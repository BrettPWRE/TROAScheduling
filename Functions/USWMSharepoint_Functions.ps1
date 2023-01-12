Function Copy-USWMItemToPWRE(){
# Function takes an item from the USWM list, and writes the necessary fields to the PWRE copy

param(
    [Parameter(Mandatory=$true)] [Microsoft.SharePoint.Client.SecurableObject] $SourceItem,
    [Parameter(Mandatory=$true)] [Microsoft.SharePoint.Client.SecurableObject] $DestinationList,
    [Parameter(Mandatory=$true)] [System.Array] $SourceFieldsAll
    )
    # Only look at fields that aren't read only, and some other requirements
    $SourceListFields = $SourceFieldsAll |  Where { (-Not ($_.ReadOnlyField)) -and (-Not ($_.Hidden)) -and ($_.InternalName -ne  "ContentType") -and ($_.InternalName -ne  "Attachments") }

    $ItemValue = @{}
    foreach($SourceField in $SourceListFields){
       if($SourceItem[$SourceField.InternalName] -ne $null){
            $FieldType = $SourceField.TypeAsString
            if($FieldType -eq "User" -or $FieldType -eq "UserMulti" -or $FieldType -eq "Lookup" -or $FieldType -eq "LookupMulti"){
                #Skip users for now
            }elseif($FieldType -eq "URL"){
                #Skip URL's for now
            }elseif($FieldType -eq "TaxonomyFieldType" -or $FieldType -eq "TaxonomyFieldTypeMulti"){
                #Add Managed Metadata labels as text
                $ItemValue.add($SourceField.InternalName,$SourceItem[$SourceField.InternalName].Label)
            }else{
                $ItemValue.add($SourceField.InternalName,$SourceItem[$SourceField.InternalName])
            }
       }
    }
    $NewItem = Add-PnPListItem -List $DestinationList -Values $ItemValue -Connection $PWRE_Ctn

    # Edit Modified Date
    $FinalItem = Set-PnPListITem -List $DestinationList -Identity $NewItem -Values @{"Modified" = $SourceItem['Modified']} -Connection $PWRE_Ctn #Weird bug that causes this line to fail occasionally, it seems like assigning output to a variable solves the issue


}

Function Integrate-DraftItem(){
# Function takes a draft item from the schedule list, and creates a copy with the necessary changes to have it show up in the integrated view
# It then updates the original draft item to persist in the draft view

param(
    [Parameter(Mandatory=$true)] [Microsoft.SharePoint.Client.SecurableObject] $SourceItem,
    [Parameter(Mandatory=$true)] [Microsoft.SharePoint.Client.SecurableObject] $SourceList,
    [Parameter(Mandatory=$true)] [System.Array] $SourceFieldsAll,
    [Parameter(Mandatory=$true)] [System.DateTime] $SubDate,
    [Parameter(Mandatory=$true)] [System.Object] $Connection

    )
    # Only look at fields that aren't read only, and some other requirements
    $SourceListFields = $SourceFieldsAll |  Where { (-Not ($_.ReadOnlyField)) -and (-Not ($_.Hidden)) -and ($_.InternalName -ne  "ContentType") -and ($_.InternalName -ne  "Attachments") }

    $ItemValue = @{}
    $MMD_Fields = @()
    foreach($SourceField in $SourceListFields){
       if($SourceItem[$SourceField.InternalName] -ne $null){
            $FieldType = $SourceField.TypeAsString
            if($FieldType -eq "User" -or $FieldType -eq "UserMulti" -or $FieldType -eq "Lookup" -or $FieldType -eq "LookupMulti"){
                #Skip users for now
            }elseif($FieldType -eq "URL"){
                #Skip URL's for now
            }elseif($FieldType -eq "TaxonomyFieldType" -or $FieldType -eq "TaxonomyFieldTypeMulti"){
                #Build list of Managed Metadata Fields. Have to set these after creating the new item
                $MMD_Fields += $SourceField
            }else{
                $ItemValue.add($SourceField.InternalName,$SourceItem[$SourceField.InternalName])
            }
       }
    }
    # Update certain fields to make the new item show up in the integrated view. This is designed to match the operation of the original flows
    $ItemValue['SubmittedDate'] = $SubDate
    $ItemValue['SubmittedYear'] = $SubDate.Year
    $ItemValue['SubmittedMonth'] = $SubmittedDate.ToString("MM") + " - " + $SubmittedDate.ToString("MMM")
    $ItemValue['Change_Status'] = $SourceItem['Changed']
    $ItemValue['SubmittedStatus'] = 'Integrated' 

    $NewItem = Add-PnPListItem -List $SourceList -Values $ItemValue -Connection $Connection
    foreach($MMD_Field in $MMD_Fields){
        Set-PnPTaxonomyFieldValue -ListItem $NewItem -InternalFieldName $MMD_Field.InternalName -TermID $SourceItem[$MMD_Field.InternalName].TermGuid -Connection $Connection
    }
    

    # Edit source item to persist in draft view
    $UpdatedDraftValues = @{
    "BaseCreateDate" = (Get-Date).AddSeconds(60);
    "SubmittedDate" = $null;
    "SubmittedMonth" = $null;
    "SubmittedYear" = $null;
    "Party_x0020_Status" = 'Submitted'
    "SubmittedStatus" = "Base";
    "Change_Status" = "Base"
    }
    $UpdatedDraftItem = Set-PnPListITem -List $SourceList -Identity $SourceItem -Values $UpdatedDraftValues -Connection $Connection


}

Function Build-ListOfScheduleLists(){
# Function takes a connection (USWM or PWRE) and returns a list of all the schedule lists
    param(
    [Parameter(Mandatory=$true)] [System.Object] $Connection
    )
    Write-Host '----- Building list of Schedule Lists -----'
    $AllLists = Get-PnPList -Connection $Connection
    $ScheduleLists = @()
    foreach($list in $AllLists){
        $Views = Get-PnPView -List $list -Connection $Connection
        if(-not $list.Title.Contains("Test") -and "Drafts" -in $Views.Title -and "Integrated" -in $Views.Title){
            $ScheduleLists += $list
        }
    }
    Return $ScheduleLists
}

Function Select-ItemsToEdit(){
    # Function outputs a grid window for user to select one or multiple draft items to integrate.
    # Can be used in the case that items were added or modified 
    # shortly after an integration, or that items in the draft view weren't marked as submitted when they were intended to be
    # The window it produces lists all draft view items with fields in order of the list view, plus 'Modified' which could be useful to determine which items need to integrate
    param(
    [Parameter(Mandatory=$true)] [System.Object] $Mode,
    [Parameter(Mandatory=$true)] [System.Object] $ListName,
    [Parameter(Mandatory=$true)] [System.Object] $Action

    )
    
    Write-Host "----- Retrieving Draft Items -----"

    Switch($Mode){
        "Test"{
            ."$PSScriptRoot\PWRELogin.ps1"
            $Ctn = $PWRE_Ctn
        }
        "USWM"{
            $Ctn = $USWM_Ctn
        }
        default{
            "No action taken"
            Exit
        }
    }

    $List = Get-PnPList -Identity $ListName -Connection $Ctn
    $DraftItems = Get-PnPListItem -List $List -Connection $Ctn | Where {$_['SubmittedStatus'] -ne "Integrated" -and $_['Title'] -eq $null -and $_['Party_x0020_Status'] -eq "Submitted"}
    $DraftView = Get-PnPView -List $List -Identity "Drafts" -Connection $Ctn
    $DraftFields = $DraftView.ViewFields 

    $FieldDetails = @{'InternalName'=@();'DisplayName'=@();'Type'=@()}
    foreach($field in $DraftFields){
        $FieldAllInfo = Get-PnPField -List $List -Identity $field -Connection $Ctn
        $FieldDetails['InternalName'] += $field
        $FieldDetails['DisplayName'] += $FieldAllInfo.Title
        $FieldDetails['Type'] += $FieldAllInfo.TypeAsString
    }

    $FormattedItems = @()
    foreach($item in $DraftItems){
            #$ItemValues = @{'ID'=$item['ID'];'Modified'=[TimeZoneInfo]::ConvertTimeBySystemTimeZoneId($item['Modified'],'Mountain Standard Time')}
            $ItemValues = @{'ID'=$item['ID']}
        foreach($field in $FieldDetails['InternalName']){
            $DisplayName = $FieldDetails['DisplayName'][$FieldDetails['InternalName'].IndexOf($field)]
            $FieldType = $FieldDetails['Type'][$FieldDetails['InternalName'].IndexOf($field)]
            if($field -eq 'StartDate' -or $field -eq '_EndDate'){
                $ItemValues.Add($DisplayName,$item[$field].ToString('M/d/yyyy'))
            }elseif($field -eq 'Modified'){
                $ItemValues.Add($DisplayName,[TimeZoneInfo]::ConvertTimeBySystemTimeZoneId($item[$field],'Mountain Standard Time'))
            }elseif($FieldType -eq "TaxonomyFieldType" -or $FieldType -eq "TaxonomyFieldTypeMulti"){
                $ItemValues.Add($DisplayName,$item[$field].Label)
            }else{
                $ItemValues.Add($DisplayName,($item[$field] -replace '<[^>]+>',''))
            }
        }
        $FormattedItems += $ItemValues
    }

    if($Action -eq "Integrate"){
        $GridTitle = "Select Items to Integrate From " + $ListName + " (Ctrl+Click for Multiple)"
    }elseif($Action -eq "Edit Modified"){
        $GridTitle = "Select Items to Edit Modified Date For (Ctrl+Click for Multiple)"
    }
    else{$GridTitle = ""}

    $Response = $FormattedItems |  ForEach {[PSCustomObject]$_} | Select-Object -Property (@('ID') + $FieldDetails['DisplayName']) | Out-GridView -Title $GridTitle -OutputMode Multiple
    $SelectedID = [Array]$Response.ID

    Return $SelectedID
}
