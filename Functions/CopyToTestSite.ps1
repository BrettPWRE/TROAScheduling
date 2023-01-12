# Dot source external files
."$PSScriptRoot\USWMSharepoint_Functions.ps1"
."$PSScriptRoot\CopyToTestSiteGUI.ps1"
."$PSScriptRoot\PWRELogin.ps1"

# Call GUI, get parameters for test site copy
$Response = Get-CopyTestSiteParameters

Switch($Response.Scope){
    "All" {
        # Make a list of all USWM schedule lists, by looking for lists with Drafts and Integrated views
        $ScheduleLists = Build-ListOfScheduleLists -Connection $USWM_Ctn
        $numScheduleLists = $ScheduleLists.Length
    }
    "Single"{
        $SingleList = Get-PnPList -Identity $Response.SingleList -Connection $USWM_Ctn
        $ScheduleLists = [Array]$SingleList
        $numScheduleLists = 1
    }
    default{
        # Relaunch main GUI
            .(Join-Path -Path (Split-Path -Path $PSscriptRoot -Parent) -ChildPath "TROAScheduling_Main.ps1")
        Exit
    }
}


foreach($USWM_List in $ScheduleLists){

    #--------- Specify List to Copy ----------------#
    $USWM_ListName = $USWM_List.Title
    Write-Host '***** Starting List ' $USWM_ListName ' *****'
    Write-Host '----- Reading list data from USWM -----'
    $USWM_Fields = Get-PnPField -List $USWM_ListName -Connection $USWM_Ctn
    $USWM_Views = Get-PnPView -List $USWM_ListName -Connection $USWM_Ctn
    $USWM_Drafts = $USWM_Views[$USWM_Views.Title.IndexOf('Drafts')]
    $USWM_Integrated = $USWM_Views[$USWM_Views.Title.IndexOf('Integrated')]
    Write-Host '----- Reading list items -----'
    $USWM_Items = Get-PnPListItem -List $USWM_ListName -Connection $USWM_Ctn | Where {$_['Modified'] -gt (Get-Date).AddMonths(-3)} # Only get items that have been modified in the last three months

    # Add List to PWRE site
    Write-Host '----- Adding list to PWRE -----'


    # If list already exists on PWRE site, check with the user if they want to delete a list in "Single" mode
    # Don't check with user if in "All" mode, assume that they want to overwrite the list
    if((Get-PnPList -Identity $USWM_ListName -Connection $PWRE_Ctn).Title -eq $USWM_ListName){
        if($Response.Scope -eq "Single"){
            Switch(Read-Host 'List already exists on PWRE site: ' $USWM_ListName ' Delete? (y/n)'){
                'y'{
                    Remove-PnPList -Identity $USWM_ListName -Connection $PWRE_Ctn -Force
                }
                default{
                    Write-Host 'No action taken'
                    Exit
                }
            }

        }else{
            Remove-PnPList -Identity $USWM_ListName -Connection $PWRE_Ctn -Force
        }

    }

    $PWRE_List = New-PnPList -Title $USWM_ListName -Template GenericList -OnQuickLaunch -Connection $PWRE_Ctn
    $PWRE_Fields = Get-PnPField -List $USWM_ListName -Connection $PWRE_Ctn

    Write-Host '----- Adding Fields to PWRE List -----'
    $Calculated_Fields = @()
    foreach($field in $USWM_Fields){
        if($field.Title -notin $PWRE_Fields.Title -and $field.InternalName -notin $PWRE_Fields.InternalName){
            # Do calculated columns at the end, so that their required columns are already created
            if($field.TypeAsString -eq 'Calculated'){
                $Calculated_Fields += $field
            # Handle managed metadata fields separately
            # Add Single line text fields for managed metadata fields
            }elseif($field.TypeAsString -eq 'TaxonomyFieldType'){ 
                $NewField = Add-PnPField -List $USWM_ListName -Type Text -InternalName $field.InternalName -DisplayName $field.Title -Connection $PWRE_Ctn
            }else{
                $NewField = Add-PnPFieldFromXml -List $USWM_ListName -FieldXml $field.SchemaXml -Connection $PWRE_Ctn
            }
        
        }
     }
    # Add calculated fields
     foreach($cfield in $Calculated_Fields){
        if($cfield.Title -eq 'Changed'){
            $Formula = '=IF([Submitted Status]="Base",IF([Base Create Date]&lt;[Modified],"Changed","Base"),"New")'
            $NewField = Add-PnPField -List $USWM_ListName -Type Calculated -InternalName $cfield.InternalName -DisplayName $cfield.Title -Formula $Formula -Connection $PWRE_Ctn
        }else{
            # Copy and pasting (or using the xml) of the existing calculated columns doesn't work right now, because of syntax issues with < and > operators.
            # Currently manually handling the "Changed" formula, and assuming it's always the same
            Write-Host 'Not prepared to handle calculated columns other than "Changed"'
            Write-Host $cfield.Title
        }

     }


     # Add Views to PWRE list
     Write-Host '----- Adding Views to PWRE List -----'
     $Draft_Fields = @()
     $Integrated_Fields = @()

     foreach($field in $USWM_Drafts.ViewFields){
        $Draft_Fields += $field.ToString()
     }
      foreach($field in $USWM_Integrated.ViewFields){
        $Integrated_Fields += $field.ToString()
     }

     $PWRE_Drafts = Add-PnPView -List $PWRE_List.Title -Title $USWM_Drafts.Title -Fields $Draft_Fields -Query $USWM_Drafts.ViewQuery -SetAsDefault -Connection $PWRE_Ctn
     Set-PnPView -Identity $PWRE_Drafts -Values @{CustomFormatter = $USWM_Drafts.CustomFormatter} -Connection $PWRE_Ctn | Out-Null
     $PWRE_Integrated =  Add-PnPView -List $PWRE_List.Title -Title $USWM_Integrated.Title -Fields $Integrated_Fields -Query $USWM_Integrated.ViewQuery -Connection $PWRE_Ctn
     Set-PnPView -Identity $PWRE_Integrated -Values @{CustomFormatter = $USWM_Integrated.CustomFormatter} -Connection $PWRE_Ctn | Out-Null
 
     #Remove-PnPView -List $PWRE_List -Identity "All Items" -Force | Out-Null


     # Copy Items to PWRE list, function at top of script
    if($USWM_Items.Length -gt 0){
        Write-Host "----- Copying List Items -----"
         foreach($item in $USWM_Items){
            Copy-USWMItemToPWRE -SourceItem $item -DestinationList $PWRE_List -SourceFieldsAll $USWM_Fields
            $ProgressString = 'Item ' + $USWM_Items.IndexOf($item) + ' of ' + ($USWM_Items.Length-1)
            $PctComplete = [math]::Round($USWM_Items.IndexOf($item)/($USWM_Items.Length-1),3) * 100
            Write-Progress -Activity 'Copy Items' -Status $ProgressString -PercentComplete $PctComplete
        
         }
         Write-Progress -Activity 'Copy Items' -Status 'Ready' -Completed
    }else{
    Write-Host '----- No Items to Copy -----'
    }

     Write-Host '***** Finished List ' $USWM_ListName ' *****'
     Write-Host '***** Finished ' ($ScheduleLists.IndexOf($USWM_List) + 1) ' of '  $numScheduleLists ' *****'

 }

 # Relaunch main GUI
 .(Join-Path -Path (Split-Path -Path $PSscriptRoot -Parent) -ChildPath "TROAScheduling_Main.ps1")