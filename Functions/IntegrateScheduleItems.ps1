# Dot source external files
."$PSScriptRoot\USWMSharepoint_Functions.ps1"
."$PSScriptRoot\IntegrationGUI.ps1"


# Call GUI, get integration setup parameters
$Response = Get-IntegrationParameters

# Connect to proper site based on "Mode" Response
Switch($Response.Mode){
    "Test"{
        ."$PSScriptRoot\PWRELogin.ps1"
        $Ctn = $PWRE_Ctn
        Write-Host "----- Begin PWRE Test Site Integration -----"
    }
    "USWM"{
        $Ctn = $USWM_Ctn
        Write-Host "----- Begin USWM Official Integration -----"
    }
    "Cancel"{
        Write-Host "Cancelling Integration"
        .(Join-Path -Path (Split-Path -Path $PSscriptRoot -Parent) -ChildPath "TROAScheduling_Main.ps1")
        Exit
    }
}

# Build list of schedule lists to integrate, based on "Scope" response
Switch($Response.Scope){
    "All"{
        $ScheduleLists = Build-ListOfScheduleLists -Connection $Ctn
        $numScheduleLists = $ScheduleLists.Length
    }
    "Party"{
        $Party = $Response.Party
        Write-Host "This functionality hasn't been developed yet"
        Exit
    }
    {($_ -eq "Single") -or ($_ -eq "Specific")}{
        $SingleList = Get-PnPList -Identity $Response.SingleList -Connection $Ctn
        $ScheduleLists = [Array]$SingleList
        $numScheduleLists = 1
    }
    default{
        Write-Host "Invalid Input"
        Exit
    }
}

$SubmittedDate = Get-Date
$list_ctr = 0
foreach($list in $ScheduleLists){
    $list_name = $list.Title
    $Fields = Get-PnPField -List $list_name -Connection $Ctn
    $list_progress = [math]::Round($list_ctr/$numScheduleLists,3) * 100
    Write-Progress -Activity "Processing List: $list_name" -Status "$list_ctr of $numScheduleLists Completed" -PercentComplete $list_progress -ID 0

    # Delete items that have been integrated on the current day (this allows integration to be rerun)
    # Only do this if not running in the "Specific Item Scope", otherwise all properly integrated items would get deleted
    if($Response.Scope -ne "Specific"){
        $Items_To_Delete = Get-PnPListItem -List $list_name -Connection $Ctn | Where {$_['SubmittedStatus'] -eq "Integrated" -and $_['SubmittedDate'].Date -eq $SubmittedDate.Date}
        foreach($item in $Items_To_Delete){
            Remove-PnPListItem -List $list_name -Identity $item -Connection $Ctn -Force -Recycle
        }
    }

    # Build list of draft items to integrate. Depends on whether the user has selected the "Specific Item" scope
    if($Response.Scope -eq "Specific"){
        $Draft_Items = @()
        foreach($i in $Response.SelectedItems){
            $Draft_Items += Get-PnPListItem -List $list.Title -Connection $Ctn -Id $i
        }
    }else{
        $Draft_Items = Get-PnPListItem -List $list.Title -Connection $Ctn | Where {$_['SubmittedStatus'] -ne "Integrated" -and $_['Title'] -eq $null -and $_['Party_x0020_Status'] -eq "Submitted"} # For some reason, most of the schedule lists have one old item with a title, ignore this item
    }
    $num_Drafts = $Draft_Items.Length
    $item_ctr = 0
    foreach($item in $Draft_Items){
        $item_progress = [math]::Round($item_ctr/$num_Drafts,3) * 100
        Write-Progress -Activity "Processing Draft Items" -Status "$item_ctr of $num_Drafts Completed" -PercentComplete $item_progress -ID 1 -ParentID 0

        ############# Integrate and Copy to Base goes here #################
        Integrate-DraftItem -SourceList $list -SourceItem $item -SourceFieldsAll $Fields -SubDate $SubmittedDate -Connection $Ctn

        $item_ctr += 1
    }
    Write-Progress -Activity "Processing Draft Items" -Status "$item_ctr of $num_Drafts Completed" -PercentComplete 100 -ID 1 -ParentID 0
    $list_ctr += 1
}
Write-Progress -Activity "Processing Draft Items" -Status "$item_ctr of $num_Drafts Completed" -Completed -ID 1 -ParentID 0
Write-Progress -Activity "Processing List: $list_name" -Status "$list_ctr of $numScheduleLists Completed" -Completed -ID 0
Write-Host "----- Integration Complete -----"

 # Relaunch main GUI
 .(Join-Path -Path (Split-Path -Path $PSscriptRoot -Parent) -ChildPath "TROAScheduling_Main.ps1")

