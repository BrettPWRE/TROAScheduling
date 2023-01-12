# Main script for interacting with the the TROA Schedule Lists
# Launches a GUI for the user to decide which action to perform (Integration, Review, or Copy to Test Site)
# Runs other scripts depending on the response, each of which will open their own GUI's to further specify parameters
."$PSScriptRoot\Functions\USWMSharepoint_Globals.ps1"
."$PSScriptRoot\Functions\SelectActionGUI.ps1"

Switch(Select-TROASchedulingAction){
    "Integration"{
        ."$PSScriptRoot\Functions\IntegrateScheduleItems.ps1"
    }
    "Review"{
        ."$PSScriptRoot\Functions\ReviewScheduleLists.ps1"
    }
    "Copy"{
        ."$PSScriptRoot\Functions\CopyToTestSite.ps1"
    }
    default{
        Write-Host "Exiting TROA Scheduling Application"
        Exit
    }

}