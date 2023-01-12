# Dot source external files
."$PSScriptRoot\USWMSharepoint_Functions.ps1"
."$PSScriptRoot\ReviewGUI.ps1"

Get-ReviewParameters

 # Relaunch main GUI
.(Join-Path -Path (Split-Path -Path $PSscriptRoot -Parent) -ChildPath "TROAScheduling_Main.ps1")