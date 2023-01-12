# Global variables used throughout Sharepoint Scheduling scripts and functions

# Sharepoint Service Credentials
# Store credentials in a "USWM.auth" file in the same folder, with the following format:
<#
username ***@uswatermaster.org
password ***
#>

$AuthFile_list = Get-Content -Path $PSScriptRoot\USWM.auth
for($i = 0;$i -lt $AuthFile_list.Length;$i++){
    $key = $AuthFile_list[$i].split(" ")[0]
    $value = $AuthFile_list[$i].split(" ")[1]

    if($key.ToLower() -eq "username"){
        $SptSvcUser = $value
    }elseif($key.ToLower() -eq "password"){
        $SptSvcPass = ConvertTo-SecureString $value -AsPlainText -Force
    }
}
$USWM_Cred = New-Object System.Management.Automation.PSCredential ($SptSvcUser, $SptSvcPass)


# URL's and Connections
# Ignore Warning Message, because this was written with a Legacy version of PnPPowershell, don't want to get warnings every time


$USWM_URL = "https://uswatermaster.sharepoint.com/sites/schedule"
if($USWM_Ctn.Url -ne $USWM_URL){
    $USWM_Ctn = Connect-PnPOnline -Url $USWM_URL -Credentials $USWM_Cred -ReturnConnection -WarningAction Ignore
}


