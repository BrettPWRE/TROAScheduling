$PWRE_URL = "https://precisionwater.sharepoint.com/sites/110_FederalWatermaster/110010_TROA_Implementation_2021"
# Watermaster won't be able to connect to the PWRE site, display a message but avoid raising an exception
if($PWRE_Ctn.Url -ne $PWRE_URL){
    try{
        $PWRE_Ctn = Connect-PnPOnline -Url $PWRE_URL -UseWebLogin -ReturnConnection -WarningAction Ignore
    }
    catch [Microsoft.SharePoint.Client.IdcrlException]{
        Write-Host "Unable to connect to the PWRE Test Site. Only actions on the USWM site are possible."
    }
}