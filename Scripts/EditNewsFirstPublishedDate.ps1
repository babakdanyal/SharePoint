#start region
#
#Use this script to update the original published date of news articles in the SharePoint site.
#
#end region
$SiteURL = "<site url goes here>"
$ListName= "Site Pages"
$Date = "2019-12-20"
$NewsID = 10
Connect-PnPOnline -Url $SiteURL -UseWebLogin

if (-not (Get-PnPContext)) {
    Write-Host "Error connecting to SharePoint Online, unable to establish context" -foregroundcolor black -backgroundcolor Red
    return
}
else{
    Set-PnPListItem -List $ListName -Identity $NewsID -Values @{"FirstPublishedDate"=$Date;} -SystemUpdate:$true
}
    