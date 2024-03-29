#-----------------------------------------------------------------------------------------------------------
# Description: This script hides the left navigation or quick launch from the SharePoint Online sites
#
# 	Reference:https://pnp.github.io/powershell/cmdlets/Get-PnPWeb.html
#-----------------------------------------------------------------------------------------------------------
Param (    
[string]$siteUrl = 'https://your-company.sharepoint.com/sites/Warehouse'
#[string]$username = 'username@your-company.com',
#[string]$password = '*******'
)
 
 
#region Authenticate using hard coded credentials  
#
#[SecureString]$SecurePass = ConvertTo-SecureString $password -AsPlainText -Force
#[System.Management.Automation.PSCredential]$PSCredentials = New-Object System.Management.Automation.PSCredential($username, $SecurePass)
#Connect-PnPOnline -Url $SiteCollection -Credentials $PSCredentials
#
#endregion Authenticate using hard coded credentials 

#region Authenticate using interactive web login    
#------------ Main ------------------ 
cls
try {
    $connection = Connect-PnPOnline -Url $siteUrl -UseWebLogin
    if (-not (Get-PnPContext)) {
        Write-Host "Error connecting to SharePoint Online, unable to establish context" -foregroundcolor black -backgroundcolor Red
        return
    }
    else
    {       
        #region Hiding quick launch

            # getting the site's Web object.
            $web = Get-PnPWeb
            
            # disable the Quick Launch Menu
            # you can toggle between $false and $true to hide and show the left navigation.

            $web.QuickLaunchEnabled = $false
            $web.Update()
            Invoke-PnPQuery

        #endregion Hiding quick launch
    }
} catch {
    Write-Host "Unable to executing commands: $_.Exception.Message" -foregroundcolor black -backgroundcolor Red
    return
}
#endregion Authenticate using interactive web login

