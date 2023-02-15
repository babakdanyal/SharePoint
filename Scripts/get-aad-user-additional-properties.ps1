#-----------------------------------------------------------------------------------------------------------
# Description: This script retrieves the additional properties of the azure ad user  
#
# Pre-requisites:
#
# Install-Module SharePointPnPPowerShellOnline
# 
# Set-ExecutionPolicy RemoteSigned# 
#
# Reference: https://pnp.github.io/powershell/cmdlets/Get-PnPAzureADUser.html
#
# Achive the same results using the following Graph API end point
# https://graph.microsoft.com/v1.0/me?$select=displayName,givenName,EmployeeID,identities
#-----------------------------------------------------------------------------------------------------------
Param (    
[string]$tenantUrl = 'https://massbow.sharepoint.com'
)

#region Authenticate using interactive web login
#------------ Main ------------------ 
cls
try {
    $connection = Connect-PnPOnline -Url $tenantUrl -Interactive
    if (-not (Get-PnPContext)) {
        Write-Host "Error connecting to SharePoint Online, unable to establish context" -foregroundcolor black -backgroundcolor Red
        return
    }
    else
    {       
        #region Getting AD User

            # getting the user's additional property i.e. EmployeeId
            Get-PnPAzureADUser -Identity "alexw@massbowf.com" -Select "EmployeeId"

        #endregion Getting AD User
    }
} catch {
    Write-Host "Unable to executing commands: $_.Exception.Message" -foregroundcolor black -backgroundcolor Red
    return
}
#endregion Authenticate using interactive web login