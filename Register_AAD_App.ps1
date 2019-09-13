# -------------------------------------------------------------------------------
# This script is meant to deploy self-signed certificate to Azure AD
# When running it, you will be asked to login to Office 365 then to Azure.
# This script generates X509 certificates and automatically applies them.
# Refer to comments below to see the entire list of things that the script is doing.
#
# -------------------------------------------------------------------------------
$ErrorActionPreference = "Stop"
# -------------------------------------------------------------------------------
# CHECK IF WE ARE RUNNING THE SCRIPT AS AN ADMIN:
# -------------------------------------------------------------------------------
$IsRunningAsAdmin = [bool](([System.Security.Principal.WindowsIdentity]::GetCurrent()).groups -match "S-1-5-32-544")
if ($IsRunningAsAdmin -eq $false) {
    Write-Host "[ERROR] MAKE SURE YOU RUN THIS SCRIPT AS AN ADMINISTRATOR" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    return
}

# -------------------------------------------------------------------------------
# Depending on how we run the script, we must find location of the script in different ways.
# We need to set execution location to the current folder before we can use relative paths.
# This gives as an ability to just double-click a .bat file without having to CD to a certain folder on disk.
# -------------------------------------------------------------------------------
if ($psISE) { $scriptLocation = Split-Path -Path $psISE.CurrentFile.FullPath }
else { $scriptLocation = $PSScriptRoot }
Set-Location $scriptLocation
Write-Host $PSScriptRoot

try {

    $appName = "DeploymentApp"
    # -------------------------------------------------------------------------------
    # We name certificate the same name as the AAD Application
    # -------------------------------------------------------------------------------
    $certificateName = $appName

    # -------------------------------------------------------------------------------
    # Generate strong passwords for the certificate
    # -------------------------------------------------------------------------------
    Add-Type -AssemblyName System.Web 
    # "&" is converted to "\u0026" when using ConvertTo-Json, so we strip this off.
    $certificatePassword = [System.Web.Security.Membership]::GeneratePassword(12, 3) -replace "&","!"
    $encryptedCertPassword = ConvertTo-SecureString -String $certificatePassword -AsPlainText -Force

    # -------------------------------------------------------------------------------
    # Generate self-signed .cer and .pfx files
    # -------------------------------------------------------------------------------
    New-PnPAzureCertificate -OutPfx "$certificateName.pfx" -OutCert "$certificateName.cer" -CertificatePassword $encryptedCertPassword
    Write-Host "[Success] Generated a self-signed certificate: $certificateName" -ForegroundColor Green

    # -------------------------------------------------------------------------------
    # Logging to Azure Cli using Office 365 Global Tenant Admin
    # -------------------------------------------------------------------------------
    Write-Host "[Pending user action] Enter Global Office 365 Admin credentials..." -ForegroundColor Yellow
    az login --allow-no-subscriptions

    Write-Host "[Success] Connected Office 365's Azure tenant" -ForegroundColor Green

    # -------------------------------------------------------------------------------
    # Select Azure Subscription Associated with Office 365 tenant
    # -------------------------------------------------------------------------------
    $subscription = az account show | ConvertFrom-Json
    $account = az account show | ConvertFrom-Json

    # -------------------------------------------------------------------------------
    # Delete the app if it already exists
    # -------------------------------------------------------------------------------
    $apps = az ad app list | ConvertFrom-Json
    $app = $apps | Where-Object { $_.displayName -eq $appName }
    if ($null -ne $app) {
        az ad app delete --id $app.appId
        Write-Host "[Success] Deleted existing AAD Application with the name $appName" -ForegroundColor Green
    }

    # -------------------------------------------------------------------------------
    # Find existing application by name
    # -------------------------------------------------------------------------------
    $apps = az ad app list | ConvertFrom-Json
    $app = $apps | Where-Object { $_.displayName -eq $appName }

    # -------------------------------------------------------------------------------
    # Assign Permissions the app. Permissions are described in the json file.
    # We are only asking for one permission: SharePoint Online: Read all site collections
    # -------------------------------------------------------------------------------
    if ($null -eq $app) {
        $app = az ad app create `
            --display-name $appName `
            --identifier-uris https://$appName.websites.net `
            --required-resource-accesses "requiredResourceManifest.json" `
        | ConvertFrom-Json

        Write-Host "[Success] Registered a new AAD Application with the name $appName" -ForegroundColor Green
    }

    # -------------------------------------------------------------------------------
    # Grant permissions (Admin-consent) programmatically. This commandlet became first available April 04 2019
    # -------------------------------------------------------------------------------
    az ad app permission admin-consent --id $app.appId

    Write-Host "[Success] $appName was granted permissions to access SharePoint Online Site collections for read-only" -ForegroundColor Green

    # -------------------------------------------------------------------------------
    # Upload a .cer file to the AAD application
    # -------------------------------------------------------------------------------
    $cerName = $certificateName + ".cer"
    az ad app credential reset --id $app.appId --cert `@$cerName 

    Write-Host "[Success] Uploaded self-signed certificate to the $appName application" -ForegroundColor Green

    $fileContentBytes = Get-Content "$certificateName.pfx" -Encoding Byte
    $pfxBlob = [System.Convert]::ToBase64String($fileContentBytes)

    $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
    $cert.Import((Get-Location).Path + "\" + $cerName)
    $thumbprint = $cert.Thumbprint

    # -------------------------------------------------------------------------------
    # Preparing an object that will contain all useful information about our Azure AD App
    # -------------------------------------------------------------------------------
    $AppDetails = @{
        thumbprint          = $thumbprint
        pfxBlob             = $pfxBlob
        appId               = $app.appId
        tenantId            = $account.tenantId
        certificatePassword = $certificatePassword
        certificateName     = $certificateName
        appName             = $appName
    }
    # -------------------------------------------------------------------------------
    # Save information about the O365 Application in a file
    # -------------------------------------------------------------------------------
    $AppDetails | ConvertTo-Json | Out-File "AppDetails.json"

    Write-Host "[Success] Saved AppDetails.json file about the '$appName' application on disk" -ForegroundColor Green

    # -------------------------------------------------------------------------------    
    # We no longer need to work with Azure associated with Office 365
    # -------------------------------------------------------------------------------
    az logout
    Write-Host "[Success] Logged out from Office 365's Azure tenant" -ForegroundColor Green
    
}
catch {
    $Error
    Write-Host "An error occurred. If it was not caused by you cancelling the deployment process - please capture the logs displayed above and contact the deployment specialist." -ForegroundColor Cyan
    Read-Host "Press Enter to exit"
    return
}

Write-Host "[Success] Application has been deployed." -ForegroundColor Cyan
Read-Host "Press Enter to close this window and exit"