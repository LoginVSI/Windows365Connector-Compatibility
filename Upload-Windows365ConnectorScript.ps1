<#
.SYNOPSIS
    Uploads a custom Windows 365 Connector script to Login Enterprise.

.DESCRIPTION
    Reads a Windows365ConnectorScript.cs file and uploads it to the
    Login Enterprise Connector Script API as the custom Windows App script.

    PowerShell 5.1 compatible.
    No admin rights required.
    Requires a Login Enterprise Public API token with appropriate access.

.EXAMPLE
    .\Upload-Windows365ConnectorScript.ps1 `
        -ApplianceUrl "https://my-login-enterprise-appliance" `
        -ScriptPath ".\WindowsApp-2.0.964.0\Windows365ConnectorScript.cs"

.EXAMPLE
    .\Upload-Windows365ConnectorScript.ps1 `
        -ApplianceUrl "https://my-login-enterprise-appliance" `
        -ScriptPath ".\WindowsApp-2.0.964.0\Windows365ConnectorScript.cs" `
        -IgnoreCertificateErrors
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ApplianceUrl,

    [Parameter(Mandatory = $true)]
    [string]$ScriptPath,

    [Parameter(Mandatory = $false)]
    [string]$ApiVersion = "v8-preview",

    [Parameter(Mandatory = $false)]
    [switch]$IgnoreCertificateErrors
)

function ConvertFrom-SecureStringToPlainText {
    param(
        [Parameter(Mandatory = $true)]
        [securestring]$SecureString
    )

    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString)

    try {
        return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
    }
    finally {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
    }
}

function Enable-IgnoreCertificateErrors {
    Add-Type @"
using System.Net;
using System.Security.Cryptography.X509Certificates;

public class TrustAllCertsPolicyForLEUpload : ICertificatePolicy {
    public bool CheckValidationResult(
        ServicePoint srvPoint,
        X509Certificate certificate,
        WebRequest request,
        int certificateProblem) {
        return true;
    }
}
"@

    [System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicyForLEUpload
}

try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    if ($IgnoreCertificateErrors) {
        Write-Warning "Ignoring SSL/TLS certificate validation errors for this PowerShell process."
        Write-Warning "Use this only for lab, test, or troubleshooting scenarios."
        Enable-IgnoreCertificateErrors
    }

    Write-Host "Checking script path..."
    if (-not (Test-Path -Path $ScriptPath -PathType Leaf)) {
        throw "Script file not found: $ScriptPath"
    }

    Write-Host "Reading script..."
    $scriptContent = [System.IO.File]::ReadAllText($ScriptPath)
    Write-Host "Script read complete. Characters: $($scriptContent.Length)"

    if ([string]::IsNullOrWhiteSpace($scriptContent)) {
        throw "Script file is empty: $ScriptPath"
    }

    if ($scriptContent -notmatch "(?im)^//\s*Version:\s*.+" -or
        $scriptContent -notmatch "(?im)^//\s*Target:\s*Windows App") {
        Write-Warning "The script does not appear to include the required metadata header."
        Write-Warning "Expected header example:"
        Write-Warning "// Version: Custom-2.0.964.0"
        Write-Warning "// Target: Windows App"
        Write-Warning "// Target version: 2.0.964.0"
        Write-Warning ""
        Write-Warning "The upload may fail if Login Enterprise cannot validate the script metadata."
    }

    $secureToken = Read-Host "Enter Login Enterprise API token" -AsSecureString
    $apiToken = ConvertFrom-SecureStringToPlainText -SecureString $secureToken

    if ([string]::IsNullOrWhiteSpace($apiToken)) {
        throw "API token is empty."
    }

    $baseUrl = $ApplianceUrl.TrimEnd("/")
    $uri = "$baseUrl/publicApi/$ApiVersion/connector-scripts/custom?target=windowsApp"

    $body = @{
        scriptContent = $scriptContent
    } | ConvertTo-Json -Depth 10

    Write-Host "Uploading to: $uri"
    Write-Host "Script path:  $ScriptPath"
    Write-Host ""

    $response = Invoke-RestMethod `
        -Method Post `
        -Uri $uri `
        -Headers @{
            Authorization = "Bearer $apiToken"
            Accept        = "application/json"
        } `
        -ContentType "application/json" `
        -Body $body `
        -TimeoutSec 120

    Write-Host "SUCCESS: Upload complete."
    Write-Host ""
    Write-Host "Version:        $($response.version)"
    Write-Host "Target:         $($response.target)"
    Write-Host "Target version: $($response.targetVersion)"
    Write-Host "Is custom:      $($response.isCustom)"
    Write-Host "Created time:   $($response.createdTime)"
}
catch {
    Write-Host "FAILED"
    Write-Host $_.Exception.Message

    if ($_.Exception.Response -ne $null) {
        try {
            $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
            $responseBody = $reader.ReadToEnd()

            if (-not [string]::IsNullOrWhiteSpace($responseBody)) {
                Write-Host ""
                Write-Host "Response body:"
                Write-Host $responseBody
            }
        }
        catch {
            Write-Host "Could not read response body."
        }
    }

    exit 1
}
finally {
    if ($apiToken) {
        Remove-Variable apiToken -ErrorAction SilentlyContinue
    }
}