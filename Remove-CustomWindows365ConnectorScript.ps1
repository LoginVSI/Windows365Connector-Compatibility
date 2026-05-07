<#
.SYNOPSIS
    Removes the custom Windows 365 Connector script from Login Enterprise.

.DESCRIPTION
    Deletes the custom Windows App connector script from a Login Enterprise appliance.
    After the custom script is deleted, Login Enterprise returns to its built-in/default
    connector script matching behavior.

    PowerShell 5.1 compatible.
    No admin rights required.
    Requires a Login Enterprise Public API token with appropriate access.

.EXAMPLE
    .\Remove-CustomWindows365ConnectorScript.ps1 `
        -ApplianceUrl "https://my-login-enterprise-appliance"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ApplianceUrl,

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

public class TrustAllCertsPolicyForLERemove : ICertificatePolicy {
    public bool CheckValidationResult(
        ServicePoint srvPoint,
        X509Certificate certificate,
        WebRequest request,
        int certificateProblem) {
        return true;
    }
}
"@

    [System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicyForLERemove
}

try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    if ($IgnoreCertificateErrors) {
        Write-Warning "Ignoring SSL/TLS certificate validation errors for this PowerShell process."
        Write-Warning "Use this only for lab, test, or troubleshooting scenarios."
        Enable-IgnoreCertificateErrors
    }

    $secureToken = Read-Host "Enter Login Enterprise API token" -AsSecureString
    $apiToken = ConvertFrom-SecureStringToPlainText -SecureString $secureToken

    if ([string]::IsNullOrWhiteSpace($apiToken)) {
        throw "API token is empty."
    }

    $baseUrl = $ApplianceUrl.TrimEnd("/")
    $uri = "$baseUrl/publicApi/$ApiVersion/connector-scripts/custom?target=windowsApp"

    Write-Host "Deleting custom Windows App connector script from:"
    Write-Host $uri
    Write-Host ""

    Invoke-RestMethod `
        -Method Delete `
        -Uri $uri `
        -Headers @{
            Authorization = "Bearer $apiToken"
            Accept        = "application/json"
        } `
        -TimeoutSec 60

    Write-Host "SUCCESS: Custom connector script deleted."
    Write-Host "Login Enterprise should now use the built-in/default connector script matching behavior."
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