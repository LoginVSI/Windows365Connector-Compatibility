<#
.SYNOPSIS
    Retrieves Windows 365 Connector scripts from Login Enterprise.

.DESCRIPTION
    Lists the built-in Windows App connector scripts available in Login
    Enterprise and downloads either a requested target version or the newest
    available version.

    The script can also back up the currently configured custom connector script.

    PowerShell 5.1 compatible.
    No administrator rights required.

.EXAMPLE
    .\Get-Windows365ConnectorScripts.ps1 `
        -ApplianceUrl "https://your-login-enterprise-appliance"

.EXAMPLE
    .\Get-Windows365ConnectorScripts.ps1 `
        -ApplianceUrl "https://your-login-enterprise-appliance" `
        -TargetVersion "2.0.1129.0" `
        -IncludeCustomBackup

.EXAMPLE
    .\Get-Windows365ConnectorScripts.ps1 `
        -ApplianceUrl "https://your-login-enterprise-appliance" `
        -IncludeCustomBackup `
        -IgnoreCertificateErrors
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ApplianceUrl,

    [Parameter(Mandatory = $false)]
    [string]$TargetVersion,

    [Parameter(Mandatory = $false)]
    [string]$OutputDirectory = ".",

    [Parameter(Mandatory = $false)]
    [string]$ApiVersion = "v8-preview",

    [Parameter(Mandatory = $false)]
    [switch]$IncludeCustomBackup,

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
        [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
    }
    finally {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
    }
}

function Enable-IgnoreCertificateErrors {
    if (-not ("TrustAllCertsPolicyForLEConnectorRetrieval" -as [type])) {
        Add-Type @"
using System.Net;
using System.Security.Cryptography.X509Certificates;

public class TrustAllCertsPolicyForLEConnectorRetrieval : ICertificatePolicy
{
    public bool CheckValidationResult(
        ServicePoint srvPoint,
        X509Certificate certificate,
        WebRequest request,
        int certificateProblem)
    {
        return true;
    }
}
"@
    }

    [System.Net.ServicePointManager]::CertificatePolicy =
        New-Object TrustAllCertsPolicyForLEConnectorRetrieval
}

function Get-HttpStatusCode {
    param(
        [Parameter(Mandatory = $true)]
        $ErrorRecord
    )

    if ($ErrorRecord.Exception.Response -and
        $ErrorRecord.Exception.Response.StatusCode) {
        return [int]$ErrorRecord.Exception.Response.StatusCode
    }

    return $null
}

$originalCertificatePolicy =
    [System.Net.ServicePointManager]::CertificatePolicy

try {
    [Net.ServicePointManager]::SecurityProtocol =
        [Net.SecurityProtocolType]::Tls12

    if ($IgnoreCertificateErrors) {
        Write-Warning "Ignoring SSL/TLS certificate validation errors."
        Write-Warning "Use this only in lab, test, or troubleshooting environments."
        Enable-IgnoreCertificateErrors
    }

    if (-not (Test-Path -Path $OutputDirectory)) {
        New-Item `
            -Path $OutputDirectory `
            -ItemType Directory `
            -Force | Out-Null
    }

    $secureToken = Read-Host `
        "Enter Login Enterprise API token" `
        -AsSecureString

    $apiToken = ConvertFrom-SecureStringToPlainText `
        -SecureString $secureToken

    if ([string]::IsNullOrWhiteSpace($apiToken)) {
        throw "API token is empty."
    }

    $baseUrl = $ApplianceUrl.TrimEnd("/")
    $apiBase = "$baseUrl/publicApi/$ApiVersion"
    $target  = "windowsApp"

    $headers = @{
        Authorization = "Bearer $apiToken"
        Accept        = "application/json"
    }

    Write-Host "`nRetrieving built-in Windows App connector scripts..."

    $defaultScripts = Invoke-RestMethod `
        -Method Get `
        -Uri "$apiBase/connector-scripts/default?target=$target" `
        -Headers $headers `
        -TimeoutSec 60 `
        -ErrorAction Stop

    if (-not $defaultScripts) {
        throw "No built-in connector scripts were returned."
    }

    Write-Host "`nAvailable built-in connector scripts:`n"

    $defaultScripts |
        Select-Object version, targetVersion, isCustom, createdTime |
        Format-Table -AutoSize

    if ($TargetVersion) {
        $selectedScript = $defaultScripts |
            Where-Object { $_.targetVersion -eq $TargetVersion } |
            Select-Object -First 1

        if (-not $selectedScript) {
            throw "Target version '$TargetVersion' was not found."
        }
    }
    else {
        $selectedScript = $defaultScripts |
            Sort-Object {
                try {
                    [version]$_.targetVersion
                }
                catch {
                    [version]"0.0"
                }
            } -Descending |
            Select-Object -First 1
    }

    $selectedTargetVersion = $selectedScript.targetVersion
    $encodedTargetVersion =
        [Uri]::EscapeDataString($selectedTargetVersion)

    Write-Host "Downloading built-in connector script for Windows App $selectedTargetVersion..."

    $defaultContent = Invoke-RestMethod `
        -Method Get `
        -Uri "$apiBase/connector-scripts/default/content?target=$target&targetVersion=$encodedTargetVersion" `
        -Headers $headers `
        -TimeoutSec 60 `
        -ErrorAction Stop

    $defaultPath = Join-Path `
        $OutputDirectory `
        "Windows365ConnectorScript-Default-$selectedTargetVersion.cs"

    [System.IO.File]::WriteAllText(
        (Resolve-Path $OutputDirectory).Path +
            "\Windows365ConnectorScript-Default-$selectedTargetVersion.cs",
        [string]$defaultContent,
        [System.Text.Encoding]::UTF8
    )

    Write-Host "SUCCESS: Downloaded built-in script."
    Write-Host "Path: $defaultPath"

    if ($IncludeCustomBackup) {
        Write-Host "`nChecking for a custom connector script..."

        try {
            $customScript = Invoke-RestMethod `
                -Method Get `
                -Uri "$apiBase/connector-scripts/custom?target=$target" `
                -Headers $headers `
                -TimeoutSec 60 `
                -ErrorAction Stop

            Write-Host "`nCurrent custom connector script:`n"

            $customScript |
                Select-Object version, targetVersion, isCustom, createdTime |
                Format-List

            $customContent = Invoke-RestMethod `
                -Method Get `
                -Uri "$apiBase/connector-scripts/custom/content?target=$target" `
                -Headers $headers `
                -TimeoutSec 60 `
                -ErrorAction Stop

            $customPath = Join-Path `
                $OutputDirectory `
                "Windows365ConnectorScript-Custom-Backup.cs"

            [System.IO.File]::WriteAllText(
                (Resolve-Path $OutputDirectory).Path +
                    "\Windows365ConnectorScript-Custom-Backup.cs",
                [string]$customContent,
                [System.Text.Encoding]::UTF8
            )

            Write-Host "SUCCESS: Backed up custom connector script."
            Write-Host "Path: $customPath"
        }
        catch {
            $statusCode = Get-HttpStatusCode -ErrorRecord $_

            if ($statusCode -eq 404) {
                Write-Host "No custom connector script is currently configured."
            }
            else {
                throw
            }
        }
    }
}
catch {
    Write-Host "`nFAILED"
    Write-Host $_.Exception.Message

    if ($_.Exception.Response) {
        try {
            $reader = New-Object System.IO.StreamReader(
                $_.Exception.Response.GetResponseStream()
            )

            $responseBody = $reader.ReadToEnd()

            if (-not [string]::IsNullOrWhiteSpace($responseBody)) {
                Write-Host "`nResponse body:"
                Write-Host $responseBody
            }
        }
        catch {
            Write-Host "Could not read the API response body."
        }
    }

    exit 1
}
finally {
    [System.Net.ServicePointManager]::CertificatePolicy =
        $originalCertificatePolicy

    if ($apiToken) {
        Remove-Variable apiToken -ErrorAction SilentlyContinue
    }
}