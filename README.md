# Windows 365 Connector Compatibility

This repository provides compatibility versions of the Windows 365 Connector workload script for Login Enterprise when the Windows App UI changes.

Microsoft may introduce Windows App updates before the next Login Enterprise release cycle. If a Windows App UI change affects your Windows 365 tests, this repository can provide a validated replacement connector script.

For the official Login Enterprise Windows 365 Connector documentation, see:

https://docs.loginvsi.com/login-enterprise/6.6/windows-365-preview

## Validated Windows App Version

For example, the script in the folder:

```text
WindowsApp-2.0.964.0
```

was validated against:

```text
Windows App version 2.0.964.0
```

If Microsoft changes the Windows App UI again, a new folder can be added with a validated script for that version.

## Script Metadata Header

Connector scripts used with Login Enterprise 6.6 and later must include metadata comments at the top of the script file.

Example:

```csharp
// Version: Custom-2.0.964.0
// Target: Windows App
// Target version: 2.0.964.0
```

The `Target version` value should match the Windows App version the script was validated against when possible.

## Login Enterprise 6.6 and Later: Upload the Script by API

Login Enterprise 6.6 and later can manage Windows 365 Connector scripts centrally through the Public API. This removes the need to replace the connector script on each Launcher machine.

Use `Upload-Windows365ConnectorScript.ps1` to upload a custom Windows 365 Connector script to your Login Enterprise appliance.

### Requirements

- PowerShell 5.1 or later
- A Login Enterprise appliance running 6.6 or later
- Windows-based Launchers running 6.6 or later
- A Login Enterprise Public API token with appropriate access
- A `Windows365ConnectorScript.cs` file with the required metadata header

No administrator rights are required to run the upload script.

### Upload Example

From the root of this repository:

```powershell
.\Upload-Windows365ConnectorScript.ps1 `
    -ApplianceUrl "https://your-login-enterprise-appliance" `
    -ScriptPath ".\WindowsApp-2.0.964.0\Windows365ConnectorScript.cs"
```

When prompted, paste your Login Enterprise API token.

### Successful Upload

A successful upload returns output similar to:

```text
SUCCESS: Upload complete.

Version:        Custom-2.0.964.0
Target:         windowsApp
Target version: 2.0.964.0
Is custom:      True
Created time:   2026-05-07T01:01:53.4486544Z
```

The important values are:

- `Target: windowsApp`
- `Is custom: True`
- The expected `Version`
- The expected `Target version`

After a custom script is uploaded, Login Enterprise uses that custom script for the Windows App target instead of automatic built-in script matching.

## Return to Built-In Script Matching

If you uploaded a custom script and want Login Enterprise to return to its built-in/default script matching behavior, run:

```powershell
.\Remove-CustomWindows365ConnectorScript.ps1 `
    -ApplianceUrl "https://your-login-enterprise-appliance"
```

When prompted, paste your Login Enterprise API token.

A successful removal returns output similar to:

```text
SUCCESS: Custom connector script deleted.
Login Enterprise should now use the built-in/default connector script matching behavior.
```

## Certificate Trust

Start by running the upload script normally.

If the machine running PowerShell trusts the Login Enterprise appliance certificate, no additional certificate steps are needed.

If you receive an SSL/TLS trust error, such as:

```text
The underlying connection was closed: Could not establish trust relationship for the SSL/TLS secure channel.
```

then the machine running PowerShell does not trust the certificate presented by the Login Enterprise appliance.

For production environments, use a certificate trusted by your organization, or make sure the issuing root CA and any intermediate CA certificates are trusted by the machine running PowerShell.

For Login Enterprise appliance certificate management, see:

https://docs.loginvsi.com/login-enterprise/6.6/managing-certificates-on-the-appliance

For Microsoft guidance on Windows certificate stores and Trusted Root Certification Authorities, see:

https://learn.microsoft.com/windows-hardware/drivers/install/trusted-root-certification-authorities-certificate-store

### Lab or Troubleshooting Use Only

For lab, test, or troubleshooting scenarios, the upload and removal scripts also support:

```powershell
-IgnoreCertificateErrors
```

Upload example:

```powershell
.\Upload-Windows365ConnectorScript.ps1 `
    -ApplianceUrl "https://your-login-enterprise-appliance" `
    -ScriptPath ".\WindowsApp-2.0.964.0\Windows365ConnectorScript.cs" `
    -IgnoreCertificateErrors
```

Remove example:

```powershell
.\Remove-CustomWindows365ConnectorScript.ps1 `
    -ApplianceUrl "https://your-login-enterprise-appliance" `
    -IgnoreCertificateErrors
```

Use `-IgnoreCertificateErrors` at your own discretion. It bypasses certificate validation for the PowerShell process and is not recommended for production use.

## Manual Replacement

For Login Enterprise 6.5, older environments, or cases where API-based script management is not available, use the manual replacement method below.

On the Launcher machine, locate the default connector script:

```text
C:\Program Files\Login VSI\Login Enterprise Launcher\Windows365ConnectorScript.cs
```

1. Back up the existing `Windows365ConnectorScript.cs` file.
2. Download the replacement `Windows365ConnectorScript.cs` from the appropriate Windows App version folder in this repository.
3. Copy the file into:

```text
C:\Program Files\Login VSI\Login Enterprise Launcher\
```

4. Do not rename the file. The Windows 365 Connector connection type depends on the filename and path remaining exactly:

```text
Windows365ConnectorScript.cs
```

After replacing the file, run your Login Enterprise test scenario again.

## Notes

- For Login Enterprise 6.6 and later, API-based script management is preferred because it centralizes the script instead of requiring per-Launcher file replacement.
- If a custom script is uploaded, Login Enterprise uses that custom script instead of automatic matching against built-in script versions.
- Use the removal script if you want to return to built-in/default script matching.
- Keep a backup of any existing script before replacing it manually.
