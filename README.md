# Windows 365 Connector Compatibility

This repository provides compatibility versions of the Windows 365 connector workload script for Login Enterprise when the Windows App UI changes.

Login Enterprise 6.5 includes an updated Windows 365 connector. However, Microsoft may introduce Windows App updates before the next Login Enterprise release cycle. This repository allows you to replace the connector workload script if a UI change affects your tests.

## Validated Windows App Version

For example, the script in the folder:

WindowsApp-2.0.964.0

was validated against:

Windows App version 2.0.964.0

If Microsoft changes the Windows App UI again, a new folder will be added with a validated script for that version.

## Installation Instructions

On the Launcher machine, locate the default connector script:

C:\Program Files\Login VSI\Login Enterprise Launcher\Windows365ConnectorScript.cs

1. Back up the existing Windows365ConnectorScript.cs file.
2. Download the replacement Windows365ConnectorScript.cs from the appropriate Windows App version folder in this repository.
3. Copy the file into:

C:\Program Files\Login VSI\Login Enterprise Launcher\

4. Do not rename the file. The Windows 365 Connector connection type depends on the filename and path remaining exactly:

Windows365ConnectorScript.cs

After replacing the file, run your Login Enterprise test scenario again.