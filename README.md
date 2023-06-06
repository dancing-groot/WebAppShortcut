# WebAppShortcut
Shortcut creator for URLs

## Description
You can use this script to create desktop shortcuts for a URL.

The shortcut can be created on the Desktop, Start Menu, Startup folder, or a combination of those.
Shortcuts are installed machine wide or only for the logged in user, depending on who launches the script.

Icons are cached locally and can either be downloaded from a URL, or copied from the same folder as the script.

You can create a single shortcut by calling the script with arguments, or create one or more by using a JSON file

## Parameters
### When using a JSON file
#### DeploymentType (string)
Set to `Install` or `Uninstall`
#### Config (string)
The name of the JSON file, located in the same folder as the script. The value can be:
- The name of the JSON file in the same folder as the script, e.g. `MyShortcuts.json`
- The full path to the JSON file, e.g. `C:\Company\Shortcuts\ThisShortcut.json`
- If the value is empty, it can be `WebApp.json` or matching the name of the script, e.g. if the script is CompanyShortcuts.ps1 then the value is `CompanyShortcuts.json`

### When using the script arguments
#### DeploymentType (string)
Set to 'Install' or 'Uninstall'
#### DisplayName (string)
The name shown for the shortcut
#### Url (string)
The website to point to, e.g. `https://github.com`
#### IconUrl (string)
The location where the shortcut icon can be downloaded or copied from the same folder as the script, e.g. `https://github.com/favicon.ico` or `github.ico`
#### OnDesktop (switch)
Use this argument as is if you want to create the shortcut on the Desktop 
#### InStartMenu (switch)
Use this argument as is if you want to create the shortcut in the Start Menu
#### InStartup (switch)
Use this argument as is if you want to create the shortcut in the Startup folder
