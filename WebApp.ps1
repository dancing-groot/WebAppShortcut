<#
.SYNOPSIS
  Shortcut creator for URLs
.DESCRIPTION
  Create a shortcut to a URL on the Desktop, Start Menu or Startup folder
  If running as SYSTEM, these will be machine wide folders
  If running as the current user, these will be the user folders
  You can enter the variables at the command line or in a JSON file.
  Additionally, the JSON file is designed to hold multiple entries
  If compiled as an EXE, you can drag and drop a JSON file to the executable
.LINK
  https://github.com/dancing-groot/WebAppShortcut
.NOTES
  Version:  2023.06.06
  Author:   Alexandre Cop
  Info:     Implemented a JSON file with multiple entries
            Icons can also be placed in the same folder as this script for deployment
#>
[cmdletbinding(DefaultParameterSetName = "Config")]
param(
  [Parameter(ParameterSetName = "Config")]
  [Parameter(ParameterSetName = "Custom")]
  [ValidateSet('Install', 'Uninstall')][string]$DeploymentType = 'Install',

  # ValueFromRemainingArguments allows to drag and drop a JSON file on the EXE
  [Parameter(ValueFromRemainingArguments = $true, Mandatory = $false, ParameterSetName = "Config")]
  [string]$Config,

  [Parameter(Mandatory = $true, ParameterSetName = "Custom")]
  [string]$DisplayName,

  [Parameter(Mandatory = $false, ParameterSetName = "Custom")]
  [string]$Url,

  [Parameter(Mandatory = $false, ParameterSetName = "Custom")]
  [string]$IconUrl,
  
  [Parameter(Mandatory = $false, ParameterSetName = "Custom")]
  [switch]$OnDesktop = $false,

  [Parameter(Mandatory = $false, ParameterSetName = "Custom")]
  [switch]$InStartMenu = $false,

  [Parameter(Mandatory = $false, ParameterSetName = "Custom")]
  [switch]$InStartup = $false
)

#region FUNCTIONS
function Add-WebApp
{
  [CmdletBinding()]
  param
  (
    [Parameter(ValueFromPipeline, ValueFromPipelinebyPropertyName)]
    [object[]]$NewShortcut
  )

  begin
  {
    if ([Security.Principal.WindowsIdentity]::GetCurrent().Name -eq "NT AUTHORITY\SYSTEM")
    {
      # Running in SYSTEM context (Shortcut applies for all users)
      $shortcutFolders = @{
        Icon      = "$($env:ProgramData)\Icons"
        Desktop   = "$env:PUBLIC\Desktop"
        Startup   = "$env:ALLUSERSPROFILE\Microsoft\Windows\Start Menu\Programs\Startup"
        StartMenu = "$env:ALLUSERSPROFILE\Microsoft\Windows\Start Menu\Programs"
      }
    }
    else
    {
      # Running in USER context (Shortcut only applies for the logged on user)
      $shortcutFolders = @{
        Icon      = "$($env:AppData)\Icons"
        Desktop   = "$([Environment]::GetFolderPath("Desktop"))"
        Startup   = "$env:AppData\Microsoft\Windows\Start Menu\Programs\Startup"
        StartMenu = "$env:AppData\Microsoft\Windows\Start Menu\Programs"
      }
    }
  }

  process
  {
    $wScriptShell = New-Object -ComObject WScript.Shell
    $localIconPath = $null

    ### Download icon for URL ###
    $localIconPath = "$($shortcutFolders.Icon)\$($NewShortcut.DisplayName).ico"
    # Check if target folder exists
    if (-not (Test-Path $shortcutFolders.Icon))
    {
      New-Item -Path $shortcutFolders.Icon -ItemType Directory
    }
    # Delete the existing icon if it exists already
    if (Test-Path $localIconPath)
    {
      Remove-Item $localIconPath -Force
    }
    # Download the icon
    if ($NewShortcut.IconUrl -match "^(http).*")
    {
      (New-Object System.Net.WebClient).DownloadFile($NewShortcut.IconUrl, $localIconPath)
    }
    else
    {
      Copy-Item -Path "$($Script:ScriptPath)\$($NewShortcut.IconUrl)" -Destination $localIconPath
    }

    ### Create shortcut on the relevant Desktop folder if required ###
    if ($NewShortcut.OnDesktop)
    {
      $shortcut = $wScriptShell.CreateShortcut("$($shortcutFolders.Desktop)\$($NewShortcut.DisplayName).lnk") 
      $shortcut.TargetPath = $NewShortcut.Url
      $shortcut.IconLocation = $localIconPath
      $shortcut.Save()
    }

    ### Create shortcut in the relevant Startup folder if required ###
    if ($NewShortcut.InStartMenu)
    {
      $shortcut = $wScriptShell.CreateShortcut("$($shortcutFolders.StartMenu)\$($NewShortcut.DisplayName).lnk") 
      $shortcut.TargetPath = $NewShortcut.Url 
      $shortcut.IconLocation = $localIconPath
      $shortcut.Save()
    }

    ### Create shortcut in the relevant Startup folder if required ###
    if ($NewShortcut.InStartup)
    {
      $shortcut = $wScriptShell.CreateShortcut("$($shortcutFolders.Startup)\$($NewShortcut.DisplayName).lnk") 
      $shortcut.TargetPath = $NewShortcut.Url
      $shortcut.IconLocation = $localIconPath
      $shortcut.Save()
    }
  }
}

function Remove-WebApp
{
  [CmdletBinding()]
  param
  (
    [Parameter(ValueFromPipeline, ValueFromPipelinebyPropertyName)]
    [object[]]$ShortcutToRemove
  )
  
  begin
  {
    if ([Security.Principal.WindowsIdentity]::GetCurrent().Name -eq "NT AUTHORITY\SYSTEM")
    {
      # Running in SYSTEM context (Shortcut applies for all users)
      $shortcutFolders = @{
        Icon      = "$($env:ProgramData)\Icons"
        Desktop   = "$env:PUBLIC\Desktop"
        Startup   = "$env:ALLUSERSPROFILE\Microsoft\Windows\Start Menu\Programs\Startup"
        StartMenu = "$env:ALLUSERSPROFILE\Microsoft\Windows\Start Menu\Programs"
      }
    }
    else
    {
      # Running in USER context (Shortcut only applies for the logged on user)
      $shortcutFolders = @{
        Icon      = "$($env:AppData)\Icons"
        Desktop   = "$([Environment]::GetFolderPath("Desktop"))"
        Startup   = "$env:AppData\Microsoft\Windows\Start Menu\Programs\Startup"
        StartMenu = "$env:AppData\Microsoft\Windows\Start Menu\Programs"
      }
    }
  }
  
  process
  {
    $shortcutName = $ShortcutToRemove.DisplayName
    $filesToRemove = @("$($shortcutFolders.Icon)\$shortcutName.ico", "$($shortcutFolders.Desktop)\$shortcutName.lnk", "$($shortcutFolders.Startup)\$shortcutName.lnk", "$($shortcutFolders.StartMenu)\$shortcutName.lnk")

    foreach ($file in $filesToRemove)
    {
      if (Test-Path $file)
      {
        Remove-Item $file -Force
      }
    }
  }
}

function Get-ScriptPath
{
  param
  (
    $Command
  )
  
  $script = @{"Path" = "."; "Stub" = "" }
  if ($Command.CommandType -eq "ExternalScript")
  {
    # Script
    $script.Path = Split-Path -Parent -Path $Command.Definition
    $script.Stub = ((Split-Path -Path $MyInvocation.ScriptName -Leaf) -split "\.")[0]
  }
  else
  {
    # Executable
    if (Split-Path -Parent -Path ([Environment]::GetCommandLineArgs()[0]))
    {
      $script.Path = Split-Path -Parent -Path ([Environment]::GetCommandLineArgs()[0])
      $script.Stub = ((Split-Path -Path ([Environment]::GetCommandLineArgs()[0]) -Leaf) -split "\.")[0]
    }
  }

  return $script
}
#endregion FUNCTIONS

#region DECLARATION
# You need to grab MyCommand from the body of the script, or it will return 'Function'
$thisScript = Get-ScriptPath -Command $MyInvocation.MyCommand
$ScriptPath = $thisScript.Path
$ScriptStub = $thisScript.Stub

switch ($PSCmdlet.ParameterSetName)
{
  'Config'
  {
    switch ($true)
    {
      (($Config -ne "") -and (Test-Path "$ScriptPath\$Config"))
      {
        $Config = "$ScriptPath\$Config"
      }

      (($Config -ne "") -and (Test-Path "$Config"))
      {
        $Config = "$Config"
      }

      (Test-Path -Path "$ScriptPath\$ScriptStub.json")
      {
        $Config = "$ScriptPath\$ScriptStub.json"
      }

      (Test-Path -Path "$ScriptPath\WebApp.json")
      {
        $Config = "$ScriptPath\WebApp.json"
      }

      default
      {
        Write-Host "No configuration file provided" -ForegroundColor Red
        exit
      }
    }

    $WebApp = (Get-Content -Path $Config | ConvertFrom-Json).WebApps
  }

  'Custom'
  {
    # This is creating an array of one so we can use the same logic as with the JSON file
    $WebApp = @(
      @{
        DisplayName = $DisplayName
        Url         = $Url
        IconUrl     = $IconUrl
        OnDesktop   = $OnDesktop
        InStartMenu = $InStartMenu
        InStartup   = $InStartup
      }
    )
  }
}
#endregion DECLARATION

#region MAIN
switch ($DeploymentType)
{
  "Install" { $WebApp | Add-WebApp }
  "Uninstall" { $WebApp | Remove-WebApp }
}
#endregion MAIN