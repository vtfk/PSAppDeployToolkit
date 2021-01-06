# PSAppDeployToolkit

## Toolkit

This is our version of [PSAppDeployToolKit](https://github.com/PSAppDeployToolkit/PSAppDeployToolkit). It's the same toolkit but with some extra features!

### Setup

- Download the latest [release](https://github.com/vtfk/PSAppDeployToolkit/releases)
- Unzip the file
- Open `Toolkit\AppDeployToolkit\CustomFunctionsConfig.xml` in your editor
    - Set `VisualStudioPlusPlusPath` to an UNC folder path containing your vc++ installation filen `(*.exe)` (no subfolders)
    - Edit `IPSegment` to your IP range (wildcard `*` accepted) and/or add more `IPSegment` elements if you have more IP ranges

### Usage

1. Add setup file(s) in the `Files` folder
1. Open `Toolkit\Deploy-Application.ps1`
1. Installation
    1. In the section `## <Perform Pre-Installation tasks here>`, in the hash table `$customInstall`, update each hash table with info for the setup file(s) you want to install.
1. Uninstallation
    1. In the section `## <Perform Pre-Uninstallation tasks here>`, in the hash table `$customUninstall`, update each hash table with info for the setup file or msi guid you want to uninstall.

## New-PDT

### Setup

- Download New-PDT.ps1 to your machine and open it in your editor
    - Edit `ToolkitPath` in the `settings` hash table to a local path or an UNC path containing the toolkit
    - Edit `LogPath` in the `settings` hash table to a local path or an UNC path to where you want the log file. If set to an empty string, no logging will be performed
- Create a shortcut on your/public desktop called `New-PDT` and set target to: `C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -NoLogo -NoProfile -ExecutionPolicy ByPass -File "Full-Path-To-New-PDT.ps1"`

### Usage

- Start `New-PDT` from desktop
- **Or** 
- Open a powershell window, go to where you have saved the `New-PDT.ps1` file and run `.\New-PDT.ps1`

## Toolkit Custom functions

### Stop-RequiredProcess

`Parameters`
- Process
    - String array with process names (without extension)

Will force stop the given processes

### Get-FileName

`Parameters`
- Filter
    - Mandatory string to find a file by name: `Example '*.msi'`
- Path
    - String to where to find the file. Default is `$dirFiles`
- FullPath
    - Switch. If set, full file path will be returned. Otherwise, only file name will be returned

Get filename based on a filter: `"*.msi", "*.mst", "*.msp", "*.exe" and so forth from $dirFiles or $dirSupportFiles or specified location.`

If Path is set to `$dirFiles`, `$dirSupportFiles` or not set at all, this function will look in both `$dirFiles` and `$dirSupportFiles` to find a file to match Filter.

If Filter is found in `$dirFiles` it will not check in `$dirSupportFiles`and vica/versa.

### Get-ClientIPSegment

Will return the gateway octet from client ip address IF client ip matches the IP segment(s) set in [`CustomFunctionsConfig.xml`](https://github.com/vtfk/PSAppDeployToolkit#setup).

If IP segment(s) is not setup in `CustomFunctionsConfig.xml` or client doesn't have any matching ip address, this function will fail.

### Get-InstallationBehavior

`Parameters`
- Behavior
    - Mandatory string to indicate if file should be installed with 'System' or 'User'

Indicate if file should be installed with 'System' or 'User'.

If 'System', file will be installed with SYSTEM account on client

If 'User', file will be installed with currently logged on user (at toolkit runtime) on client

### Start-CustomInstall

Parses the hash table `$customInstall` and installs the file(s) with set info

### Start-CustomUninstall

Parses the hash table `$customUninstall` and uninstalls the file(s)/guid(s) with set info

### Wait-ForProgramToFinish

`Parameters`
- ProcessName
    - Mandatory string to indicate which process name (without extension) to wait for to finish

### Wait-TimeForProgramToFinish

`Parameters`
- Minutes
    - Integer to indicate how many minutes to wait. Default is 1 minute

### Get-MsiProperty

`Parameters`
- Path
    - Mandatory string to the msi file
- Property
    - Mandatory string to indicate which property to retrieve. Set of values is present. If this is set, `CustomProperty` will not be available
- CustomProperty
    - Mandatory string to indicate which custom property to retrieve. If this is set, `Property` will not be available

Get MSI properties (Must be single quoted. `Example: "'ProductCode'"`)