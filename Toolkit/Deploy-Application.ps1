<#
.SYNOPSIS
	This script performs the installation or uninstallation of an application(s).
	# LICENSE #
	PowerShell App Deployment Toolkit - Provides a set of functions to perform common application deployment tasks on Windows.
	Copyright (C) 2017 - Sean Lillis, Dan Cunningham, Muhammad Mashwani, Aman Motazedian.
	This program is free software: you can redistribute it and/or modify it under the terms of the GNU Lesser General Public License as published by the Free Software Foundation, either version 3 of the License, or any later version. This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
	You should have received a copy of the GNU Lesser General Public License along with this program. If not, see <http://www.gnu.org/licenses/>.
.DESCRIPTION
	The script is provided as a template to perform an install or uninstall of an application(s).
	The script either performs an "Install" deployment type or an "Uninstall" deployment type.
	The install deployment type is broken down into 3 main sections/phases: Pre-Install, Install, and Post-Install.
	The script dot-sources the AppDeployToolkitMain.ps1 script which contains the logic and functions required to install or uninstall an application.
.PARAMETER DeploymentType
	The type of deployment to perform. Default is: Install.
.PARAMETER DeployMode
	Specifies whether the installation should be run in Interactive, Silent, or NonInteractive mode. Default is: Interactive. Options: Interactive = Shows dialogs, Silent = No dialogs, NonInteractive = Very silent, i.e. no blocking apps. NonInteractive mode is automatically set if it is detected that the process is not user interactive.
.PARAMETER AllowRebootPassThru
	Allows the 3010 return code (requires restart) to be passed back to the parent process (e.g. SCCM) if detected from an installation. If 3010 is passed back to SCCM, a reboot prompt will be triggered.
.PARAMETER TerminalServerMode
	Changes to "user install mode" and back to "user execute mode" for installing/uninstalling applications for Remote Destkop Session Hosts/Citrix servers.
.PARAMETER DisableLogging
	Disables logging to file for the script. Default is: $false.
.EXAMPLE
    powershell.exe -Command "& { & '.\Deploy-Application.ps1' -DeployMode 'Silent'; Exit $LastExitCode }"
.EXAMPLE
    powershell.exe -Command "& { & '.\Deploy-Application.ps1' -AllowRebootPassThru; Exit $LastExitCode }"
.EXAMPLE
    powershell.exe -Command "& { & '.\Deploy-Application.ps1' -DeploymentType 'Uninstall'; Exit $LastExitCode }"
.EXAMPLE
    Deploy-Application.exe -DeploymentType "Install" -DeployMode "Silent"
.NOTES
	Toolkit Exit Code Ranges:
	60000 - 68999: Reserved for built-in exit codes in Deploy-Application.ps1, Deploy-Application.exe, and AppDeployToolkitMain.ps1
	69000 - 69999: Recommended for user customized exit codes in Deploy-Application.ps1
	70000 - 79999: Recommended for user customized exit codes in AppDeployToolkitExtensions.ps1
.LINK
	http://psappdeploytoolkit.com
#>
[CmdletBinding()]
Param (
	[Parameter(Mandatory=$false)]
	[ValidateSet('Install','Uninstall','Repair')]
	[string]$DeploymentType = 'Install',
	[Parameter(Mandatory=$false)]
	[ValidateSet('Interactive','Silent','NonInteractive')]
	[string]$DeployMode = 'Interactive',
	[Parameter(Mandatory=$false)]
	[switch]$AllowRebootPassThru = $false,
	[Parameter(Mandatory=$false)]
	[switch]$TerminalServerMode = $false,
	[Parameter(Mandatory=$false)]
	[switch]$DisableLogging = $false
)

Try {
	## Set the script execution policy for this process
	Try { Set-ExecutionPolicy -ExecutionPolicy 'ByPass' -Scope 'Process' -Force -ErrorAction 'Stop' } Catch {}

	##*===============================================
	##* VARIABLE DECLARATION
	##*===============================================
	## Variables: Application
	[string]$appVendor = ''
	[string]$appName = ''
	[string]$appVersion = ''
	[string]$appArch = ''
	[string]$appLang = 'EN'
	[string]$appRevision = '01'
	[string]$appScriptVersion = '1.0.0'
	[string]$appScriptDate = 'XX/XX/20XX'
	[string]$appScriptAuthor = '<author name>'
	##*===============================================
	## Variables: Install Titles (Only set here to override defaults set by the toolkit)
	[string]$installName = ''
	[string]$installTitle = ''

	##* Do not modify section below
	#region DoNotModify

	## Variables: Exit Code
	[int32]$mainExitCode = 0

	## Variables: Script
	[string]$deployAppScriptFriendlyName = 'Deploy Application'
	[version]$deployAppScriptVersion = [version]'3.8.3'
	[string]$deployAppScriptDate = '30/09/2020'
	[hashtable]$deployAppScriptParameters = $psBoundParameters

	## Variables: Environment
	If (Test-Path -LiteralPath 'variable:HostInvocation') { $InvocationInfo = $HostInvocation } Else { $InvocationInfo = $MyInvocation }
	[string]$scriptDirectory = Split-Path -Path $InvocationInfo.MyCommand.Definition -Parent

	## Dot source the required App Deploy Toolkit Functions
	Try {
		[string]$moduleAppDeployToolkitMain = "$scriptDirectory\AppDeployToolkit\AppDeployToolkitMain.ps1"
		If (-not (Test-Path -LiteralPath $moduleAppDeployToolkitMain -PathType 'Leaf')) { Throw "Module does not exist at the specified location [$moduleAppDeployToolkitMain]." }
		If ($DisableLogging) { . $moduleAppDeployToolkitMain -DisableLogging } Else { . $moduleAppDeployToolkitMain }
	}
	Catch {
		If ($mainExitCode -eq 0){ [int32]$mainExitCode = 60008 }
		Write-Error -Message "Module [$moduleAppDeployToolkitMain] failed to load: `n$($_.Exception.Message)`n `n$($_.InvocationInfo.PositionMessage)" -ErrorAction 'Continue'
		## Exit the script, returning the exit code to SCCM
		If (Test-Path -LiteralPath 'variable:HostInvocation') { $script:ExitCode = $mainExitCode; Exit } Else { Exit $mainExitCode }
	}

	#endregion
	##* Do not modify section above
	##*===============================================
	##* END VARIABLE DECLARATION
	##*===============================================

	If ($deploymentType -ine 'Uninstall' -and $deploymentType -ine 'Repair') {
		##*===============================================
		##* PRE-INSTALLATION
		##*===============================================
		[string]$installPhase = 'Pre-Installation'

		## Show Welcome Message, close Internet Explorer if required, allow up to 3 deferrals, verify there is enough disk space to complete the install, and persist the prompt
		#Show-InstallationWelcome -CloseApps 'iexplore' -AllowDefer -DeferTimes 3 -CheckDiskSpace -PersistPrompt
		
		## Show Progress Message (with the default message)
		#Show-InstallationProgress
		
		## <Perform Pre-Installation tasks here>
        
        #region Custom pre steps

        # Check if required process is running and close it if it's running
		#Stop-RequiredProcess -Process "AcroRd32"
		
		<# installation behavior (
                System : Install with System privileges.
				User : Install with currently logged on user (Logon requirement should be set to "Only when a user is logged on"
				
				If not provided or provided with faulty value, "System" will be used!!!
				
				If settings this to user, -FullPath MUST be used with Get-FileName in above commands. This is because of scheduled task requires full path to filename
        )#>
        
		# Get installation files, add -FullPath to return full file path (needed if path is outside $dirFiles or in subdirectories of $dirFiles)
		#
		# Change type execution order by shuffling msi, msp and/or exe around:
		#	$customInstall = @(@("msi","msp","exe"),
		#	$customInstall = @(@("exe","msp","msi"),
		#
		# Types are executed by the order they are in the object,
		# meaning the first element of a msi object will be executed before the second element of a msi object
		# Same procedure for msp and exe

		# Get-FileName -Filter "*.msi" -Path "$dirFiles\data" -FullPath  ||  Get-FileName -Filter "*.msi" -Path "C:\SRC\Something" -FullPath
		# Get-FileName -Filter "*.msp" -Path "$dirFiles\data" -FullPath  ||  Get-FileName -Filter "*.msp" -Path "C:\SRC\Something" -FullPath
		# Get-FileName -Filter "*.exe" -Path "$dirFiles\data" -FullPath  ||  Get-FileName -Filter "*.exe" -Path "C:\SRC\Something" -FullPath
		$customInstall = @(@("msi","msp","exe"),
			@{
				type = "msi"
				msiFile = Get-FileName -Filter "*.msi"
				mstFile = Get-FileName -Filter "*.mst"
				params = "/qn"
				workingDirectory = "" # not used if installBehaviour is set to User
				installBehavior = Get-InstallationBehavior -Behavior System
			},
			@{
				type = "msi"
				msiFile = ""
				mstFile = ""
				params = ""
				workingDirectory = "" # not used if installBehaviour is set to User
				installBehavior = Get-InstallationBehavior -Behavior System
			},
			@{
				type = "msp"
				file = Get-FileName -Filter "*.msp"
				params = "/qn"
				installBehavior = Get-InstallationBehavior -Behavior System
			},
			@{
				type = "msp"
				file = ""
				params = ""
				installBehavior = Get-InstallationBehavior -Behavior System
			},
			@{
				type = "exe"
				file = Get-FileName -Filter "*.exe"
				params = "/S"
				workingDirectory = "" # not used if installBehaviour is set to User
				waitForProcess = "" # this can be used if 'file' spawns a separate file for install. this will then monitor the given process and not continue until it's finished
				windowStyle = "Hidden" # Hidden|Maximized|Minimized|Normal
				installBehavior = Get-InstallationBehavior -Behavior System
				ignoreExitCodes = "" # comma separated list of exit codes. Will not be evaluated if installBehavior is User!
			},
			@{
				type = "exe"
				file = ""
				params = ""
				workingDirectory = "" # not used if installBehaviour is set to User
				waitForProcess = "" # this can be used if 'file' spawns a separate file for install. this will then monitor the given process and not continue until it's finished
				windowStyle = "Hidden" # Hidden|Maximized|Minimized|Normal
				installBehavior = Get-InstallationBehavior -Behavior System
				ignoreExitCodes = "" # comma separated list of exit codes. Will not be evaluated if installBehavior is User!
			}
		)
		
		# get certificate, must return full file path
		$certFiles = @(
			@{
				file = Get-FileName -Filter "*.cer" -FullPath
			},
			@{
				file = ""
			}
		)

        # Get current ip
		#$currentIP = Get-ClientIPSegment
		
		# uninstall VC++ 2015,2017,2019,2015_2019 x86_x64 - (WORKAROUND TO ALLOW SHITTY APPS TO INSTALL AN APP-REQUIRED OLDER VC++ THAN ALREADY INSTALLED.....)
		[bool]$allowShittyApps = $False

        #endregion

		##*===============================================
		##* INSTALLATION
		##*===============================================
		[string]$installPhase = 'Installation'

		## <Perform Installation tasks here>

		# install certificates as a trusted publisher before installation to allow silent install of drivers
		$certFiles | % {
			if ($_.file)
			{
				Start-CertificateInstallation -Path $_.file -HaltScriptOnError
			}
		}
		
		# uninstall VC++ 2015,2017,2019,2015_2019 x86_x64
		if ($allowShittyApps)
		{
			$allowShittyApps = Start-VCUninstall
		}
		
		#region Custom install steps

        Start-CustomInstall

		#endregion
		
		# install VC++ from "VisualStudioPlusPlusPath" (2015_2019 x86_x64)
		if ($allowShittyApps)
		{
			Start-VCInstall
		}

		##*===============================================
		##* POST-INSTALLATION
		##*===============================================
		[string]$installPhase = 'Post-Installation'

		## <Perform Post-Installation tasks here>

		## Display a message at the end of the install
		#If (-not $useDefaultMsi) { Show-InstallationPrompt -Message 'You can customize text to appear at the end of an install or remove it completely for unattended installations.' -ButtonRightText 'OK' -Icon Information -NoWait }
	}
	ElseIf ($deploymentType -ieq 'Uninstall')
	{
		##*===============================================
		##* PRE-UNINSTALLATION
		##*===============================================
		[string]$installPhase = 'Pre-Uninstallation'

		## Show Welcome Message, close Internet Explorer with a 60 second countdown before automatically closing
		#Show-InstallationWelcome -CloseApps 'iexplore' -CloseAppsCountdown 60
		
		## Show Progress Message (with the default message)
		#Show-InstallationProgress
		
		## <Perform Pre-Uninstallation tasks here>
        
        #region Custom pre steps

        # Check if required process is running and close it if it's running
		#Stop-RequiredProcess -Process "AcroRd32"
		
		<# uninstallation behavior (
                System : Uninstall with System privileges.
				User : Uninstall with currently logged on user (Logon requirement should be set to "Only when a user is logged on"
				
				If not provided or provided with faulty value, "System" will be used!!!
				
				If settings this to user, -FullPath MUST be used with Get-FileName in above commands. This is because of scheduled task requires full path to filename
        )#>

		# Get uninstallation files, add -FullPath to return full file path (needed if path is outside $dirFiles or in subdirectories of $dirFiles)
		#
		# Change type execution order by shuffling msi, and/or exe around:
		#	$customUninstall = @(@("msi","exe"),
		#	$customUninstall = @(@("exe",msi"),
		#
		# Types are executed by the order they are in the object,
		# meaning the first element of a msi object will be executed before the second element of a msi object
		# Same procedure for exe

		# Get-FileName -Filter "*.msi" -Path "$dirFiles\data" -FullPath  ||  Get-FileName -Filter "*.msi" -Path "C:\SRC\Something" -FullPath
		# Get-FileName -Filter "*.exe" -Path "$dirFiles\data" -FullPath  ||  Get-FileName -Filter "*.exe" -Path "C:\SRC\Something" -FullPath
		$customUninstall = @(@("msi","exe"),
			@{
				type = "msi"
				msiFile = Get-FileName -Filter "*.msi"
				params = "/qn"
				workingDirectory = "" # not used if uninstallBehaviour is set to User
				uninstallBehavior = Get-InstallationBehavior -Behavior System
			},
			@{
				type = "msi"
				msiFile = ""
				params = ""
				workingDirectory = "" # not used if uninstallBehaviour is set to User
				uninstallBehavior = Get-InstallationBehavior -Behavior System
			},
			@{
				type = "exe"
				file = Get-FileName -Filter "*.exe"
				params = "/S"
				workingDirectory = "" # not used if uninstallBehaviour is set to User
				waitForProcess = "" # this can be used if 'file' spawns a separate file for uninstall. this will then monitor the given process and not continue until it's finished
				windowStyle = "Hidden" # Hidden|Maximized|Minimized|Normal
				uninstallBehavior = Get-InstallationBehavior -Behavior System
				ignoreExitCodes = "" # comma separated list of exit codes. Will not be evaluated if uninstallBehavior is User!
			},
			@{
				type = "exe"
				file = ""
				params = ""
				workingDirectory = "" # not used if uninstallBehaviour is set to User
				waitForProcess = "" # this can be used if 'file' spawns a separate file for uninstall. this will then monitor the given process and not continue until it's finished
				windowStyle = "Hidden" # Hidden|Maximized|Minimized|Normal
				uninstallBehavior = Get-InstallationBehavior -Behavior System
				ignoreExitCodes = "" # comma separated list of exit codes. Will not be evaluated if uninstallBehavior is User!
			}
		)

        #endregion

		##*===============================================
		##* UNINSTALLATION
		##*===============================================
		[string]$installPhase = 'Uninstallation'

		# <Perform Uninstallation tasks here>
		
        #region Custom uninstall steps

        Start-CustomUninstall

		#endregion
		
		# wait for uninstall to finish
		#Wait-ForProgramToFinish -ProcessName "_uninstall*"
		#Wait-TimeForProgramToFinish -Minutes 1

		##*===============================================
		##* POST-UNINSTALLATION
		##*===============================================
		[string]$installPhase = 'Post-Uninstallation'

		## <Perform Post-Uninstallation tasks here>


	}
	ElseIf ($deploymentType -ieq 'Repair')
	{
		##*===============================================
		##* PRE-REPAIR
		##*===============================================
		[string]$installPhase = 'Pre-Repair'

		## Show Progress Message (with the default message)
		Show-InstallationProgress

		## <Perform Pre-Repair tasks here>

		##*===============================================
		##* REPAIR
		##*===============================================
		[string]$installPhase = 'Repair'

		## Handle Zero-Config MSI Repairs
		If ($useDefaultMsi) {
			[hashtable]$ExecuteDefaultMSISplat =  @{ Action = 'Repair'; Path = $defaultMsiFile; }; If ($defaultMstFile) { $ExecuteDefaultMSISplat.Add('Transform', $defaultMstFile) }
			Execute-MSI @ExecuteDefaultMSISplat
		}
		# <Perform Repair tasks here>

		##*===============================================
		##* POST-REPAIR
		##*===============================================
		[string]$installPhase = 'Post-Repair'

		## <Perform Post-Repair tasks here>


    }
	##*===============================================
	##* END SCRIPT BODY
	##*===============================================

	## Call the Exit-Script function to perform final cleanup operations
	Exit-Script -ExitCode $mainExitCode
}
Catch {
	[int32]$mainExitCode = 60001
	[string]$mainErrorMessage = "$(Resolve-Error)"
	Write-Log -Message $mainErrorMessage -Severity 3 -Source $deployAppScriptFriendlyName
	Show-DialogBox -Text $mainErrorMessage -Icon 'Stop'
	Exit-Script -ExitCode $mainExitCode
}
