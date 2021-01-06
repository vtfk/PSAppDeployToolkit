# Import config xml file ($directoryRoot comes from AppDeployToolkitExtenions.ps1)
[string]$configFilePath = Join-Path -Path $directoryRoot -ChildPath "CustomFunctionsConfig.xml"
[XML]$CustomConfig = [XML](Get-Content -Path $configFilePath)

# Check if required process is running and close it if it's running
Function Stop-RequiredProcess
{
    ## ExtensionAppDeployScript
    param(
        [Parameter(Mandatory = $True)]
        [string[]]$Process
    )

    foreach ($ProcessName in $Process)
    {
        Write-Log -Message "Checking if $ProcessName is running" -Severity 1 -Source "CustomFunctions.ps1" -ScriptSection "Stop-RequiredProcesses" -LogType CMTrace
        Get-Process -Name $ProcessName -ErrorAction SilentlyContinue | % { Write-Log -Message "Closing '$($_.Name)' with Id '$($_.Id)'" -Severity 1 -Source "CustomFunctions.ps1" -ScriptSection "Stop-RequiredProcesses" -LogType CMTrace; $_ | Stop-Process -Force -Confirm:$False }

        Start-Sleep -Milliseconds 1500

        if ((Get-Process -Name $ProcessName -ErrorAction SilentlyContinue))
        {
            Write-Log -Message "$ProcessName is still running. Aborting..." -Severity 3 -Source "CustomFunctions.ps1" -ScriptSection "Stop-RequiredProcesses" -LogType CMTrace
            Exit-Script -ExitCode 16001
        }
    }
}

# Get filename based on a filter: "*.msi", "*.mst", "*.msp", "*.exe" and so forth from $dirFiles or $dirSupportFiles or specified location.
Function Get-FileName
{
    ## ExtensionAppDeployScript
    param (
        [Parameter(Mandatory = $True)]
        [ValidateNotNullOrEmpty()]
        [string]$Filter,

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]$Path = $dirFiles,

        [Parameter()]
        [switch]$FullPath
    )

    $file = Get-ChildItem -Path $Path -Filter $Filter -File | Select -ExpandProperty FullName
    if (!$file)
    {
        if ($Path.StartsWith($dirFiles))
        {
            # maybe it exists in $dirSupportFiles :(
            Write-Log -Message "'$Filter' not found in '$Path'" -Severity 2 -Source "CustomFunctions.ps1" -ScriptSection "Get-FileName" -LogType CMTrace
            $Path = $Path.Replace($dirFiles, $dirSupportFiles)
            Write-Log -Message "Will try to find '$Filter' in '$dirSupportFiles'" -Severity 2 -Source "CustomFunctions.ps1" -ScriptSection "Get-FileName" -LogType CMTrace
            $file = Get-ChildItem -Path $Path -Filter $Filter | Select -ExpandProperty VersionInfo | Select -ExpandProperty FileName
        }
        elseif ($Path.StartsWith($dirSupportFiles))
        {
            # maybe it exists in $dirFiles :)
            Write-Log -Message "'$Filter' not found in '$Path'" -Severity 2 -Source "CustomFunctions.ps1" -ScriptSection "Get-FileName" -LogType CMTrace
            $Path = $Path.Replace($dirSupportFiles, $dirFiles)
            Write-Log -Message "Will try to find '$Filter' in '$dirFiles'" -Severity 2 -Source "CustomFunctions.ps1" -ScriptSection "Get-FileName" -LogType CMTrace
            $file = Get-ChildItem -Path $Path -Filter $Filter | Select -ExpandProperty VersionInfo | Select -ExpandProperty FileName
        }
    }

    if (!$file)
    {
        Write-Log -Message "'$Filter' not found in '$Path'" -Severity 2 -Source "CustomFunctions.ps1" -ScriptSection "Get-FileName" -LogType CMTrace
        return ""
    }
    else
    {
        if ($file.Count -le 0 -or $file.Count -gt 1)
        {
            Write-Log -Message "'$Filter' in '$Path' : $($file.Count) items found ('$($file -join "', '")'). Give a more specific filter" -Severity 2 -Source "CustomFunctions.ps1" -ScriptSection "Get-FileName" -LogType CMTrace
            return ""
        }
        else
        {
            Write-Log -Message "Using '$Filter' in '$Path' : '$file'" -Severity 1 -Source "CustomFunctions.ps1" -ScriptSection "Get-FileName" -LogType CMTrace
        }

        if ($FullPath)
        {
            Write-Log -Message "Returning full path : '$file'" -Severity 1 -Source "CustomFunctions.ps1" -ScriptSection "Get-FileName" -LogType CMTrace
        }
        else
        {
            $file = [System.IO.Path]::GetFileName($file)
            Write-Log -Message "Returning file name only : '$file'" -Severity 1 -Source "CustomFunctions.ps1" -ScriptSection "Get-FileName" -LogType CMTrace
        }

        return $file
    }
}

# Get second octet of IP Address
Function Get-ClientIPSegment
{
    ## ExtensionAppDeployScript
    ## GET CLIENT GATEWAY IP SEGMENT
    $IPAddresses = Get-WmiObject Win32_NetworkAdapterConfiguration | Where { $ipAddr = $_.IpAddress; $_.IpEnabled -eq $True -and ($CustomConfig.CustomExtensions.IPSegments.IPSegment | Where { $ipAddr -like $_ }) } | Select IPAddress
    if ($null -eq $IPAddresses)
    {
        Write-Log -Message "This machine is not connected to any of the given networks!" -Severity 2 -Source "CustomFunctions.ps1" -ScriptSection "Get-ClientIPSegment" -LogType "CMTrace" -WriteHost $True
        return -1
    }

    if ($null -ne $IPAddresses.Count)
    {
        Write-Log -Message "Found $($IPAddresses.Count) adapters with IP like '$($CustomConfig.CustomExtensions.IPSegments.IPSegment -join "' or '")' :" -Severity 1 -Source "CustomFunctions.ps1" -ScriptSection "Get-ClientIPSegment" -LogType "CMTrace"
        foreach ($IPA in $IPAddresses)
        {
            Write-Log -Message "- $($IPA.IPAddress)" -Severity 1 -Source "CustomFunctions.ps1" -ScriptSection "Get-ClientIPSegment" -LogType "CMTrace"
        }
        [string]$IPAddress = $IPAddresses[0].IPAddress
        Write-Log -Message "Using IPAddress '$IPAddress'" -Severity 1 -Source "CustomFunctions.ps1" -ScriptSection "Get-ClientIPSegment" -LogType "CMTrace"
        [string]$GatewayIP = $IPAddress.Substring(3, 3)
        Write-Log -Message "Found IP segment '$GatewayIP'" -Severity 1 -Source "CustomFunctions.ps1" -ScriptSection "Get-ClientIPSegment" -LogType "CMTrace"

        return $GatewayIP
    }
    else
    {
        [string]$IPAddress = $IPAddresses.IPAddress
        Write-Log -Message "Using IPAddress '$IPAddress'" -Severity 1 -Source "CustomFunctions.ps1" -ScriptSection "Get-ClientIPSegment" -LogType "CMTrace"
        [string]$GatewayIP = $IPAddress.Substring(3, 3)
        Write-Log -Message "Found IP segment '$GatewayIP'" -Severity 1 -Source "CustomFunctions.ps1" -ScriptSection "Get-ClientIPSegment" -LogType "CMTrace"

        return $GatewayIP
    }
}

<# Get Installation Behavior
        Installation behavior (
        System : Install with System privileges.
        User : Install with currently logged on user (Logon requirement should be set to "Only when a user is logged on"
        If not provided or provided with faulty value, "System" will be used!!!
)
#>
Function Get-InstallationBehavior
{
    ## ExtensionAppDeployScript
    param(
        [Parameter(Mandatory = $True)]
        [ValidateSet("System", "User")]
        [string]$Behavior
    )

    if (!$Behavior)
    {
        Write-Log -Message "Unexpected parameter given. Will use System as fallback" -Severity 2 -Source "CustomFunctions.ps1" -ScriptSection "Get-InstallationBehavior" -LogType CMTrace
        return "System"
    }
    elseif ($Behavior -eq "System" -or $Behavior -eq "User")
    {
        Write-Log -Message "'$Behavior' parameter given. Will use '$Behavior'" -Severity 1 -Source "CustomFunctions.ps1" -ScriptSection "Get-InstallationBehavior" -LogType CMTrace
        return $Behavior
    }
    else
    {
        Write-Log -Message "Unexpected parameter given. Will use System as fallback" -Severity 2 -Source "CustomFunctions.ps1" -ScriptSection "Get-InstallationBehavior" -LogType CMTrace
        return "System"
    }
}

# Custom install based on arguments given in Deploy-Application.ps1
Function Start-CustomInstall
{
    ## ExtensionAppDeployScript
    if (!$customInstall -or !$customInstall.Count -or $customInstall.Count -le 1 -or $customInstall[0].GetType().Name -ne "Object[]" -or $customInstall[1].GetType().Name -ne "Hashtable") {
        Write-Log -Message "Required variable 'customInstall' missing or invalid configured. Read the comments in Deploy-Application.ps1" -Severity 3 -Source "CustomFunctions.ps1" -ScriptSection "Start-CustomInstall" -LogType CMTrace
        Exit-Script -ExitCode 60070
    }

    $customInstall[0] | % {
        $type = $_
        $customInstall | Where { $_.type -eq $type } | % {
            if ($_.type -eq "msi") { Start-CustomMsi -Obj $_ -Type Install }
            elseif ($_.type -eq "msp") { Start-CustomMsp -Obj $_ }
            elseif ($_.type -eq "exe") { Start-CustomExe -Obj $_ -Type Install }
            else { Write-Log -Message "Skipping unknown/unsupported file type ($type)" -Severity 2 -Source "CustomFunctions.ps1" -ScriptSection "Start-CustomInstall" -LogType CMTrace }
        }
    }
}

# Custom uninstall based on arguments given in Deploy-Application.ps1
Function Start-CustomUninstall
{
    ## ExtensionAppDeployScript
    if (!$customUninstall -or !$customUninstall.Count -or $customUninstall.Count -le 1 -or $customUninstall[0].GetType().Name -ne "Object[]" -or $customUninstall[1].GetType().Name -ne "Hashtable") {
        Write-Log -Message "Required variable 'customUninstall' missing or invalid configured. Read the comments in Deploy-Application.ps1" -Severity 3 -Source "CustomFunctions.ps1" -ScriptSection "Start-CustomUninstall" -LogType CMTrace
        Exit-Script -ExitCode 60070
    }

    $customUninstall[0] | % {
        $type = $_
        $customUninstall | Where { $_.type -eq $type } | % {
            if ($_.type -eq "msi") { Start-CustomMsi -Obj $_ -Type Uninstall }
            elseif ($_.type -eq "exe") { Start-CustomExe -Obj $_ -Type Uninstall }
            else { Write-Log -Message "Skipping unknown/unsupported file type ($type)" -Severity 2 -Source "CustomFunctions.ps1" -ScriptSection "Start-CustomUninstall" -LogType CMTrace }
        }
    }
}

Function Start-CustomMsi
{
    ## ExtensionAppDeployScript
    param(
        [Parameter(Mandatory = $True)]
        $Obj,

        [Parameter()]
        [ValidateSet("Install", "Uninstall")]
        [string]$Type = "Install"
    )

    if ($Type -eq "Install")
    {
        if ($Obj.msiFile)
        {
            # add msi info
            $msiSplat = @{
                Action = "Install"
                Path = $Obj.msiFile
            }
            if ($Obj.mstFile) { $msiSplat.Add("Transform", $Obj.mstFile) }
            if ($Obj.params) { $msiSplat.Add("Parameters", $Obj.params) }
            if ($Obj.workingDirectory) { $msiSplat.Add("WorkingDirectory", $Obj.workingDirectory) }

            # Install Windows Installer application
            if (![string]::IsNullOrEmpty($msiSplat.Path))
            {
                if (!$Obj.installBehavior -or $Obj.installBehavior -eq "System")
                {
                    # Install Windows Installer application as System
                    Write-Log -Message "Installing '$($msiSplat.Path)' with System account" -Severity 1 -Source "CustomFunctions.ps1" -ScriptSection "Start-CustomMsi" -LogType CMTrace
                    Execute-MSI @msiSplat
                }
                elseif ($Obj.installBehavior -eq "User")
                {
                    # Install Windows Installer application as User
                    $loggedOnUsers = Get-LoggedOnUser
                    Write-Log -Message "Logged on users: '$($loggedOnUsers.NTAccount -join ', ')'" -Severity 1 -Source "CustomFunctions.ps1" -ScriptSection "Start-CustomMsi" -LogType CMTrace
                    if (!$loggedOnUsers -or $loggedOnUsers.Count -le 0)
                    {
                        throw "No logged on users on this system. Installation aborted!"
                    }
                    else
                    {
                        # Add msi user info
                        $msiSplatUser = @{
                            Path = "msiexec.exe"
                            RunLevel = "HighestAvailable"
                            Wait = $True
                        }
                        
                        [string]$parameteresUser = "/i `"$($msiSplat.Path)`""
                        if (![string]::IsNullOrEmpty($msiSplat.Parameters)) { $parameteresUser += " $($msiSplat.Parameters)" }
                        if (![string]::IsNullOrEmpty($msiSplat.Transform)) { $parameteresUser += " TRANSFORMS=`"$($msiSplat.Transform)`"" }
                        $msiSplatUser.Add("Parameters", $parameteresUser)

                        # add username
                        if ($loggedOnUsers.Count -eq 1)
                        {
                            $msiSplatUser.Add("UserName", $loggedOnUsers.NTAccount)
                            Write-Log -Message "Installing '$($msiSplat.Path)' with logged on user account '$($loggedOnUsers.NTAccount)'" -Severity 1 -Source "CustomFunctions.ps1" -ScriptSection "Start-CustomMsi" -LogType CMTrace
                        }
                        elseif ($loggedOnUsers.Count -gt 1)
                        {
                            $msiSplatUser.Add("UserName", $loggedOnUsers[0].NTAccount)
                            Write-Log -Message "Installing '$($msiSplat.Path)' with logged on user account -- using '$($loggedOnUsers[0].NTAccount)' from '$($loggedOnUsers.NTAccount -join ', ')'" -Severity 1 -Source "CustomFunctions.ps1" -ScriptSection "Start-CustomMsi" -LogType CMTrace
                        }

                        # execute process as user
                        Execute-ProcessAsUser @msiSplatUser
                    }
                }
            }
        }
    }
    elseif ($Type -eq "Uninstall")
    {
        if ($Obj.msiFile)
        {
            # add msi info
            $msiSplat = @{
                Action = "Uninstall"
                Path = $Obj.msiFile
            }
            if ($Obj.mstFile) { $msiSplat.Add("Transform", $Obj.mstFile) }
            if ($Obj.params) { $msiSplat.Add("Parameters", $Obj.params) }
            if ($Obj.workingDirectory) { $msiSplat.Add("WorkingDirectory", $Obj.workingDirectory) }

            # Uninstall Windows Installer application
            if (![string]::IsNullOrEmpty($msiSplat.Path))
            {
                if (!$Obj.uninstallBehavior -or $Obj.uninstallBehavior -eq "System")
                {
                    # Uninstall Windows Installer application as System
                    Write-Log -Message "Uninstalling '$($msiSplat.Path)' with System account" -Severity 1 -Source "CustomFunctions.ps1" -ScriptSection "Start-CustomMsi" -LogType CMTrace
                    Execute-MSI @msiSplat
                }
                elseif ($Obj.uninstallBehavior -eq "User")
                {
                    # Uninstall Windows Installer application as User
                    $loggedOnUsers = Get-LoggedOnUser
                    Write-Log -Message "Logged on users: '$($loggedOnUsers.NTAccount -join ', ')'" -Severity 1 -Source "CustomFunctions.ps1" -ScriptSection "Start-CustomMsi" -LogType CMTrace
                    if (!$loggedOnUsers -or $loggedOnUsers.Count -le 0)
                    {
                        throw "No logged on users on this system. Uninstallation aborted!"
                    }
                    else
                    {
                        # Add msi user info
                        $msiSplatUser = @{
                            Path = "msiexec.exe"
                            RunLevel = "HighestAvailable"
                            Wait = $True
                        }
                        
                        [string]$parameteresUser = "/x `"$($msiSplat.Path)`""
                        if (![string]::IsNullOrEmpty($msiSplat.Parameters)) { $parameteresUser += " $($msiSplat.Parameters)" }
                        if (![string]::IsNullOrEmpty($msiSplat.Transform)) { $parameteresUser += " TRANSFORMS=`"$($msiSplat.Transform)`"" }
                        $msiSplatUser.Add("Parameters", $parameteresUser)

                        # add username
                        if ($loggedOnUsers.Count -eq 1)
                        {
                            $msiSplatUser.Add("UserName", $loggedOnUsers.NTAccount)
                            Write-Log -Message "Uninstalling '$($msiSplat.Path)' with logged on user account '$($loggedOnUsers.NTAccount)'" -Severity 1 -Source "CustomFunctions.ps1" -ScriptSection "Start-CustomMsi" -LogType CMTrace
                        }
                        elseif ($loggedOnUsers.Count -gt 1)
                        {
                            $msiSplatUser.Add("UserName", $loggedOnUsers[0].NTAccount)
                            Write-Log -Message "Uninstalling '$($msiSplat.Path)' with logged on user account -- using '$($loggedOnUsers[0].NTAccount)' from '$($loggedOnUsers.NTAccount -join ', ')'" -Severity 1 -Source "CustomFunctions.ps1" -ScriptSection "Start-CustomMsi" -LogType CMTrace
                        }

                        # execute process as user
                        Execute-ProcessAsUser @msiSplatUser
                    }
                }
            }
        }
    }
}

Function Start-CustomMsp
{
    ## ExtensionAppDeployScript
    param(
        [Parameter(Mandatory = $True)]
        $Obj
    )

    if ($Obj.file)
    {
        # add msp info
        $mspSplat = @{
            Action = "Patch"
            Path = $Obj.file
        }
        if ($Obj.params) { $mspSplat.Add("Parameters", $Obj.params) }

        # Patch Windows Installer application
        if (![string]::IsNullOrEmpty($mspSplat.Path))
        {
            if (!$Obj.installBehavior -or $Obj.installBehavior -eq "System")
            {
                # Patch Windows Installer application as System
                Write-Log -Message "Patching '$($mspSplat.Path)' with System account" -Severity 1 -Source "CustomFunctions.ps1" -ScriptSection "Start-CustomMsp" -LogType CMTrace
                Execute-MSI @mspSplat
            }
            elseif ($Obj.installBehavior -eq "User")
            {
                # Patch Windows Installer application as User
                $loggedOnUsers = Get-LoggedOnUser
                Write-Log -Message "Logged on users: '$($loggedOnUsers.NTAccount -join ', ')'" -Severity 1 -Source "CustomFunctions.ps1" -ScriptSection "Start-CustomMsp" -LogType CMTrace
                if (!$loggedOnUsers -or $loggedOnUsers.Count -le 0)
                {
                    throw "No logged on users on this system. Installation aborted!"
                }
                else
                {
                    # Add msp user info
                    $mspSplatUser = @{
                        Path = "msiexec.exe"
                        RunLevel = "HighestAvailable"
                        Wait = $True
                    }
                    [string]$parameteresUser = "/update `"$($mspSplat.Path)`""
                    if (![string]::IsNullOrEmpty($mspSplat.Parameters)) { $parameteresUser += " $($mspSplat.Parameters)" }
                    $mspSplatUser.Add("Parameters", $parameteresUser)

                    # add username
                    if ($loggedOnUsers.Count -eq 1)
                    {
                        $mspSplatUser.Add("UserName", $loggedOnUsers.NTAccount)
                        Write-Log -Message "Patching '$($mspSplat.Path)' with logged on user account '$($loggedOnUsers.NTAccount)'" -Severity 1 -Source "CustomFunctions.ps1" -ScriptSection "Start-CustomMsp" -LogType CMTrace
                    }
                    elseif ($loggedOnUsers.Count -gt 1)
                    {
                        $mspSplatUser.Add("UserName", $loggedOnUsers[0].NTAccount)
                        Write-Log -Message "Patching '$($mspSplat.Path)' with logged on user account -- using '$($loggedOnUsers[0].NTAccount)' from '$($loggedOnUsers.NTAccount -join ', ')'" -Severity 1 -Source "CustomFunctions.ps1" -ScriptSection "Start-CustomMsp" -LogType CMTrace
                    }

                    # execute process as user
                    Execute-ProcessAsUser @mspSplatUser
                }
            }
        }
    }
}

Function Start-CustomExe
{
    ## ExtensionAppDeployScript
    param(
        [Parameter(Mandatory = $True)]
        $Obj,

        [Parameter()]
        [ValidateSet("Install", "Uninstall")]
        [string]$Type = "Install"
    )

    if ($Type -eq "Install")
    {
        if ($Obj.file)
        {
            # add executable info
            $exeSplat = @{
                WindowStyle = $Obj.windowStyle
                Path = $Obj.file
                WaitForMsiExec = $True
            }
            if ($Obj.params) { $exeSplat.Add("Parameter", $Obj.params) }
            if ($Obj.workingDirectory) { $exeSplat.Add("WorkingDirectory", $Obj.workingDirectory) }
            if ($Obj.ignoreExitCodes) { $exeSplat.Add("IgnoreExitCodes", $Obj.ignoreExitCodes) }

            # Install exe application
            if (![string]::IsNullOrEmpty($exeSplat.Path))
            {
                if (!$Obj.installBehavior -or $Obj.installBehavior -eq "System")
                {
                    # Install exe application as System
                    Write-Log -Message "Installing '$($exeSplat.Path)' with System account and WindowStyle '$($exeSplat.WindowStyle)'" -Severity 1 -Source "CustomFunctions.ps1" -ScriptSection "Start-CustomExe" -LogType CMTrace
                    Execute-Process @exeSplat

                    if ($Obj.waitForProcess)
                    {
                        # wait for given process to finish before continuing
                        Wait-ForProgramToFinish -ProcessName $Obj.waitForProcess -CalledFromCustom
                    }
                }
                elseif ($Obj.installBehavior -eq "User")
                {
                    # Install exe application as User
                    $loggedOnUsers = Get-LoggedOnUser
                    Write-Log -Message "Logged on users: '$($loggedOnUsers.NTAccount -join ', ')'" -Severity 1 -Source "CustomFunctions.ps1" -ScriptSection "Start-CustomExe" -LogType CMTrace
                    if (!$loggedOnUsers -or $loggedOnUsers.Count -le 0)
                    {
                        throw "No logged on users on this system. Installation aborted!"
                    }
                    else
                    {
                        # Add exe user info
                        $exeSplatUser = @{
                            Path = $exeSplat.Path
                            RunLevel = "HighestAvailable"
                            Wait = $True
                        }
                        if (![string]::IsNullOrEmpty($exeSplat.Parameter))
                        {
                            $exeSplatUser.Add("Parameters", $exeSplat.Parameter)
                        }

                        # add username
                        if ($loggedOnUsers.Count -eq 1)
                        {
                            $exeSplatUser.Add("UserName", $loggedOnUsers.NTAccount)
                            Write-Log -Message "Installing '$($exeSplat.Path)' with logged on user account '$($loggedOnUsers.NTAccount)'" -Severity 1 -Source "CustomFunctions.ps1" -ScriptSection "Start-CustomExe" -LogType CMTrace
                        }
                        elseif ($loggedOnUsers.Count -gt 1)
                        {
                            $exeSplatUser.Add("UserName", $loggedOnUsers[0].NTAccount)
                            Write-Log -Message "Installing '$($exeSplat.Path)' with logged on user account -- using '$($loggedOnUsers[0].NTAccount)' from '$($loggedOnUsers.NTAccount -join ', ')'" -Severity 1 -Source "CustomFunctions.ps1" -ScriptSection "Start-CustomExe" -LogType CMTrace
                        }

                        # execute process as user
                        Execute-ProcessAsUser @exeSplatUser

                        if ($Obj.waitForProcess)
                        {
                            # wait for given process to finish before continuing (not sure if this will work when 'Path' is run in another user context)
                            Wait-ForProgramToFinish -ProcessName $Obj.waitForProcess -CalledFromCustom
                        }
                    }
                }
            }
        }
    }
    elseif ($Type -eq "Uninstall")
    {
        if ($Obj.file)
        {
            # add executable info
            $exeSplat = @{
                WindowStyle = $Obj.windowStyle
                Path = $Obj.file
                WaitForMsiExec = $True
            }
            if ($Obj.params) { $exeSplat.Add("Parameter", $Obj.params) }
            if ($Obj.workingDirectory) { $exeSplat.Add("WorkingDirectory", $Obj.workingDirectory) }
            if ($Obj.ignoreExitCodes) { $exeSplat.Add("IgnoreExitCodes", $Obj.ignoreExitCodes) }

            # Uninstall exe application
            if (![string]::IsNullOrEmpty($exeSplat.Path))
            {
                if (!$Obj.uninstallBehavior -or $Obj.uninstallBehavior -eq "System")
                {
                    # Uninstall exe application as System
                    Write-Log -Message "Uninstalling '$($exeSplat.Path)' with System account and WindowStyle '$($exeSplat.WindowStyle)'" -Severity 1 -Source "CustomFunctions.ps1" -ScriptSection "Start-CustomExe" -LogType CMTrace
                    Execute-Process @exeSplat

                    if ($Obj.waitForProcess)
                    {
                        # wait for given process to finish before continuing
                        Wait-ForProgramToFinish -ProcessName $Obj.waitForProcess -CalledFromCustom
                    }
                }
                elseif ($Obj.uninstallBehavior -eq "User")
                {
                    # Uninstall exe application as User
                    $loggedOnUsers = Get-LoggedOnUser
                    Write-Log -Message "Logged on users: '$($loggedOnUsers.NTAccount -join ', ')'" -Severity 1 -Source "CustomFunctions.ps1" -ScriptSection "Start-CustomExe" -LogType CMTrace
                    if (!$loggedOnUsers -or $loggedOnUsers.Count -le 0)
                    {
                        throw "No logged on users on this system. Uninstallation aborted!"
                    }
                    else
                    {
                        # Add exe user info
                        $exeSplatUser = @{
                            Path = $exeSplat.Path
                            RunLevel = "HighestAvailable"
                            Wait = $True
                        }
                        if (![string]::IsNullOrEmpty($exeSplat.Parameter))
                        {
                            $exeSplatUser.Add("Parameters", $exeSplat.Parameter)
                        }

                        # add username
                        if ($loggedOnUsers.Count -eq 1)
                        {
                            $exeSplatUser.Add("UserName", $loggedOnUsers.NTAccount)
                            Write-Log -Message "Uninstalling '$($exeSplat.Path)' with logged on user account '$($loggedOnUsers.NTAccount)'" -Severity 1 -Source "CustomFunctions.ps1" -ScriptSection "Start-CustomExe" -LogType CMTrace
                        }
                        elseif ($loggedOnUsers.Count -gt 1)
                        {
                            $exeSplatUser.Add("UserName", $loggedOnUsers[0].NTAccount)
                            Write-Log -Message "Uninstalling '$($exeSplat.Path)' with logged on user account -- using '$($loggedOnUsers[0].NTAccount)' from '$($loggedOnUsers.NTAccount -join ', ')'" -Severity 1 -Source "CustomFunctions.ps1" -ScriptSection "Start-CustomExe" -LogType CMTrace
                        }

                        # execute process as user
                        Execute-ProcessAsUser @exeSplatUser

                        if ($Obj.waitForProcess)
                        {
                            # wait for given process to finish before continuing (not sure if this will work when 'Path' is run in another user context)
                            Wait-ForProgramToFinish -ProcessName $Obj.waitForProcess -CalledFromCustom
                        }
                    }
                }
            }
        }
    }
}

# Wait for install/uninstall to be finished
Function Wait-ForProgramToFinish
{
    ## ExtensionAppDeployScript
    param(
        [Parameter(Mandatory = $True)]
        [string]$ProcessName,

        [Parameter()]
        [switch]$CalledFromCustom
    )

    # get filename only
    $procName = [System.IO.Path]::GetFileNameWithoutExtension($ProcessName)
    
    # wait for (un)install to be completed
    if (!$CalledFromCustom)
    {
        Write-Log -Message "Waiting for $($DeploymentType.ToLower())($procName) to be completed" -Severity 1 -Source "CustomFunctions.ps1" -ScriptSection "Wait-ForProgramToFinish\$($DeploymentType)ation" -LogType CMTrace
    }
    else
    {
        Write-Log -Message "Waiting for '$procName' to be completed" -Severity 1 -Source "CustomFunctions.ps1" -ScriptSection "Wait-ForProgramToFinish\$($DeploymentType)ation" -LogType CMTrace
    }

    while ((Get-Process -Name $procName))
    {
        Start-Sleep -Milliseconds 150
    }

    if (!$CalledFromCustom)
    {
        Write-Log -Message "$DeploymentType completed" -Severity 1 -Source "CustomFunctions.ps1" -ScriptSection "Wait-ForProgramToFinish\$($DeploymentType)ation" -LogType CMTrace
    }
    else
    {
        Write-Log -Message "'$procName' completed" -Severity 1 -Source "CustomFunctions.ps1" -ScriptSection "Wait-ForProgramToFinish\$($DeploymentType)ation" -LogType CMTrace
    }
}

# Wait x seconds for install/unintall to be finished
Function Wait-TimeForProgramToFinish
{
    ## ExtensionAppDeployScript
    param(
        [Parameter()]
        [int]$Minutes = 1
    )

    # wait for (un)install to be completed
    Write-Log -Message "Waiting $Minutes minutes for $($DeploymentType.ToLower()) to be completed" -Severity 1 -Source "CustomFunctions.ps1" -ScriptSection "Wait-TimeForProgramToFinish\$($DeploymentType)ation" -LogType CMTrace
    Start-Sleep -Seconds ($Minutes*60)
    Write-Log -Message "Wait for $($DeploymentType.ToLower()) completed" -Severity 1 -Source "CustomFunctions.ps1" -ScriptSection "Wait-TimeForProgramToFinish\$($DeploymentType)ation" -LogType CMTrace
}

# Install certificate as a trusted publisher
Function Start-CertificateInstallation
{
    ## ExtensionAppDeployScript
    param(
        [Parameter(Mandatory = $True)]
        [string]$Path,

        [Parameter(Mandatory = $True)]
        [switch]$HaltScriptOnError
    )

    # get filename
    $fileName = [System.IO.Path]::GetFileName($Path)

    # import Path
    Write-Log -Message "Adding trusted certificate $fileName" -Severity 1 -Source "CustomFunctions.ps1" -ScriptSection "Start-CertificateInstallation" -LogType CMTrace
    try
    {
        Execute-Process -Path "C:\Windows\System32\certutil.exe" -Parameters "-addstore TrustedPublisher `"$Path`"" -ErrorAction Stop
        Write-Log -Message "Added certificate $fileName" -Severity 1 -Source "CustomFunctions.ps1" -ScriptSection "Start-CertificateInstallation" -LogType CMTrace
    }
    catch
    {
        Write-Log -Message "Failed to add certificate '$Path' : $_" -Severity 3 -Source "CustomFunctions.ps1" -ScriptSection "Start-CertificateInstallation" -LogType CMTrace

        if ($HaltScriptOnError)
        {
            Write-Log -Message "HaltScriptOnError is given. Aborting installation..." -Severity 2 -Source "CustomFunctions.ps1" -ScriptSection "Start-CertificateInstallation" -LogType CMTrace
            Exit-Script -ExitCode 60001
        }
    }
}

# Find VC++ redist to uninstall
Function Find-VCUninstall
{
    ## ExtensionAppDeployScript
    param(
        [Parameter(Mandatory = $True)]
        [string]$Path,

        [Parameter()]
        [string]$DisplayNameSearchString,

        [Parameter()]
        [string]$DisplayVersionSearchString
    )

    $hkey = "HKEY_LOCAL_MACHINE"
    $hkeyReplacement = "HKLM:"

    if ($Path.StartsWith($hkey))
    {
        $Path = $Path.Replace($hkey, $hkeyReplacement)
    }

    $uninstItems = @()

    Get-ChildItem -Path $Path | % {
        $subPath = $_.Name.Replace($hkey, $hkeyReplacement)

        $displayName = Get-ItemProperty -Path $subPath -Name DisplayName -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Select -ExpandProperty DisplayName
        if ($displayName -and $DisplayNameSearchString)
        {
            $displayName = $displayName | Where { $_ -match $DisplayNameSearchString }
        }

        $displayVersion = Get-ItemProperty -Path $subPath -Name DisplayVersion -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Select -ExpandProperty DisplayVersion
        if ($displayVersion -and $DisplayVersionSearchString)
        {
            $displayVersion = $displayVersion | Where { $_ -match $DisplayVersionSearchString }
        }

        $uninstallString = Get-ItemProperty -Path $subPath -Name UninstallString -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Select -ExpandProperty UninstallString
        $quietUninstallString = Get-ItemProperty -Path $subPath -Name QuietUninstallString -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Select -ExpandProperty QuietUninstallString

        if ($displayName)
        {
            if ($quietUninstallString)
            {
                $uninstItems += (New-Object PSObject -Property @{ DisplayName = $displayName; DisplayVersion = $displayVersion; Uninstall = $quietUninstallString })
            }
            elseif ($uninstallString)
            {
                $uninstItems += (New-Object PSObject -Property @{ DisplayName = $displayName; DisplayVersion = $displayVersion; Uninstall = $uninstallString })
            }
        }
    }

    return $uninstItems
}

# Uninstall VC++ redist item found in Find-VCUninstall
Function Uninstall-VCRedistributable
{
    ## ExtensionAppDeployScript
    param(
        [Parameter(Mandatory = $True, ValueFromPipeline = $True)]
        $UninstallItem
    )

    Process
    {
        $uninstallSplit = $UninstallItem.Uninstall.Split([string[]]"`"", [System.StringSplitOptions]::RemoveEmptyEntries)
        $executable = $uninstallSplit[0]
        $parameters = $uninstallSplit[1].Trim()

        Write-Log -Message "Starting uninstallation for '$($UninstallItem.DisplayName)' ($($UninstallItem.DisplayVersion))" -Severity 1 -Source "CustomFunctions.ps1" -ScriptSection "Uninstall-VCRedistributable" -LogType CMTrace
        Execute-Process -Path $executable -Parameters $parameters -WindowStyle Hidden -CreateNoWindow -WaitForMsiExec
    }
}

# uninstall VC++ 2015,2017,2019,2015_2019 x86_x64
Function Start-VCUninstall
{
    ## ExtensionAppDeployScript
    if ([string]::IsNullOrEmpty($CustomConfig.CustomExtensions.VisualStudioPlusPlusPath))
    {
        Write-Log -Message "'VisualStudioPlusPlusPath' is not defined or is set to an empty value in 'CustomFunctionsConfig.xml'" -Severity 3 -Source "CustomFunctions.ps1" -ScriptSection "Start-VCUninstall" -LogType CMTrace
        return $False
    }
    elseif (!(Test-Path -Path $CustomConfig.CustomExtensions.VisualStudioPlusPlusPath))
    {
        Write-Log -Message "'VisualStudioPlusPlusPath' defined in 'CustomFunctionsConfig.xml' does not exist or is not reachable!" -Severity 3 -Source "CustomFunctions.ps1" -ScriptSection "Start-VCUninstall" -LogType CMTrace
        return $False
    }

    $uninstallPaths = @("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall")
    $uninstallItems = @()

    # before uninstallation
    $uninstallPaths | % {
        $uninstallItems += Find-VCUninstall -Path $_ -DisplayNameSearchString "C\+\+ 2015.+?Redistributable|C\+\+ 2017.+?Redistributable|C\+\+ 2019.+?Redistributable"
    }

    # call uninstall on any VC++ found
    if ($uninstallItems)
    {
        $uninstallItems | Uninstall-VCRedistributable

        return $True
    }
    else
    {
        return $False
    }
}

# install VC++ 2015-2019 x86_x64
Function Start-VCInstall
{
    ## ExtensionAppDeployScript
    $folderPath = $CustomConfig.CustomExtensions.VisualStudioPlusPlusPath

    Get-ChildItem -Path $folderPath -Filter "*.exe" -File | % {
        $exeSplat = @{
            WindowStyle = "Hidden"
            Path = $_.FullName
            WaitForMsiExec = $True
            Parameter = "/quiet /install"
        }

        # get VC++ version
        $fileVersion = $_ | Select -ExpandProperty VersionInfo | Select -ExpandProperty FileVersion

        Write-Log -Message "Installing '$([System.IO.Path]::GetFileName($exeSplat.Path))' ($fileVersion)" -Severity 1 -Source "CustomFunctions.ps1" -ScriptSection "Start-VCInstall" -LogType CMTrace
        Execute-Process @exeSplat
    }
}

# Get MSI properties (Must be single quoted. Exmaple: "'ProductCode'")
Function Get-MsiProperty
{
    ## ExtensionAppDeployScript
	param(
        [Parameter(Mandatory = $True)]
        [string]$Path,
        
        [Parameter(Mandatory = $True, ParameterSetName = "PreDefined")]
        [ValidateSet("Manufacturer", "ProductCode", "ProductName", "ProductVersion", "UpgradeCode")]
        [string]$Property,
        
        [Parameter(Mandatory = $True, ParameterSetName = "Custom")]
        [ValidateScript({ $_ -like "'*'"})]
        [string]$CustomProperty
	)
	    
	function Get-Property($Object, $PropertyName, [object[]]$ArgumentList)
	{
		return $Object.GetType().InvokeMember($PropertyName, 'Public, Instance, GetProperty', $null, $Object, $ArgumentList)
	}
	 
	function Invoke-Method($Object, $MethodName, $ArgumentList)
	{
		return $Object.GetType().InvokeMember($MethodName, 'Public, Instance, InvokeMethod', $null, $Object, $ArgumentList)
	}
	 
	$ErrorActionPreference = 'Stop'
	Set-StrictMode -Version Latest
	 
	$msiOpenDatabaseModeReadOnly = 0
	$Installer = New-Object -ComObject WindowsInstaller.Installer
	 
	$Database = Invoke-Method $Installer OpenDatabase @($Path, $msiOpenDatabaseModeReadOnly)
    
    if ($Property) { $View = Invoke-Method $Database OpenView @("SELECT Value FROM Property WHERE Property='$Property'") }
    elseif ($CustomProperty) { $View = Invoke-Method $Database OpenView @("SELECT Value FROM Property WHERE Property=$CustomProperty") }
	 
	Invoke-Method $View Execute
	 
	$Record = Invoke-Method $View Fetch
	if ($Record)
	{
		Write-Output(Get-Property $Record StringData 1)
	}
	 
	Invoke-Method $View Close @( )
	Remove-Variable -Name Record, View, Database, Installer
	 
}