Function New-LogItem
{
    param(
        [Parameter(Mandatory = $True)]
        $LogItem
    )

    if ([string]::IsNullOrEmpty($settings.LogPath))
    {
        # logging is disabled
        return;
    }

    if ($settings.LogPath.ToLower().EndsWith(".xlsx"))
    {
        # logging to Excel
        if (!(Get-Module ImportExcel -ListAvailable))
        {
            # ImportExcel module not found. Logging disabled
            Write-Error "Module 'ImportExcel' not found. Logging disabled" -ErrorAction Continue
            return;
        }

        $LogItem | Export-Excel -Path $settings.LogPath -Append -AutoSize -TableName "Usage" -TableStyle Medium1
    }
    elseif ($settings.LogPath.ToLower().EndsWith(".csv"))
    {
        # logging to csv
        $LogItem | Export-Csv -Path $settings.LogPath -Encoding UTF8 -Append -Delimiter ';' -NoTypeInformation
    }
    else
    {
        # logging to regular file
        $LogItem | Out-String | Out-File -FilePath $settings.LogPath -Encoding utf8 -Append
    }
}

# remove any variables still existing
Clear-Variable -Name defaultVars -Force -Confirm:$False -ErrorAction SilentlyContinue -WarningAction Ignore

# default variables
$defaultVars = @{
    FolderName = "New PDT $(Get-Date -Format 'dd.MM.yy HH:mm:ss')"
    Arch = "x64"
    Author = $env:USERNAME
}

# settings
$settings = @{
    OpenPDT = @{
        Active = $True
        Fallback = $True
        Program = "code"
        ProgramDisplayName = "Visual Studio Code"
        FallbackProgram = "C:\Windows\system32\WindowsPowerShell\v1.0\powershell_ise.exe"
        FallbackProgramDisplayName = "PowerShell ISE"
    }
    OutputFolder = "$env:USERPROFILE\Desktop"
    ToolkitPath = "Local-Or-UNC-Path-To-Folder-Containing-Toolkit"
    LogPath = "Local-Or-UNC-Path-To-Log-File-As-.xlsx-,-.csv-or-anythingotherextension" # Set to a empty string to disable logging. If set to .xlsx, PowerShell module "ImportExcel" must be installed. If set to .csv, "Export-CSV" will be used. Otherwise normal file logging will be used
}

# get variables
$folderName = Read-Host -Prompt (Write-Host "Folder name:" -ForegroundColor Green)
$pdtVendor = Read-Host -Prompt (Write-Host "Application Vendor:" -ForegroundColor Green)
$pdtName = Read-Host -Prompt (Write-Host "Application Name:" -ForegroundColor Green)
$pdtVersion = Read-Host -Prompt (Write-Host "Application Version:" -ForegroundColor Green)
$pdtArch = Read-Host -Prompt (Write-Host "Application Architecture (x86|x64) (default '$($defaultVars.Arch)'):" -ForegroundColor Green)
$pdtAuthor = Read-Host -Prompt (Write-Host "Author (default '$($defaultVars.Author)'):" -ForegroundColor Green)
$pdtDate = (Get-Date -Format "dd/MM/yyyy").Replace(".", "/")

# set default folderName if none specified
if (!$folderName)
{
    $folderName = $defaultVars.FolderName
}

# set default arch if none specified
if (!$pdtArch)
{
    $pdtArch = $defaultVars.Arch
}

# set default author if none specified
if (!$pdtAuthor)
{
    $pdtAuthor = $defaultVars.Author
}

$outputFolder = Join-Path -Path $settings.OutputFolder -ChildPath $folderName
$outputPath = Join-Path -Path $outputFolder -ChildPath "Deploy-Application.ps1"

# copy folder to desktop
Copy-Item -Path $settings.ToolkitPath -Destination $outputFolder -Container -Force -Recurse -Confirm:$False

#$pdtScript = Get-Content -Path $outputPath -Raw -Encoding UTF8
$pdtScript = Get-Content -Path $outputPath -Raw

# replace vendor
if ($pdtVendor)
{
    $pdtScript = $pdtScript -replace "(\`$appVendor = '')|(\`$appVendor = '.+')","`$appVendor = '$pdtVendor'"
}

# replace name
if ($pdtName)
{
    $pdtScript = $pdtScript -replace "(\`$appName = '')|(\`$appName = '.+')","`$appName = '$pdtName'"
}

# replace version
if ($pdtVersion)
{
    $pdtScript = $pdtScript -replace "(\`$appVersion = '')|(\`$appVersion = '.+')","`$appVersion = '$pdtVersion'"
}

# replace arch
if ($pdtArch)
{
    $pdtScript = $pdtScript -replace "(\`$appArch = '')|(\`$appArch = '.+')","`$appArch = '$pdtArch'"
}

# replace author
if ($pdtAuthor)
{
    $pdtScript = $pdtScript -replace "(\`$appScriptAuthor = '')|(\`$appScriptAuthor = '.+')","`$appScriptAuthor = '$pdtAuthor'"
}

# replace date
if ($pdtDate)
{
    $pdtScript = $pdtScript -replace "(\`$appScriptDate = '')|(\`$appScriptDate = '.+')","`$appScriptDate = '$pdtDate'"
}

$pdtScript | Out-File -FilePath $outputPath -Encoding default -Force -NoNewline -Confirm:$False

# open PDT
if ($settings.OpenPDT.Active)
{
    $outputPath = "`"$outputPath`""

    # launching $settings.OpenPDT.ProgramDisplayName
    Write-Host "`nTrying to launch '$($settings.OpenPDT.ProgramDisplayName)' : " -ForegroundColor Cyan -NoNewline
    try
    {
        Start-Process -FilePath $settings.OpenPDT.Program -ArgumentList $outputPath -NoNewWindow -ErrorAction Stop
        Write-Host "Launched" -ForegroundColor Green
        New-LogItem -LogItem (New-Object PSObject -Property @{ Date = (Get-Date -Format 'dd.MM.yyyy HH:mm:ss'); User = $env:USERNAME; FolderName = $folderName; Vendor = $pdtVendor; Name = $pdtName; Version = $pdtVersion; Architecture = $pdtArch; Author = $pdtAuthor; Code = "Launched '$($settings.OpenPDT.ProgramDisplayName)'" } | Select Date, User, FolderName, Vendor, Name, Version, Architecture, Author, Code)
    }
    catch
    {
        $codeError = $_.ToString()
        Write-Host "Failed: $codeError" -ForegroundColor Yellow

        if ($settings.OpenPDT.Fallback)
        {
            # launching $settings.OpenPDT.FallbackProgramDisplayName
            Write-Host "`tTrying to launch '$($settings.OpenPDT.FallbackProgramDisplayName)' : " -ForegroundColor Cyan -NoNewline
            try
            {
                Start-Process -FilePath $settings.OpenPDT.FallbackProgram -ArgumentList $outputPath -ErrorAction Stop
                Write-Host "Launched" -ForegroundColor Green
                New-LogItem -LogItem (New-Object PSObject -Property @{ Date = (Get-Date -Format 'dd.MM.yyyy HH:mm:ss'); User = $env:USERNAME; FolderName = $folderName; Vendor = $pdtVendor; Name = $pdtName; Version = $pdtVersion; Architecture = $pdtArch; Author = $pdtAuthor; Code = "Launched '$($settings.OpenPDT.FallbackProgramDisplayName)'" } | Select Date, User, FolderName, Vendor, Name, Version, Architecture, Author, Code)
            }
            catch
            {
                New-LogItem -LogItem (New-Object PSObject -Property @{ Date = (Get-Date -Format 'dd.MM.yyyy HH:mm:ss'); User = $env:USERNAME; FolderName = $folderName; Vendor = $pdtVendor; Name = $pdtName; Version = $pdtVersion; Architecture = $pdtArch; Author = $pdtAuthor; Code = "Failed '$($settings.OpenPDT.ProgramDisplayName)': $codeError ::: Failed '$($settings.OpenPDT.FallbackProgramDisplayName)': $_" } | Select Date, User, FolderName, Vendor, Name, Version, Architecture, Author, Code)
                Write-Host "Failed: $_" -ForegroundColor Red
                Start-Sleep -Seconds 3
            }
        }
    }
}
else
{
    New-LogItem -LogItem (New-Object PSObject -Property @{ Date = (Get-Date -Format 'dd.MM.yyyy HH:mm:ss'); User = $env:USERNAME; FolderName = $folderName; Vendor = $pdtVendor; Name = $pdtName; Version = $pdtVersion; Architecture = $pdtArch; Author = $pdtAuthor; Code = "Launch deactivated" } | Select Date, User, FolderName, Vendor, Name, Version, Architecture, Author, Code)
}