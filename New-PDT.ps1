# remove any variables still existing
Clear-Variable -Name folderName -Force -Confirm:$False -ErrorAction SilentlyContinue -WarningAction Ignore
Clear-Variable -Name pdtVendor -Force -Confirm:$False -ErrorAction SilentlyContinue -WarningAction Ignore
Clear-Variable -Name pdtName -Force -Confirm:$False -ErrorAction SilentlyContinue -WarningAction Ignore
Clear-Variable -Name pdtVersion -Force -Confirm:$False -ErrorAction SilentlyContinue -WarningAction Ignore
Clear-Variable -Name pdtArch -Force -Confirm:$False -ErrorAction SilentlyContinue -WarningAction Ignore
Clear-Variable -Name pdtDate -Force -Confirm:$False -ErrorAction SilentlyContinue -WarningAction Ignore
Clear-Variable -Name pdtAuthor -Force -Confirm:$False -ErrorAction SilentlyContinue -WarningAction Ignore

# default variables
$defaultFolderName = "New PDT $(Get-Date -Format 'dd.MM.yy HH:mm:ss')"
$defaultArch = "x64"
$defaultAuthor = (Get-ADUser $env:USERNAME -Properties DisplayName | Select -ExpandProperty DisplayName)
$openPDT = $true
$openPDTFallback = $true
$openPDTProgramPriProgram = "code"
$openPDTProgramPriDisplayName = "Visual Studio Code"
$openPDTProgramSecProgram = "C:\Windows\system32\WindowsPowerShell\v1.0\powershell_ise.exe"
$openPDTProgramSecDisplayName = "PowerShell ISE"

# get variables
$folderName = Read-Host -Prompt (Write-Host "Folder name:" -ForegroundColor Green)
$pdtVendor = Read-Host -Prompt (Write-Host "Application Vendor:" -ForegroundColor Green)
$pdtName = Read-Host -Prompt (Write-Host "Application Name:" -ForegroundColor Green)
$pdtVersion = Read-Host -Prompt (Write-Host "Application Version:" -ForegroundColor Green)
$pdtArch = Read-Host -Prompt (Write-Host "Application Architecture (x86|x64) (default '$defaultArch'):" -ForegroundColor Green)
$pdtAuthor = Read-Host -Prompt (Write-Host "Author (default '$defaultAuthor'):" -ForegroundColor Green)
$pdtDate = (Get-Date -Format "dd/MM/yyyy").Replace(".", "/")

# set default folderName if none specified
if (!$folderName)
{
    $folderName = $defaultFolderName
}

# set default arch if none specified
if (!$pdtArch)
{
    $pdtArch = $defaultArch
}

# set default author if none specified
if (!$pdtAuthor)
{
    $pdtAuthor = $defaultAuthor
}

$outputFolder = "$env:USERPROFILE\Desktop\$folderName"
$outputPath = "$outputFolder\Deploy-Application.ps1"

# copy folder to desktop
Copy-Item -Path "D:\SCRIPTS\Common Windows Installer Wrapper" -Destination $outputFolder -Container -Force -Recurse -Confirm:$False

if ($pdtVendor -or $pdtName -or $pdtVersion -or $pdtArch -or $pdtAuthor -or $pdtDate)
{
    #$pdtScript = Get-Content -Path $outputPath -Raw -Encoding UTF8
    $pdtScript = Get-Content -Path $outputPath -Raw

    # replace vendor
    if ($pdtVendor)
    {
        $pdtScript = $pdtScript.Replace("`$appVendor = ''", "`$appVendor = '$pdtVendor'")
    }

    # replace name
    if ($pdtName)
    {
        $pdtScript = $pdtScript.Replace("`$appName = ''", "`$appName = '$pdtName'")
    }

    # replace version
    if ($pdtVersion)
    {
        $pdtScript = $pdtScript.Replace("`$appVersion = ''", "`$appVersion = '$pdtVersion'")
    }

    # replace arch
    if ($pdtArch)
    {
        $pdtScript = $pdtScript.Replace("`$appArch = ''", "`$appArch = '$pdtArch'")
    }

    # replace author
    if ($pdtAuthor)
    {
        $pdtScript = $pdtScript.Replace("`$appScriptAuthor = ''", "`$appScriptAuthor = '$pdtAuthor'")
    }

    # replace date
    if ($pdtDate)
    {
        $pdtScript = $pdtScript.Replace("`$appScriptDate = ''", "`$appScriptDate = '$pdtDate'")
    }

    $pdtScript | Out-File -FilePath $outputPath -Encoding default -Force -NoNewline -Confirm:$False

    # open PDT in code
    if ($openPDT)
    {
        $outputPath = "`"$outputPath`""

        # launching Visual Studio code
        Write-Host "`nTrying to launch $openPDTProgramPriDisplayName : " -ForegroundColor Cyan -NoNewline
        try
        {
            Start-Process -FilePath $openPDTProgramPriProgram -ArgumentList $outputPath -NoNewWindow -ErrorAction Stop
            Write-Host "Launched" -ForegroundColor Green
            New-Object PSObject -Property @{ Date = (Get-Date -Format 'dd.MM.yyyy HH:mm:ss'); User = $env:USERNAME; FolderName = $folderName; Vendor = $pdtVendor; Name = $pdtName; Version = $pdtVersion; Architecture = $pdtArch; Author = $pdtAuthor; Code = "Launched '$openPDTProgramPriDisplayName'" } | Select Date, User, FolderName, Vendor, Name, Version, Architecture, Author, Code | Export-Excel -Path "D:\SCRIPTS\Common Windows Installer Wrapper usage.xlsx" -Append -AutoSize -TableName "Usage" -TableStyle Medium1
        }
        catch
        {
            $codeError = $_.ToString()
            Write-Host "Failed: $codeError" -ForegroundColor Yellow

            if ($openPDTFallback)
            {
                # launching PowerShell ISE
                Write-Host "`tTrying to launch $openPDTProgramSecDisplayName : " -ForegroundColor Cyan -NoNewline
                try
                {
                    Start-Process -FilePath $openPDTProgramSecProgram -ArgumentList $outputPath -ErrorAction Stop
                    Write-Host "Launched" -ForegroundColor Green
                    New-Object PSObject -Property @{ Date = (Get-Date -Format 'dd.MM.yyyy HH:mm:ss'); User = $env:USERNAME; FolderName = $folderName; Vendor = $pdtVendor; Name = $pdtName; Version = $pdtVersion; Architecture = $pdtArch; Author = $pdtAuthor; Code = "Launched '$openPDTProgramSecDisplayName'" } | Select Date, User, FolderName, Vendor, Name, Version, Architecture, Author, Code | Export-Excel -Path "D:\SCRIPTS\Common Windows Installer Wrapper usage.xlsx" -Append -AutoSize -TableName "Usage" -TableStyle Medium1
                }
                catch
                {
                    New-Object PSObject -Property @{ Date = (Get-Date -Format 'dd.MM.yyyy HH:mm:ss'); User = $env:USERNAME; FolderName = $folderName; Vendor = $pdtVendor; Name = $pdtName; Version = $pdtVersion; Architecture = $pdtArch; Author = $pdtAuthor; Code = "Failed '$openPDTProgramPriDisplayName': $codeError - Failed '$openPDTProgramSecDisplayName': $_" } | Select Date, User, FolderName, Vendor, Name, Version, Architecture, Author, Code | Export-Excel -Path "D:\SCRIPTS\Common Windows Installer Wrapper usage.xlsx" -Append -AutoSize -TableName "Usage" -TableStyle Medium1
                    Write-Host "Failed: $_" -ForegroundColor Red
                    Start-Sleep -Seconds 3
                }
            }
        }
    }
    else
    {
        New-Object PSObject -Property @{ Date = (Get-Date -Format 'dd.MM.yyyy HH:mm:ss'); User = $env:USERNAME; FolderName = $folderName; Vendor = $pdtVendor; Name = $pdtName; Version = $pdtVersion; Architecture = $pdtArch; Author = $pdtAuthor; Code = "Launch deactivated" } | Select Date, User, FolderName, Vendor, Name, Version, Architecture, Author, Code | Export-Excel -Path "D:\SCRIPTS\Common Windows Installer Wrapper usage.xlsx" -Append -AutoSize -TableName "Usage" -TableStyle Medium1
    }
}