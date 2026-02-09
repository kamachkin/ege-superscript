# Windows installation and configuration automation superscript - MODIFIED VERSION
# Auto-detects drive with 'soft' folder
# Only connects to WiFi if no wired connection available
# No user prompts or confirmations

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8
$ErrorActionPreference = "Continue"

# Find drive with 'soft' folder
$script:softDir = $null
foreach ($drive in Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Free -gt 0 }) {
    $testPath = Join-Path $drive.Root "soft"
    if (Test-Path $testPath) {
        $script:softDir = $testPath
        break
    }
}

if (-not $script:softDir) {
    Write-Host "ERROR: Could not find 'soft' folder on any drive!" -ForegroundColor Red
    Exit 1
}

Write-Host "Found software directory: $script:softDir" -ForegroundColor Green

# Wi-Fi settings
$WiFiSSID = "34"
$WiFiPassword = "12315900"

# Counters
$script:softwareSuccess = 0
$script:softwareError = 0
$script:softwareSkipped = 0
$script:shortcutSuccess = 0
$script:shortcutError = 0
$script:shortcutSkipped = 0

# Logging
$logPath = "$env:TEMP\SuperScript_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$startTime = Get-Date

function Write-Log {
    param (
        [Parameter(Mandatory=$true)]
        [AllowEmptyString()]
        [string]$Message,
        [Parameter(Mandatory=$false)]
        [ValidateSet("INFO", "SUCCESS", "WARNING", "ERROR", "SKIP")]
        [string]$Level = "INFO"
    )
    
    if ([string]::IsNullOrWhiteSpace($Message)) { return }
    
    $timeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timeStamp] [$Level] $Message"
    
    $color = switch ($Level) {
        "ERROR"   { "Red" }
        "WARNING" { "Yellow" }
        "SUCCESS" { "Green" }
        "SKIP"    { "Cyan" }
        default   { "White" }
    }
    
    Write-Host $logMessage -ForegroundColor $color
    try {
        Add-Content -Path $logPath -Value $logMessage -ErrorAction SilentlyContinue
    } catch {}
}

function Write-Section {
    param([string]$Title)
    Write-Host "`n============================================================================" -ForegroundColor Cyan
    Write-Host "  $Title" -ForegroundColor Cyan
    Write-Host "============================================================================`n" -ForegroundColor Cyan
    Write-Log "=== $Title ===" -Level "INFO"
}

function Test-WiredConnection {
    try {
        $wiredAdapter = Get-NetAdapter | Where-Object { 
            $_.PhysicalMediaType -eq "802.3" -and 
            $_.Status -eq "Up" -and
            $_.MediaConnectionState -eq "Connected"
        }
        return ($null -ne $wiredAdapter)
    } catch {
        return $false
    }
}

function Test-SoftwareInstalled {
    param(
        [string]$DisplayName,
        [string]$FilePath = $null
    )
    
    if ($FilePath -and (Test-Path $FilePath)) {
        return $true
    }
    
    $registryPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    
    foreach ($path in $registryPaths) {
        try {
            $installed = Get-ItemProperty $path -ErrorAction SilentlyContinue | 
                         Where-Object { $_.DisplayName -like "*$DisplayName*" }
            if ($installed) {
                return $true
            }
        } catch {
            continue
        }
    }
    
    return $false
}

function Connect-WiFi {
    Write-Section "STAGE 1: NETWORK CONNECTION"
    
    # Check if wired connection exists
    if (Test-WiredConnection) {
        Write-Log "Wired connection detected - skipping WiFi connection" -Level "SUCCESS"
        
        $InternetTest = Test-Connection -ComputerName "8.8.8.8" -Count 2 -Quiet
        if ($InternetTest) {
            Write-Log "Internet access available via wired connection!" -Level "SUCCESS"
            return $true
        } else {
            Write-Log "Wired connection has no internet access" -Level "WARNING"
        }
    }
    
    Write-Log "No wired connection - attempting WiFi connection..."
    
    try {
        $WiFiAdapter = Get-NetAdapter | Where-Object { 
            $_.Name -like "*Wi-Fi*" -or 
            $_.Name -like "*Wireless*" -or 
            $_.Name -like "*WLAN*" -or 
            $_.PhysicalMediaType -eq "Native 802.11" -or
            $_.PhysicalMediaType -eq "802.11"
        } | Where-Object { $_.Status -ne "Disabled" } | Select-Object -First 1

        if (-not $WiFiAdapter) {
            Write-Log "Wi-Fi adapter not found" -Level "WARNING"
            return $false
        }

        Write-Log "Wi-Fi adapter found: $($WiFiAdapter.Name)" -Level "SUCCESS"

        if ($WiFiAdapter.Status -ne "Up") {
            Enable-NetAdapter -Name $WiFiAdapter.Name -Confirm:$false
            Start-Sleep -Seconds 3
        }

        $WiFiProfileXML = @"
<?xml version="1.0"?>
<WLANProfile xmlns="http://www.microsoft.com/networking/WLAN/profile/v1">
    <name>$WiFiSSID</name>
    <SSIDConfig>
        <SSID>
            <name>$WiFiSSID</name>
        </SSID>
    </SSIDConfig>
    <connectionType>ESS</connectionType>
    <connectionMode>auto</connectionMode>
    <MSM>
        <security>
            <authEncryption>
                <authentication>WPA2PSK</authentication>
                <encryption>AES</encryption>
                <useOneX>false</useOneX>
            </authEncryption>
            <sharedKey>
                <keyType>passPhrase</keyType>
                <protected>false</protected>
                <keyMaterial>$WiFiPassword</keyMaterial>
            </sharedKey>
        </security>
    </MSM>
</WLANProfile>
"@

        $TempProfilePath = "$env:TEMP\WiFiProfile_$WiFiSSID.xml"
        $WiFiProfileXML | Out-File -FilePath $TempProfilePath -Encoding UTF8

        netsh wlan add profile filename="$TempProfilePath" user=all | Out-Null
        Remove-Item -Path $TempProfilePath -Force -ErrorAction SilentlyContinue

        netsh wlan connect name="$WiFiSSID" | Out-Null
        Start-Sleep -Seconds 5

        $InternetTest = Test-Connection -ComputerName "8.8.8.8" -Count 2 -Quiet
        if ($InternetTest) {
            Write-Log "Successfully connected to WiFi with internet access!" -Level "SUCCESS"
            return $true
        } else {
            Write-Log "WiFi connected but no internet access" -Level "WARNING"
            return $false
        }
    }
    catch {
        Write-Log "WiFi connection error: $_" -Level "ERROR"
        return $false
    }
}

function Install-NuGetProvider {
    Write-Section "STAGE 1.5: PREPARING ENVIRONMENT"
    
    $internetTest = Test-Connection -ComputerName "8.8.8.8" -Count 2 -Quiet
    if (-not $internetTest) {
        Write-Log "No internet connection - skipping NuGet installation" -Level "WARNING"
        return $false
    }
    
    try {
        $NuGet = Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue
        if (-not $NuGet) {
            Write-Log "Installing NuGet provider..."
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -ForceBootstrap -Confirm:$false -ErrorAction Stop | Out-Null
            Write-Log "NuGet provider installed" -Level "SUCCESS"
            return $true
        } else {
            Write-Log "NuGet provider already installed" -Level "SUCCESS"
            return $true
        }
    } catch {
        Write-Log "Could not install NuGet provider: $_" -Level "WARNING"
        return $false
    }
}

function Install-Software {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Name,
        [Parameter(Mandatory=$true)]
        [string]$FilePath,
        [Parameter(Mandatory=$true)]
        [string]$Arguments,
        [Parameter(Mandatory=$false)]
        [string]$CheckPath = $null,
        [Parameter(Mandatory=$false)]
        [string]$CheckName = $null
    )
    
    $checkDisplayName = if ($CheckName) { $CheckName } else { $Name }
    
    if (Test-SoftwareInstalled -DisplayName $checkDisplayName -FilePath $CheckPath) {
        Write-Log "[SKIP] $Name (already installed)" -Level "SKIP"
        $script:softwareSkipped++
        return
    }
    
    $installStartTime = Get-Date
    Write-Log "Installing: $Name"
    
    try {
        if (-not (Test-Path $FilePath)) {
            Write-Log "File not found: $FilePath" -Level "WARNING"
            $script:softwareError++
            return
        }
        
        $process = Start-Process -FilePath $FilePath -ArgumentList $Arguments -Wait -PassThru -NoNewWindow
        
        if ($process.ExitCode -eq 0 -or $null -eq $process.ExitCode) {
            $duration = ((Get-Date) - $installStartTime).TotalSeconds
            Write-Log "[OK] $Name installed ($($duration.ToString("F1")) sec.)" -Level "SUCCESS"
            $script:softwareSuccess++
        } else {
            Write-Log "[ERROR] Installation error $Name (exit code: $($process.ExitCode))" -Level "ERROR"
            $script:softwareError++
        }
    }
    catch {
        Write-Log "[ERROR] Installation error $Name : $_" -Level "ERROR"
        $script:softwareError++
    }
}

function Start-SoftwareInstallation {
    Write-Section "STAGE 2: SOFTWARE INSTALLATION"
    
    Write-Log "Starting installation of programs..."
    Write-Log "Software directory: $script:softDir"
    
    # Code editors and IDEs
    Write-Host "`n--- Code Editors and IDEs ---" -ForegroundColor Yellow
    
    Install-Software -Name "WingIDE" `
        -FilePath "$script:softDir\wing.exe" `
        -Arguments "/VERYSILENT /SUPPRESSMSBOXES /NORESTART /SP-" `
        -CheckPath "${env:ProgramFiles(x86)}\Wing 101 8\bin\wing-101.exe" `
        -CheckName "Wing"
    
    Install-Software -Name "Sublime Text" `
        -FilePath "$script:softDir\sublime_text_build_4169_x64_setup.exe" `
        -Arguments "/VERYSILENT /SUPPRESSMSBOXES /NORESTART /SP-" `
        -CheckPath "$env:ProgramFiles\Sublime Text\sublime_text.exe" `
        -CheckName "Sublime Text"
    
    Install-Software -Name "Visual Studio Code" `
        -FilePath "$script:softDir\VSCodeUserSetup-x64-1.89.1.exe" `
        -Arguments "/VERYSILENT /NORESTART /MERGETASKS=!runcode" `
        -CheckPath "$env:LOCALAPPDATA\Programs\Microsoft VS Code\Code.exe" `
        -CheckName "Microsoft Visual Studio Code"
    
    Install-Software -Name "PyCharm Community" `
        -FilePath "$script:softDir\pycharm-community-2025.1.1.1.exe" `
        -Arguments "/S" `
        -CheckPath "${env:ProgramFiles(x86)}\JetBrains\PyCharm Community Edition 2025.1.1.1\bin\pycharm64.exe" `
        -CheckName "PyCharm"
    
    Install-Software -Name "IntelliJ IDEA Community" `
        -FilePath "$script:softDir\ideaIC-2025.1.1.1.exe" `
        -Arguments "/S" `
        -CheckPath "${env:ProgramFiles(x86)}\JetBrains\IntelliJ IDEA Community Edition 2025.1.1.1\bin\idea64.exe" `
        -CheckName "IntelliJ IDEA"
    
    Install-Software -Name "Code::Blocks" `
        -FilePath "$script:softDir\codeblocks-20.03mingw-setup.exe" `
        -Arguments "/S" `
        -CheckPath "$env:ProgramFiles\CodeBlocks\codeblocks.exe" `
        -CheckName "CodeBlocks"
    
    Install-Software -Name "PascalABC.NET" `
        -FilePath "$script:softDir\PascalABCNETSetup.exe" `
        -Arguments "/S" `
        -CheckPath "${env:ProgramFiles(x86)}\PascalABC.NET\PascalABCNET.exe" `
        -CheckName "PascalABC"
    
    if (Test-Path "$script:softDir\kumir2-2.1.0-rc11-install.exe") {
        Install-Software -Name "KuMir" `
            -FilePath "$script:softDir\kumir2-2.1.0-rc11-install.exe" `
            -Arguments "/S" `
            -CheckPath "${env:ProgramFiles(x86)}\Kumir-2.1.0-rc11\bin\kumir2-classic.exe" `
            -CheckName "Kumir"
    } elseif (Test-Path "$script:softDir\kumir-setup.exe") {
        Install-Software -Name "KuMir" `
            -FilePath "$script:softDir\kumir-setup.exe" `
            -Arguments "/S" `
            -CheckPath "${env:ProgramFiles(x86)}\Kumir-2.1.0-rc11\bin\kumir2-classic.exe" `
            -CheckName "Kumir"
    }
    
    Install-Software -Name "Visual Studio 2022" `
        -FilePath "$script:softDir\vs2022\vs_setup.exe" `
        -Arguments "--quiet --norestart --wait --add Microsoft.VisualStudio.Workload.NetDesktop --includeRecommended" `
        -CheckPath "$env:ProgramFiles\Microsoft Visual Studio\2022\Community\Common7\IDE\devenv.exe" `
        -CheckName "Visual Studio Community 2022"
    
    # Programming languages
    Write-Host "`n--- Programming Languages ---" -ForegroundColor Yellow
    
    Install-Software -Name "Python 3.12.4" `
        -FilePath "$script:softDir\python-3.12.4-amd64.exe" `
        -Arguments "/quiet InstallAllUsers=1 PrependPath=1 Include_test=0" `
        -CheckPath "$env:ProgramFiles\Python312\python.exe" `
        -CheckName "Python 3.12"
    
    Install-Software -Name "Python 3.8.10" `
        -FilePath "$script:softDir\python-3.8.10-amd64.exe" `
        -Arguments "/quiet InstallAllUsers=1 PrependPath=1 Include_test=0" `
        -CheckPath "$env:ProgramFiles\Python38\python.exe" `
        -CheckName "Python 3.8"
    
    Install-Software -Name "JDK 11.0.21" `
        -FilePath "$script:softDir\jdk-11.0.21_windows-x64_bin.exe" `
        -Arguments "/s" `
        -CheckPath "$env:ProgramFiles\Java\jdk-11\bin\java.exe" `
        -CheckName "Java SE Development Kit"
    
    Install-Software -Name "OpenJDK 21" `
        -FilePath "msiexec.exe" `
        -Arguments "/i `"$script:softDir\OpenJDK21U-jdk_x64_windows_hotspot_21.0.7_6.msi`" /qn" `
        -CheckName "Eclipse Temurin JDK"
    
    # Utilities
    Write-Host "`n--- Utilities ---" -ForegroundColor Yellow
    
    Install-Software -Name "Total Commander" `
        -FilePath "$script:softDir\tcmd1103x64.exe" `
        -Arguments "/AHMGDU" `
        -CheckPath "$env:ProgramFiles\totalcmd\TOTALCMD64.EXE" `
        -CheckName "Total Commander"
    
    Install-Software -Name "7-Zip" `
        -FilePath "$script:softDir\7z2406-x64.exe" `
        -Arguments "/S" `
        -CheckPath "$env:ProgramFiles\7-Zip\7zFM.exe" `
        -CheckName "7-Zip"
    
    Install-Software -Name "Far Manager" `
        -FilePath "msiexec.exe" `
        -Arguments "/i `"$script:softDir\Far30b6300.x86.20240407.msi`" /quiet /norestart" `
        -CheckPath "${env:ProgramFiles(x86)}\Far Manager\Far.exe" `
        -CheckName "Far Manager"
    
    # Office applications
    Write-Host "`n--- Office Applications ---" -ForegroundColor Yellow
    
    Install-Software -Name "LibreOffice" `
        -FilePath "msiexec.exe" `
        -Arguments "/i `"$script:softDir\LibreOffice_7.6.6_Win_x86-64.msi`" /qn" `
        -CheckPath "$env:ProgramFiles\LibreOffice\program\swriter.exe" `
        -CheckName "LibreOffice"
    
    # Microsoft Office
    if (Test-Path "$script:softDir\MicrosoftOffice2019\Setup.exe") {
        $officeInstalled = Test-SoftwareInstalled -DisplayName "Microsoft Office" `
            -FilePath "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\WINWORD.EXE"
        
        if ($officeInstalled) {
            Write-Log "[SKIP] Microsoft Office 2019 (already installed)" -Level "SKIP"
            $script:softwareSkipped++
        } else {
            Write-Log "Installing: Microsoft Office 2019"
            try {
                Push-Location "$script:softDir\MicrosoftOffice2019"
                $process = Start-Process -FilePath ".\Setup.exe" -ArgumentList "/configure .\configuration.xml" -Wait -PassThru -NoNewWindow
                Pop-Location
                if ($process.ExitCode -eq 0) {
                    Write-Log "[OK] Microsoft Office 2019 installed" -Level "SUCCESS"
                    $script:softwareSuccess++
                } else {
                    Write-Log "[ERROR] Microsoft Office 2019 error" -Level "ERROR"
                    $script:softwareError++
                }
            } catch {
                Pop-Location
                Write-Log "[ERROR] Microsoft Office 2019 error: $_" -Level "ERROR"
                $script:softwareError++
            }
        }
    }
    
    # Eclipse (extraction)
    Write-Host "`n--- Extracting Eclipse ---" -ForegroundColor Yellow
    
    if (Test-Path "C:\eclipse\eclipse\eclipse.exe") {
        Write-Log "[SKIP] Eclipse IDE (already extracted)" -Level "SKIP"
        $script:softwareSkipped++
    } elseif (Test-Path "$script:softDir\eclipse-java-2020-09-R-win32-x86_64.zip") {
        Write-Log "Installing: Eclipse IDE"
        try {
            Expand-Archive -Path "$script:softDir\eclipse-java-2020-09-R-win32-x86_64.zip" -DestinationPath "C:\eclipse" -Force
            Write-Log "[OK] Eclipse IDE extracted" -Level "SUCCESS"
            $script:softwareSuccess++
        } catch {
            Write-Log "[ERROR] Eclipse extraction error: $_" -Level "ERROR"
            $script:softwareError++
        }
    }
    
    Write-Log "Software installation completed. Success: $script:softwareSuccess | Skipped: $script:softwareSkipped | Errors: $script:softwareError"
}

function Start-WindowsUpdate {
    Write-Section "STAGE 3: WINDOWS UPDATE"
    
    try {
        $internetTest = Test-Connection -ComputerName "8.8.8.8" -Count 2 -Quiet
        if (-not $internetTest) {
            Write-Log "No internet connection, skipping updates" -Level "WARNING"
            return
        }
        
        $NuGet = Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue
        if (-not $NuGet) {
            Write-Log "NuGet provider not installed, cannot continue with updates" -Level "ERROR"
            return
        }
        
        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue
        
        $Module = Get-Module -ListAvailable -Name PSWindowsUpdate
        if (-not $Module) {
            Write-Log "Installing PSWindowsUpdate..."
            Install-Module -Name PSWindowsUpdate -Force -AllowClobber -Confirm:$false -ErrorAction Stop
        }
        
        Import-Module PSWindowsUpdate -Force -ErrorAction Stop
        
        # Multiple cycles to ensure all updates are installed
        $MaxCycles = 5
        $CurrentCycle = 0
        $TotalInstalled = 0
        
        Write-Log "Starting Windows Update process (max $MaxCycles cycles)..."
        
        do {
            $CurrentCycle++
            Write-Log "=== Update Cycle $CurrentCycle of $MaxCycles ===" -Level "INFO"
            
            try {
                # Search for updates
                Write-Log "Searching for available updates..."
                $Updates = Get-WindowsUpdate -ErrorAction Stop
                
                if ($Updates.Count -eq 0) {
                    Write-Log "No more updates available." -Level "SUCCESS"
                    break
                }
                
                Write-Log "Found $($Updates.Count) updates in this cycle"
                
                # Install updates without reboot
                Write-Log "Installing updates..."
                $InstallResults = Install-WindowsUpdate -AcceptAll -Install -IgnoreReboot -Confirm:$false -ErrorAction Stop
                
                $SuccessCount = ($InstallResults | Where-Object { $_.Result -eq "Installed" -or $_.Result -eq "Downloaded" }).Count
                $TotalInstalled += $SuccessCount
                
                Write-Log "Installed $SuccessCount updates in this cycle" -Level "SUCCESS"
                
                # Wait a bit before next cycle
                if ($CurrentCycle -lt $MaxCycles) {
                    Write-Log "Waiting 10 seconds before next cycle..."
                    Start-Sleep -Seconds 10
                }
            }
            catch {
                Write-Log "Error in cycle $CurrentCycle : $_" -Level "WARNING"
                # Continue to next cycle even if there's an error
                Start-Sleep -Seconds 5
            }
            
        } while ($CurrentCycle -lt $MaxCycles)
        
        Write-Log "Windows Update completed! Total updates installed: $TotalInstalled" -Level "SUCCESS"
        Write-Log "Note: Some updates may require a restart to complete installation" -Level "INFO"
    }
    catch {
        Write-Log "Critical update error: $_" -Level "ERROR"
    }
}

function Install-VSCodePythonExtension {
    Write-Section "STAGE 4: INSTALLING PYTHON EXTENSION FOR VS CODE"
    
    try {
        $codePath = "$env:LOCALAPPDATA\Programs\Microsoft VS Code\bin\code.cmd"
        
        if (-not (Test-Path $codePath)) {
            Write-Log "VS Code not found" -Level "WARNING"
            return
        }
        
        Write-Log "Installing Python extension for VS Code..."
        $process = Start-Process -FilePath $codePath -ArgumentList "--install-extension", "ms-python.python", "--force" -Wait -PassThru -NoNewWindow
        
        if ($process.ExitCode -eq 0) {
            Write-Log "Python extension successfully installed!" -Level "SUCCESS"
        }
        
        $process = Start-Process -FilePath $codePath -ArgumentList "--install-extension", "ms-python.vscode-pylance", "--force" -Wait -PassThru -NoNewWindow
        
        if ($process.ExitCode -eq 0) {
            Write-Log "Pylance extension successfully installed!" -Level "SUCCESS"
        }
    }
    catch {
        Write-Log "Error installing VS Code extensions: $_" -Level "ERROR"
    }
}

function Create-Shortcut {
    param (
        [string]$TargetPath,
        [string]$ShortcutName,
        [string]$Description = ""
    )
    
    try {
        if ($TargetPath.Contains("%")) {
            $TargetPath = [System.Environment]::ExpandEnvironmentVariables($TargetPath)
        }
        
        if (-not (Test-Path -LiteralPath $TargetPath)) {
            Write-Log "[WARNING] File not found: $TargetPath" -Level "WARNING"
            return $false
        }
        
        $desktopPath = [Environment]::GetFolderPath("Desktop")
        $shortcutPath = Join-Path -Path $desktopPath -ChildPath "$ShortcutName.lnk"
        
        if (Test-Path $shortcutPath) {
            Write-Log "[SKIP] $ShortcutName (already exists)" -Level "SKIP"
            $script:shortcutSkipped++
            return $true
        }
        
        $WshShell = New-Object -ComObject WScript.Shell
        $Shortcut = $WshShell.CreateShortcut($shortcutPath)
        $Shortcut.TargetPath = $TargetPath
        $Shortcut.WorkingDirectory = Split-Path -Parent $TargetPath
        if ($Description) { $Shortcut.Description = $Description }
        $Shortcut.Save()
        
        Write-Log "[OK] Created: $ShortcutName" -Level "SUCCESS"
        $script:shortcutSuccess++
        return $true
    }
    catch {
        Write-Log "[ERROR] Error creating shortcut '$ShortcutName': $_" -Level "ERROR"
        $script:shortcutError++
        return $false
    }
}

function Create-AllShortcuts {
    Write-Section "STAGE 5: CREATING DESKTOP SHORTCUTS"
    
    $programFiles = $env:ProgramFiles
    $programFilesX86 = ${env:ProgramFiles(x86)}
    $localAppData = $env:LOCALAPPDATA
    
    Write-Host "--- File Managers ---" -ForegroundColor Yellow
    Create-Shortcut "$programFilesX86\Far Manager\Far.exe" "Far Manager 3"
    Create-Shortcut "$programFiles\totalcmd\TOTALCMD64.EXE" "Total Commander"
    
    Write-Host "`n--- Development Environments ---" -ForegroundColor Yellow
    Create-Shortcut "$programFiles\CodeBlocks\codeblocks.exe" "CodeBlocks"
    Create-Shortcut "$programFiles\Sublime Text\sublime_text.exe" "Sublime Text"
    Create-Shortcut "$programFiles\Microsoft Visual Studio\2022\Community\Common7\IDE\devenv.exe" "Visual Studio 2022"
    Create-Shortcut "$programFilesX86\Wing 101 8\bin\wing-101.exe" "Wing 101"
    Create-Shortcut "$programFilesX86\Kumir-2.1.0-rc11\bin\kumir2-classic.exe" "Kumir-Standard"
    Create-Shortcut "$programFiles\Python312\Lib\idlelib\idle.pyw" "IDLE (Python 3.12)"
    Create-Shortcut "$programFiles\Python38\Lib\idlelib\idle.pyw" "IDLE (Python 3.8)"
    Create-Shortcut "$programFilesX86\JetBrains\IntelliJ IDEA Community Edition 2025.1.1.1\bin\idea64.exe" "IntelliJ IDEA"
    Create-Shortcut "$programFilesX86\JetBrains\PyCharm Community Edition 2025.1.1.1\bin\pycharm64.exe" "PyCharm"
    Create-Shortcut "$programFilesX86\PascalABC.NET\PascalABCNET.exe" "PascalABC.NET"
    Create-Shortcut "$localAppData\Programs\Microsoft VS Code\Code.exe" "VS Code"
    Create-Shortcut "C:\eclipse\eclipse\eclipse.exe" "Eclipse"
    
    Write-Host "`n--- Office Applications ---" -ForegroundColor Yellow
    Create-Shortcut "$programFilesX86\Microsoft Office\root\Office16\EXCEL.EXE" "Excel"
    Create-Shortcut "$programFilesX86\Microsoft Office\root\Office16\WINWORD.EXE" "Word"
    Create-Shortcut "$programFiles\LibreOffice\program\scalc.exe" "LibreOffice Calc"
    Create-Shortcut "$programFiles\LibreOffice\program\swriter.exe" "LibreOffice Writer"
    
    Write-Host "`n--- Interpreters ---" -ForegroundColor Yellow
    Create-Shortcut "$programFiles\Python312\python.exe" "Python 3.12"
    Create-Shortcut "$programFiles\Python38\python.exe" "Python 3.8"
    
    Write-Host "`n--- Utilities ---" -ForegroundColor Yellow
    Create-Shortcut "$programFiles\7-Zip\7zFM.exe" "7-Zip"
    Create-Shortcut "$env:windir\system32\notepad.exe" "Notepad"
    Create-Shortcut "$programFiles\Windows NT\Accessories\wordpad.exe" "WordPad"
    
    Write-Log "Shortcut creation completed. Success: $script:shortcutSuccess | Skipped: $script:shortcutSkipped | Errors: $script:shortcutError"
}

function Start-WindowsActivation {
    Write-Section "STAGE 6: WINDOWS AND OFFICE ACTIVATION (FINAL STEP)"
    
    Write-Log "This is the FINAL step - running activation after all installations" -Level "INFO"
    
    $needsWindowsActivation = $false
    $needsOfficeActivation = $false
    
    # Check Windows activation
    Write-Log "Checking Windows activation status..."
    try {
        $activationStatus = Get-CimInstance -ClassName SoftwareLicensingProduct | 
                           Where-Object { $_.PartialProductKey -and $_.Name -like "*Windows*" }
        
        $windowsActivated = $activationStatus | Where-Object { $_.LicenseStatus -eq 1 }
        
        if (-not $windowsActivated) {
            Write-Log "Windows is NOT activated" -Level "WARNING"
            $needsWindowsActivation = $true
        } else {
            Write-Log "Windows is already activated" -Level "SUCCESS"
        }
    }
    catch {
        Write-Log "Could not check Windows activation, will attempt activation" -Level "WARNING"
        $needsWindowsActivation = $true
    }
    
    # Check Office activation
    Write-Log "Checking Office activation status..."
    $osppPath = "${env:ProgramFiles(x86)}\Microsoft Office\Office16\OSPP.VBS"
    
    if (Test-Path $osppPath) {
        try {
            $output = cscript //NoLogo "$osppPath" /dstatus 2>&1 | Out-String
            
            if ($output -match "LICENSE STATUS:\s*---LICENSED---") {
                Write-Log "Office is already activated" -Level "SUCCESS"
            } else {
                Write-Log "Office is NOT activated" -Level "WARNING"
                $needsOfficeActivation = $true
            }
        }
        catch {
            Write-Log "Could not check Office activation, will attempt activation" -Level "WARNING"
            $needsOfficeActivation = $true
        }
    } else {
        Write-Log "Office not found or OSPP.VBS missing" -Level "INFO"
    }
    
    # Run activator if needed
    if ($needsWindowsActivation -or $needsOfficeActivation) {
        Write-Log "" -Level "INFO"
        Write-Log "================================================" -Level "INFO"
        Write-Log "RUNNING ACTIVATION SCRIPT FROM https://get.activated.win" -Level "INFO"
        Write-Log "================================================" -Level "INFO"
        
        if ($needsWindowsActivation -and $needsOfficeActivation) {
            Write-Log "Will attempt to activate: Windows AND Office" -Level "INFO"
        } elseif ($needsWindowsActivation) {
            Write-Log "Will attempt to activate: Windows only" -Level "INFO"
        } else {
            Write-Log "Will attempt to activate: Office only" -Level "INFO"
        }
        
        try {
            Write-Log "Downloading and executing activation script..."
            Invoke-Expression (Invoke-RestMethod -Uri "https://get.activated.win")
            Write-Log "Activation script executed successfully!" -Level "SUCCESS"
            Write-Log "Please verify activation status manually after completion" -Level "INFO"
        }
        catch {
            Write-Log "Error running activation script: $_" -Level "ERROR"
            Write-Log "You can manually activate later using:" -Level "INFO"
            Write-Log "  Open PowerShell as Admin: irm https://get.activated.win | iex" -Level "INFO"
        }
    } else {
        Write-Log "================================================" -Level "SUCCESS"
        Write-Log "No activation needed - everything is already activated!" -Level "SUCCESS"
        Write-Log "================================================" -Level "SUCCESS"
    }
}

# MAIN EXECUTION
Clear-Host

Write-Host @"
============================================================================
                                                                             
                    INSTALLATION AUTOMATION SUPERSCRIPT                      
                       Auto-Unattended Version                          
                                                                             
============================================================================
"@ -ForegroundColor Cyan

Write-Log "========================================"
Write-Log "SUPERSCRIPT START"
Write-Log "========================================"
Write-Log "Software directory: $script:softDir"
Write-Log "Log file: $logPath"

try {
    # STAGE 1: Network Connection
    $wifiConnected = Connect-WiFi
    
    if (-not $wifiConnected) {
        Write-Log "Continuing without internet connection..." -Level "WARNING"
    }
    
    # STAGE 1.5: Prepare Environment
    $nugetInstalled = $false
    if ($wifiConnected) {
        $nugetInstalled = Install-NuGetProvider
    }
    
    # STAGE 2: Software Installation
    Start-SoftwareInstallation
    
    # STAGE 3: Windows Update (with multiple cycles)
    if ($wifiConnected -and $nugetInstalled) {
        Start-WindowsUpdate
    }
    
    # STAGE 4: VS Code Extensions
    if ($wifiConnected) {
        Install-VSCodePythonExtension
    }
    
    # STAGE 5: Desktop Shortcuts
    Create-AllShortcuts
    
    # STAGE 6: Windows & Office Activation (LAST!)
    if ($wifiConnected) {
        Write-Log "All tasks completed. Running activation as final step..." -Level "INFO"
        Start-WindowsActivation
    }
    
    $totalTime = ((Get-Date) - $startTime).TotalMinutes
    
    Write-Section "EXECUTION COMPLETED"
    Write-Host "Total execution time: $($totalTime.ToString("F2")) minutes" -ForegroundColor Green
    Write-Host ""
    Write-Host "Summary:" -ForegroundColor Cyan
    Write-Host "  Software installed: $script:softwareSuccess" -ForegroundColor Green
    Write-Host "  Software skipped: $script:softwareSkipped" -ForegroundColor Cyan
    Write-Host "  Software errors: $script:softwareError" -ForegroundColor $(if ($script:softwareError -gt 0) { "Red" } else { "Green" })
    Write-Host ""
    Write-Host "  Shortcuts created: $script:shortcutSuccess" -ForegroundColor Green
    Write-Host "  Shortcuts skipped: $script:shortcutSkipped" -ForegroundColor Cyan
    Write-Host "  Shortcut errors: $script:shortcutError" -ForegroundColor $(if ($script:shortcutError -gt 0) { "Red" } else { "Green" })
    Write-Host ""
    Write-Host "Log file: $logPath" -ForegroundColor Yellow
    Write-Host ""
    
    Write-Log "========================================"
    Write-Log "SUPERSCRIPT EXECUTION COMPLETED"
    Write-Log "Total time: $($totalTime.ToString("F2")) minutes"
    Write-Log "========================================"
}
catch {
    Write-Host ""
    Write-Host "CRITICAL ERROR OCCURRED!" -ForegroundColor Red
    Write-Host "Error: $_" -ForegroundColor Red
    Write-Log "[ERROR] CRITICAL ERROR: $_"
}

Start-Sleep -Seconds 5
