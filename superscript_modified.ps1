# Windows installation and configuration automation superscript - MODIFIED VERSION
# Auto-detects drive with 'soft' folder
# Only connects to WiFi if no wired connection available
# No user prompts or confirmations

try {
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    $OutputEncoding = [System.Text.Encoding]::UTF8
} catch {}
$ErrorActionPreference = "Continue"
$ConfirmPreference = 'None'

# Self-elevate to Administrator if not already elevated (single UAC prompt)
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process -FilePath 'powershell.exe' -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

# Find drive with 'soft' folder
$script:softDir = $null
foreach ($drive in Get-PSDrive -PSProvider FileSystem -ErrorAction SilentlyContinue | Where-Object { $_.Root -and (Test-Path $_.Root) }) {
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

# Load WiFi networks from wifi.txt (format: SSID;Password per line)
$script:WiFiNetworks = @()
$wifiFile = Join-Path (Split-Path $script:softDir -Parent) "wifi.txt"
if (-not (Test-Path $wifiFile)) {
    $wifiFile = Join-Path $script:softDir "wifi.txt"
}
if (Test-Path $wifiFile) {
    Get-Content $wifiFile -Encoding UTF8 | ForEach-Object {
        $line = $_.Trim()
        if ($line -and -not $line.StartsWith('#')) {
            $parts = $line.Split(';', 2)
            if ($parts.Count -eq 2 -and $parts[0].Trim()) {
                $script:WiFiNetworks += @{ SSID = $parts[0].Trim(); Password = $parts[1].Trim() }
            }
        }
    }
    Write-Host "Loaded $($script:WiFiNetworks.Count) WiFi network(s) from wifi.txt" -ForegroundColor Green
} else {
    Write-Host "wifi.txt not found - WiFi auto-connect disabled" -ForegroundColor Yellow
}

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

function Update-EnvironmentPath {
    $machinePath = [System.Environment]::GetEnvironmentVariable("Path", "Machine")
    $userPath = [System.Environment]::GetEnvironmentVariable("Path", "User")
    $env:Path = "$machinePath;$userPath"
}

function Disable-SleepAndHibernation {
    Write-Log "Disabling sleep, hibernation, screen lock and screensaver..."
    try {
        # Disable sleep on AC and DC
        powercfg /change standby-timeout-ac 0
        powercfg /change standby-timeout-dc 0
        # Disable hibernate
        powercfg /change hibernate-timeout-ac 0
        powercfg /change hibernate-timeout-dc 0
        powercfg /hibernate off
        # Disable monitor timeout
        powercfg /change monitor-timeout-ac 0
        powercfg /change monitor-timeout-dc 0

        # Disable lock screen on idle (Machine policy)
        $regPath = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Personalization'
        if (-not (Test-Path $regPath)) { New-Item -Path $regPath -Force | Out-Null }
        Set-ItemProperty -Path $regPath -Name 'NoLockScreen' -Value 1 -Type DWord -Force

        # Disable lock screen timeout (never lock)
        $regPath2 = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System'
        Set-ItemProperty -Path $regPath2 -Name 'InactivityTimeoutSecs' -Value 0 -Type DWord -Force

        # Disable screensaver for all users via registry
        $regPath3 = 'HKCU:\Control Panel\Desktop'
        Set-ItemProperty -Path $regPath3 -Name 'ScreenSaveActive' -Value '0' -Force
        Set-ItemProperty -Path $regPath3 -Name 'ScreenSaverIsSecure' -Value '0' -Force
        Set-ItemProperty -Path $regPath3 -Name 'ScreenSaveTimeOut' -Value '0' -Force
        Set-ItemProperty -Path $regPath3 -Name 'SCRNSAVE.EXE' -Value '' -Force

        # Also set via Group Policy registry keys
        $regPath4 = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Control Panel\Desktop'
        if (-not (Test-Path $regPath4)) { New-Item -Path $regPath4 -Force | Out-Null }
        Set-ItemProperty -Path $regPath4 -Name 'ScreenSaveActive' -Value '0' -Force
        Set-ItemProperty -Path $regPath4 -Name 'ScreenSaverIsSecure' -Value '0' -Force
        Set-ItemProperty -Path $regPath4 -Name 'ScreenSaveTimeOut' -Value '0' -Force

        # Disable "Require sign-in" after sleep
        powercfg /SETACVALUEINDEX SCHEME_CURRENT SUB_NONE CONSOLELOCK 0
        powercfg /SETDCVALUEINDEX SCHEME_CURRENT SUB_NONE CONSOLELOCK 0
        powercfg /SETACTIVE SCHEME_CURRENT

        Write-Log "Sleep, hibernation, screen lock and screensaver disabled" -Level "SUCCESS"
    } catch {
        Write-Log "Error disabling sleep: $_" -Level "WARNING"
    }
}

function Test-InternetAccess {
    try {
        $tcp = New-Object System.Net.Sockets.TcpClient
        $tcp.Connect('8.8.8.8', 53)
        $tcp.Close()
        return $true
    } catch {
        return $false
    }
}

function Connect-WiFi {
    Write-Section "STAGE 1: NETWORK CONNECTION"
    
    # Check if wired connection exists
    if (Test-WiredConnection) {
        Write-Log "Wired connection detected - skipping WiFi connection" -Level "SUCCESS"
        
        if (Test-InternetAccess) {
            Write-Log "Internet access available via wired connection!" -Level "SUCCESS"
            return $true
        } else {
            Write-Log "Wired connection has no internet access" -Level "WARNING"
        }
    }
    
    if ($script:WiFiNetworks.Count -eq 0) {
        Write-Log "No WiFi networks configured (wifi.txt missing or empty)" -Level "WARNING"
        Write-Log "Continuing without internet connection..." -Level "WARNING"
        return $false
    }
    
    Write-Log "No wired connection - attempting WiFi connection..."
    Write-Log "Found $($script:WiFiNetworks.Count) WiFi network(s) to try"
    
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

        # Try each WiFi network from the list
        foreach ($network in $script:WiFiNetworks) {
            $ssid = $network.SSID
            $password = $network.Password
            Write-Log "Trying WiFi network: $ssid ..."

            $WiFiProfileXML = @"
<?xml version="1.0"?>
<WLANProfile xmlns="http://www.microsoft.com/networking/WLAN/profile/v1">
    <name>$ssid</name>
    <SSIDConfig>
        <SSID>
            <name>$ssid</name>
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
                <keyMaterial>$password</keyMaterial>
            </sharedKey>
        </security>
    </MSM>
</WLANProfile>
"@

            $TempProfilePath = "$env:TEMP\WiFiProfile_$ssid.xml"
            $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
            [System.IO.File]::WriteAllText($TempProfilePath, $WiFiProfileXML, $utf8NoBom)

            netsh wlan add profile filename="$TempProfilePath" user=all | Out-Null
            Remove-Item -Path $TempProfilePath -Force -ErrorAction SilentlyContinue

            for ($retry = 1; $retry -le 3; $retry++) {
                netsh wlan connect name="$ssid" | Out-Null
                Start-Sleep -Seconds 5
                if (Test-InternetAccess) {
                    Write-Log "Connected to '$ssid' with internet access!" -Level "SUCCESS"
                    return $true
                }
                Write-Log "Attempt $retry for '$ssid' - no internet yet..." -Level "WARNING"
                Start-Sleep -Seconds 3
            }
            Write-Log "Network '$ssid' failed after 3 attempts" -Level "WARNING"
        }

        Write-Log "All WiFi networks failed" -Level "WARNING"
        Write-Log "Continuing without internet connection..." -Level "WARNING"
        return $false
    }
    catch {
        Write-Log "WiFi connection error: $_" -Level "ERROR"
        return $false
    }
}

function Install-NuGetProvider {
    Write-Section "STAGE 1.5: PREPARING ENVIRONMENT"
    
    $internetTest = Test-InternetAccess
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
        
        if ($process.ExitCode -eq 0 -or $process.ExitCode -eq 3010 -or $null -eq $process.ExitCode) {
            $duration = ((Get-Date) - $installStartTime).TotalSeconds
            $rebootNote = if ($process.ExitCode -eq 3010) { " (reboot required)" } else { "" }
            Write-Log "[OK] $Name installed ($($duration.ToString("F1")) sec.)$rebootNote" -Level "SUCCESS"
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
    
    Install-Software -Name "VSCodium" `
        -FilePath "$script:softDir\VSCodiumSetup-x64-1.112.01907.exe" `
        -Arguments "/VERYSILENT /NORESTART /MERGETASKS=!runcode,addcontextmenufiles,addcontextmenufolders,addtopath" `
        -CheckPath "$env:ProgramFiles\VSCodium\VSCodium.exe" `
        -CheckName "VSCodium"
    
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
                if ($process.ExitCode -eq 0 -or $process.ExitCode -eq 3010) {
                    Write-Log "[OK] Microsoft Office 2019 installed" -Level "SUCCESS"
                    $script:softwareSuccess++
                } else {
                    Write-Log "[ERROR] Microsoft Office 2019 error (exit code: $($process.ExitCode))" -Level "ERROR"
                    $script:softwareError++
                }
            } catch {
                Write-Log "[ERROR] Microsoft Office 2019 error: $_" -Level "ERROR"
                $script:softwareError++
            } finally {
                Pop-Location
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
        $internetTest = Test-InternetAccess
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

function Register-VSCodeExtensions {
    param([string]$ExtensionsDir)
    
    try {
        $jsonPath = Join-Path $ExtensionsDir 'extensions.json'
        $obsPath  = Join-Path $ExtensionsDir '.obsolete'
        
        # Load existing extensions.json or create empty array
        $extensions = @()
        if (Test-Path $jsonPath) {
            $extensions = @(Get-Content $jsonPath -Raw | ConvertFrom-Json)
        }
        
        # Clean .obsolete — remove ms-python entries
        if (Test-Path $obsPath) {
            try {
                $obs = Get-Content $obsPath -Raw | ConvertFrom-Json
                $props = $obs.PSObject.Properties | Where-Object { $_.Name -notlike 'ms-python.*' }
                if ($props) {
                    $newObs = [ordered]@{}
                    foreach ($p in $props) { $newObs[$p.Name] = $p.Value }
                    $newObs | ConvertTo-Json -Compress | Set-Content $obsPath -Encoding UTF8
                } else {
                    Remove-Item $obsPath -Force
                }
                Write-Log "Cleaned ms-python entries from .obsolete"
            } catch {
                Write-Log "Could not clean .obsolete: $_" -Level "WARNING"
            }
        }
        
        # Find ms-python extension folders on disk
        $extFolders = Get-ChildItem -Path $ExtensionsDir -Directory -Filter 'ms-python*' -ErrorAction SilentlyContinue
        if (-not $extFolders) { return }
        
        $now = [long]((Get-Date).ToUniversalTime() - [datetime]'1970-01-01').TotalMilliseconds
        
        foreach ($folder in $extFolders) {
            $pkgPath = Join-Path $folder.FullName 'package.json'
            if (-not (Test-Path $pkgPath)) { continue }
            
            $pkg = Get-Content $pkgPath -Raw | ConvertFrom-Json
            $extId = "$($pkg.publisher).$($pkg.name)"
            
            # Skip if already registered
            $existing = $extensions | Where-Object { $_.identifier.id -eq $extId }
            if ($existing) { continue }
            
            # Determine forward-slash path for VS Code location
            $absPath = $folder.FullName -replace '\\','/'
            if ($absPath -match '^([A-Za-z]):(.*)') {
                $absPath = "/$($Matches[1].ToLower()):$($Matches[2])"
            }
            
            $entry = @{
                identifier       = @{ id = $extId }
                version          = $pkg.version
                location         = @{ '$mid' = 1; path = $absPath; scheme = 'file' }
                relativeLocation = $folder.Name
                metadata         = @{
                    installedTimestamp    = $now
                    pinned               = $true
                    source               = 'vsix'
                    publisherDisplayName = 'Microsoft'
                    isPreReleaseVersion  = $false
                    hasPreReleaseVersion = $false
                }
            }
            $extensions += $entry
            Write-Log "Registered extension in extensions.json: $extId v$($pkg.version)"
        }
        
        $extensions | ConvertTo-Json -Depth 10 | Set-Content $jsonPath -Encoding UTF8
    } catch {
        Write-Log "Failed to register extensions in extensions.json: $_" -Level "WARNING"
    }
}

function Install-VSCodePythonExtension {
    Write-Section "STAGE 4: INSTALLING VSCODIUM EXTENSIONS (ALL USERS)"
    
    try {
        $codeExePath = "$env:ProgramFiles\VSCodium\VSCodium.exe"
        
        if (-not (Test-Path $codeExePath)) {
            Write-Log "VSCodium not found at $codeExePath" -Level "WARNING"
            return
        }
        
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        
        # List of VSIX files to install
        $vsixFiles = @(
            "ms-python.python-2026.4.0.vsix",
            "ms-python.debugpy-2025.18.0-win32-x64.vsix",
            "ms-python.vscode-python-envs-1.24.0.vsix"
        )
        
        $installedAny = $false
        
        foreach ($vsixName in $vsixFiles) {
            $vsixPath = "$script:softDir\$vsixName"
            if (-not (Test-Path $vsixPath)) {
                Write-Log "VSIX not found: $vsixName" -Level "WARNING"
                continue
            }
            
            Write-Log "Installing extension: $vsixName ..."
            
            try {
                $tempExtract = "$env:TEMP\vsix_extract_$([System.IO.Path]::GetRandomFileName())"
                [System.IO.Compression.ZipFile]::ExtractToDirectory($vsixPath, $tempExtract)
                
                $manifestPath = "$tempExtract\extension.vsixmanifest"
                if (Test-Path $manifestPath) {
                    [xml]$manifest = Get-Content $manifestPath
                    $identity = $manifest.PackageManifest.Metadata.Identity
                    $folderName = "$($identity.Publisher).$($identity.Id)-$($identity.Version)".ToLower()
                } else {
                    $folderName = [System.IO.Path]::GetFileNameWithoutExtension($vsixName)
                }
                
                $targetDir = "$env:USERPROFILE\.vscode-oss\extensions\$folderName"
                $extensionSrc = "$tempExtract\extension"
                
                if (Test-Path $extensionSrc) {
                    if (-not (Test-Path $targetDir)) { New-Item -Path $targetDir -ItemType Directory -Force | Out-Null }
                    Copy-Item -Path "$extensionSrc\*" -Destination $targetDir -Recurse -Force
                    Write-Log "Extension installed: $folderName" -Level "SUCCESS"
                    $installedAny = $true
                } else {
                    Write-Log "Extension folder not found in VSIX: $vsixName" -Level "ERROR"
                }
                
                Remove-Item -Path $tempExtract -Recurse -Force -ErrorAction SilentlyContinue
            } catch {
                Write-Log "Failed to install $vsixName : $_" -Level "ERROR"
            }
        }
        
        # Register extensions in extensions.json and clean .obsolete for current user
        if ($installedAny) {
            Register-VSCodeExtensions -ExtensionsDir "$env:USERPROFILE\.vscode-oss\extensions"
        }
        
        # Copy extensions to all user profiles (Default + existing users)
        $examExtDir = "$env:USERPROFILE\.vscode-oss\extensions"
        $extFolders = Get-ChildItem -Path $examExtDir -Directory -Filter "ms-python*" -ErrorAction SilentlyContinue
        
        if ($extFolders -and $extFolders.Count -gt 0) {
            $pythonExe = "$env:ProgramFiles\Python312\python.exe"
            $settingsJson = @{ 'python.defaultInterpreterPath' = $pythonExe } | ConvertTo-Json
            
            $profilePaths = @("C:\Users\Default")
            $skipProfiles = @($env:USERNAME, 'Public', 'Default', 'Default User', 'All Users')
            Get-ChildItem -Path "C:\Users" -Directory -ErrorAction SilentlyContinue | 
                Where-Object { $_.Name -notin $skipProfiles -and -not $_.Name.StartsWith('.') } |
                ForEach-Object { $profilePaths += $_.FullName }
            
            foreach ($profileDir in $profilePaths) {
                if (-not (Test-Path $profileDir)) { continue }
                $profileName = Split-Path $profileDir -Leaf
                
                try {
                    foreach ($ext in $extFolders) {
                        $targetExtDir = Join-Path $profileDir ".vscode-oss\extensions\$($ext.Name)"
                        if (-not (Test-Path $targetExtDir)) {
                            New-Item -Path $targetExtDir -ItemType Directory -Force | Out-Null
                            Copy-Item -Path "$($ext.FullName)\*" -Destination $targetExtDir -Recurse -Force
                        }
                    }
                    
                    # Register extensions in this profile's extensions.json
                    Register-VSCodeExtensions -ExtensionsDir (Join-Path $profileDir ".vscode-oss\extensions")
                    
                    $targetSettingsDir = Join-Path $profileDir "AppData\Roaming\VSCodium\User"
                    $targetSettingsPath = Join-Path $targetSettingsDir "settings.json"
                    if (-not (Test-Path $targetSettingsDir)) {
                        New-Item -Path $targetSettingsDir -ItemType Directory -Force | Out-Null
                    }
                    if (Test-Path $targetSettingsPath) {
                        $raw = Get-Content $targetSettingsPath -Raw
                        if ($raw -notmatch 'python\.defaultInterpreterPath') {
                            $obj = $raw | ConvertFrom-Json
                            $obj | Add-Member -NotePropertyName 'python.defaultInterpreterPath' -NotePropertyValue $pythonExe -Force
                            $obj | ConvertTo-Json -Depth 10 | Set-Content $targetSettingsPath -Encoding UTF8
                        }
                    } else {
                        Set-Content -Path $targetSettingsPath -Value $settingsJson -Encoding UTF8
                    }
                    Write-Log "Extensions + settings copied to profile: $profileName" -Level "SUCCESS"
                } catch {
                    Write-Log "Could not configure profile $profileName : $_" -Level "WARNING"
                }
            }
        } else {
            Write-Log "No ms-python extensions found in $examExtDir to copy" -Level "WARNING"
        }
        
        # Settings for current user
        try {
            $pythonExe = "$env:ProgramFiles\Python312\python.exe"
            $examSettingsDir = "$env:APPDATA\VSCodium\User"
            if (-not (Test-Path $examSettingsDir)) {
                New-Item -Path $examSettingsDir -ItemType Directory -Force | Out-Null
            }
            $examSettingsPath = "$examSettingsDir\settings.json"
            if (Test-Path $examSettingsPath) {
                $raw = Get-Content $examSettingsPath -Raw
                if ($raw -notmatch 'python\.defaultInterpreterPath') {
                    $obj = $raw | ConvertFrom-Json
                    $obj | Add-Member -NotePropertyName 'python.defaultInterpreterPath' -NotePropertyValue $pythonExe -Force
                    $obj | ConvertTo-Json -Depth 10 | Set-Content $examSettingsPath -Encoding UTF8
                }
            } else {
                @{ 'python.defaultInterpreterPath' = $pythonExe } | ConvertTo-Json | Set-Content $examSettingsPath -Encoding UTF8
            }
            Write-Log "VSCodium settings configured for current user" -Level "SUCCESS"
        } catch {
            Write-Log "Could not configure VSCodium settings: $_" -Level "WARNING"
        }
    }
    catch {
        Write-Log "Error installing VSCodium extensions: $_" -Level "ERROR"
    }
}

function Create-Shortcut {
    param (
        [string]$TargetPath,
        [string]$ShortcutName,
        [string]$Description = "",
        [string]$Arguments = ""
    )
    
    try {
        if ($TargetPath.Contains("%")) {
            $TargetPath = [System.Environment]::ExpandEnvironmentVariables($TargetPath)
        }
        
        if (-not (Test-Path -LiteralPath $TargetPath)) {
            Write-Log "[WARNING] File not found: $TargetPath" -Level "WARNING"
            return $false
        }
        
        $desktopPath = "$env:PUBLIC\Desktop"
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
        if ($Arguments) { $Shortcut.Arguments = $Arguments }
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
    Create-Shortcut "$programFiles\Python312\pythonw.exe" "IDLE (Python 3.12)" -Arguments "-m idlelib"
    Create-Shortcut "$programFiles\Python38\pythonw.exe" "IDLE (Python 3.8)" -Arguments "-m idlelib"
    Create-Shortcut "$programFilesX86\JetBrains\IntelliJ IDEA Community Edition 2025.1.1.1\bin\idea64.exe" "IntelliJ IDEA"
    Create-Shortcut "$programFilesX86\JetBrains\PyCharm Community Edition 2025.1.1.1\bin\pycharm64.exe" "PyCharm"
    Create-Shortcut "$programFilesX86\PascalABC.NET\PascalABCNET.exe" "PascalABC.NET"
    Create-Shortcut "$programFiles\VSCodium\VSCodium.exe" "VSCodium"
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
    
    # Search for OSPP.VBS in multiple possible locations
    $osppPaths = @(
        "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\OSPP.VBS",
        "$env:ProgramFiles\Microsoft Office\root\Office16\OSPP.VBS",
        "${env:ProgramFiles(x86)}\Microsoft Office\Office16\OSPP.VBS",
        "$env:ProgramFiles\Microsoft Office\Office16\OSPP.VBS"
    )
    $osppPath = $osppPaths | Where-Object { Test-Path $_ } | Select-Object -First 1
    
    # Also check if Office is installed at all (registry or executable)
    $officeInstalled = (Test-Path "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\WINWORD.EXE") -or
                       (Test-Path "$env:ProgramFiles\Microsoft Office\root\Office16\WINWORD.EXE") -or
                       (Test-SoftwareInstalled -DisplayName "Microsoft Office")
    
    if ($osppPath) {
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
    } elseif ($officeInstalled) {
        Write-Log "Office is installed but OSPP.VBS not found - will run activator" -Level "WARNING"
        $needsOfficeActivation = $true
    } else {
        Write-Log "Office not installed" -Level "INFO"
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
            Write-Log "Launching activation script in separate window..."
            Start-Process -FilePath 'powershell.exe' -ArgumentList '-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', 'irm https://get.activated.win | iex; Write-Host "`nActivation completed. Press any key to close..." -ForegroundColor Green; $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")' -Verb RunAs
            Write-Log "Activation window opened (interactive - will stay open)" -Level "SUCCESS"
        }
        catch {
            Write-Log "Error launching activation script: $_" -Level "ERROR"
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
    # Disable sleep/hibernation so setup is not interrupted
    Disable-SleepAndHibernation
    
    # STAGE 1: Network Connection
    $wifiConnected = Connect-WiFi
    
    if (-not $wifiConnected) {
        Write-Log "Continuing without internet connection..." -Level "WARNING"
    }
    
    # STAGE 2: Start NuGet + Windows Update in a SEPARATE VISIBLE WINDOW
    $updateProcess = $null
    if ($wifiConnected) {
        Write-Log "Starting Windows Update in separate window..." -Level "INFO"
        $updateScriptPath = "$env:TEMP\WindowsUpdate_Runner.ps1"
        $updateScriptContent = @'
$host.UI.RawUI.WindowTitle = "Windows Update"
$ErrorActionPreference = 'Continue'
$wuLog = "$env:TEMP\WindowsUpdate_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

function Write-Status {
    param([string]$Msg, [string]$Color = 'White')
    $ts = Get-Date -Format 'HH:mm:ss'
    $line = "[$ts] $Msg"
    Write-Host $line -ForegroundColor $Color
    Add-Content -Path $wuLog -Value $line -ErrorAction SilentlyContinue
}

Write-Host "============================================" -ForegroundColor Cyan
Write-Host "    WINDOWS UPDATE (built-in COM API)       " -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

try {
    $MaxCycles = 2
    $MaxRetries = 2
    $CurrentCycle = 0
    $TotalInstalled = 0

    function Wait-ForInternet {
        param([int]$TimeoutSec = 120)
        $end = (Get-Date).AddSeconds($TimeoutSec)
        while ((Get-Date) -lt $end) {
            try {
                $tcp = New-Object System.Net.Sockets.TcpClient
                $tcp.Connect('8.8.8.8', 53)
                $tcp.Close()
                return $true
            } catch {
                Write-Status "Waiting for internet connection..." Yellow
                Start-Sleep -Seconds 5
            }
        }
        return $false
    }

    do {
        $CurrentCycle++
        $retryCount = 0
        $cycleSuccess = $false

        while (-not $cycleSuccess -and $retryCount -lt $MaxRetries) {
            try {
                Write-Host "--------------------------------------------" -ForegroundColor DarkGray
                if ($retryCount -gt 0) {
                    Write-Status "Retry $retryCount - waiting for network..." Yellow
                    if (-not (Wait-ForInternet -TimeoutSec 120)) {
                        Write-Status "Network not available after 2 min, skipping" Red
                        break
                    }
                    Write-Status "Network restored! Resuming..." Green
                    Start-Sleep -Seconds 3
                }

                Write-Status "Cycle $CurrentCycle of $MaxCycles - Searching for updates..." Cyan

                $session = New-Object -ComObject Microsoft.Update.Session
                $searcher = $session.CreateUpdateSearcher()
                $searchResult = $searcher.Search("IsInstalled=0 and Type='Software'")

                if ($searchResult.Updates.Count -eq 0) {
                    Write-Status "No more updates available!" Green
                    $cycleSuccess = $true
                    $CurrentCycle = $MaxCycles  # exit outer loop
                    break
                }

                Write-Status "Found $($searchResult.Updates.Count) update(s):" Yellow
                $updatesToDownload = New-Object -ComObject Microsoft.Update.UpdateColl
                foreach ($update in $searchResult.Updates) {
                    Write-Host "  - $($update.Title)" -ForegroundColor White
                    if (-not $update.EulaAccepted) { $update.AcceptEula() }
                    $updatesToDownload.Add($update) | Out-Null
                }

                Write-Host ""
                Write-Status "Downloading updates..." Yellow
                $downloader = $session.CreateUpdateDownloader()
                $downloader.Updates = $updatesToDownload
                $downloadResult = $downloader.Download()

                $updatesToInstall = New-Object -ComObject Microsoft.Update.UpdateColl
                foreach ($update in $searchResult.Updates) {
                    if ($update.IsDownloaded) {
                        $updatesToInstall.Add($update) | Out-Null
                    }
                }

                if ($updatesToInstall.Count -eq 0) {
                    Write-Status "No updates downloaded - network may have dropped" Yellow
                    $retryCount++
                    continue
                }

                # Ensure required services are running before install
                foreach ($svcName in @('wuauserv','msiserver','TrustedInstaller','BITS')) {
                    $svc = Get-Service -Name $svcName -ErrorAction SilentlyContinue
                    if ($svc -and $svc.Status -ne 'Running') {
                        Write-Status "Starting service $svcName..." Yellow
                        Start-Service -Name $svcName -ErrorAction SilentlyContinue
                        Start-Sleep -Seconds 2
                    }
                }

                Write-Status "Installing $($updatesToInstall.Count) update(s)..." Yellow
                $comInstallOK = $false
                try {
                    $installer = New-Object -ComObject Microsoft.Update.Installer
                    $installer.Updates = $updatesToInstall
                    $installResult = $installer.Install()

                    $installed = 0
                    for ($i = 0; $i -lt $updatesToInstall.Count; $i++) {
                        $res = $installResult.GetUpdateResult($i)
                        if ($res.ResultCode -eq 2) { $installed++ }
                    }

                    $TotalInstalled += $installed
                    Write-Status "Installed $installed update(s) in this cycle" Green
                    $comInstallOK = $true

                    if ($installResult.RebootRequired) {
                        Write-Status "Some updates require a reboot" Yellow
                    }
                } catch {
                    Write-Status "COM Installer failed: $($_.Exception.Message)" Yellow
                    Write-Status "Falling back to UsoClient..." Yellow
                }

                if (-not $comInstallOK) {
                    # Fallback: trigger install via UsoClient (works on Win10 22H2+)
                    try {
                        Start-Process -FilePath 'UsoClient.exe' -ArgumentList 'StartInstall' -Wait -NoNewWindow -ErrorAction Stop
                        Write-Status "UsoClient StartInstall triggered, waiting 60s..." Yellow
                        Start-Sleep -Seconds 60
                        $TotalInstalled += $updatesToInstall.Count
                        Write-Status "UsoClient install completed (approximate count)" Green
                        $comInstallOK = $true
                    } catch {
                        Write-Status "UsoClient also failed: $($_.Exception.Message)" Red
                    }
                }

                if ($comInstallOK) {
                    $cycleSuccess = $true
                } else {
                    throw "Both COM and UsoClient install methods failed"
                }
            } catch {
                $retryCount++
                Write-Status "Error (attempt $retryCount/$MaxRetries): $($_.Exception.Message)" Red
                if ($retryCount -lt $MaxRetries) {
                    Write-Status "Retrying after service restart..." Yellow
                }
            }
        }

        if ($CurrentCycle -lt $MaxCycles -and $cycleSuccess) { Start-Sleep -Seconds 10 }
    } while ($CurrentCycle -lt $MaxCycles)

    Write-Host ""
    Write-Host "============================================" -ForegroundColor Green
    Write-Status "COMPLETED! Total updates installed: $TotalInstalled" Green
    Write-Host "============================================" -ForegroundColor Green
} catch {
    Write-Status "Critical error: $_" Red
}

Write-Host ""
Write-Status "Log saved to: $wuLog"
Write-Host "This window will close in 30 seconds..." -ForegroundColor DarkGray
Start-Sleep -Seconds 30
'@
        Set-Content -Path $updateScriptPath -Value $updateScriptContent -Encoding UTF8 -Force
        $updateProcess = Start-Process -FilePath 'powershell.exe' -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$updateScriptPath`"" -PassThru
        Write-Log "Windows Update window opened (PID: $($updateProcess.Id))" -Level "SUCCESS"
    }
    
    # STAGE 3: Software Installation (runs while updates download in background)
    Start-SoftwareInstallation
    
    # Refresh PATH so newly installed tools are available
    Update-EnvironmentPath
    
    # STAGE 4: VSCodium Extensions (local VSIX, no internet needed)
    Install-VSCodePythonExtension
    
    # STAGE 5: Desktop Shortcuts
    Create-AllShortcuts

    # === All USB interaction is now complete ===
    Write-Host ""
    Write-Host "============================================" -ForegroundColor Green
    Write-Host "  USB FLASH DRIVE IS NO LONGER NEEDED"      -ForegroundColor Green
    Write-Host "  You can safely remove it now."             -ForegroundColor Green
    Write-Host "============================================" -ForegroundColor Green
    Write-Host ""
    Write-Log "All files read from USB - flash drive can be removed" -Level "SUCCESS"
    
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

    # === Generate desktop report ===
    try {
        $reportPath = "$env:PUBLIC\Desktop\Installation_Report.txt"
        $report = @()
        $report += '============================================================'
        $report += '  OTCHET OB USTANOVKE'
        $reportDate = Get-Date -Format 'dd.MM.yyyy HH:mm:ss'
        $reportTime = $totalTime.ToString('F2')
        $report += "  Data: $reportDate"
        $report += "  Vremya vypolneniya: $reportTime minut"
        $report += '============================================================'
        $report += ''

        # --- Network ---
        $report += '--- SET (Network) ---'
        if ($wifiConnected) {
            $report += '  [OK] Internet: connected'
        } else {
            $report += '  [!]  Internet: NOT connected'
        }
        $report += ''

        # --- Software ---
        $report += '--- PROGRAMMY (Software) ---'
        $report += '  Installed: ' + $script:softwareSuccess
        $report += '  Skipped: ' + $script:softwareSkipped
        $report += '  Errors: ' + $script:softwareError
        $report += ''

        # Check each program
        $checks = @(
            @{ Name = 'WingIDE'; Path = "${env:ProgramFiles(x86)}\Wing IDE 101 9\bin\wing-101.exe" },
            @{ Name = 'Sublime Text'; Path = "$env:ProgramFiles\Sublime Text\sublime_text.exe" },
            @{ Name = 'VSCodium'; Path = "$env:ProgramFiles\VSCodium\VSCodium.exe" },
            @{ Name = 'PyCharm Community'; Path = "${env:ProgramFiles(x86)}\JetBrains\PyCharm Community*\bin\pycharm64.exe" },
            @{ Name = 'IntelliJ IDEA Community'; Path = "${env:ProgramFiles(x86)}\JetBrains\IntelliJ IDEA Community*\bin\idea64.exe" },
            @{ Name = 'Python 3.12'; Path = "$env:ProgramFiles\Python312\python.exe" },
            @{ Name = 'Git'; Path = "$env:ProgramFiles\Git\bin\git.exe" },
            @{ Name = 'Java (JDK)'; Path = "$env:ProgramFiles\Java\jdk*\bin\java.exe" },
            @{ Name = '7-Zip'; Path = "$env:ProgramFiles\7-Zip\7z.exe" },
            @{ Name = 'Firefox'; Path = "$env:ProgramFiles\Mozilla Firefox\firefox.exe" },
            @{ Name = 'Google Chrome'; Path = "$env:ProgramFiles\Google\Chrome\Application\chrome.exe" },
            @{ Name = 'LibreOffice'; Path = "$env:ProgramFiles\LibreOffice\program\soffice.exe" },
            @{ Name = 'Office (Word)'; Path = "$env:ProgramFiles\Microsoft Office\root\Office16\WINWORD.EXE" },
            @{ Name = 'GIMP'; Path = "$env:ProgramFiles\GIMP 2\bin\gimp-2.*.exe" },
            @{ Name = 'Notepad++'; Path = "$env:ProgramFiles\Notepad++\notepad++.exe" },
            @{ Name = 'VLC'; Path = "$env:ProgramFiles\VideoLAN\VLC\vlc.exe" }
        )
        foreach ($c in $checks) {
            $found = $false
            if ($c.Path -match '\*') {
                $found = $null -ne (Get-Item $c.Path -ErrorAction SilentlyContinue | Select-Object -First 1)
            } else {
                $found = Test-Path $c.Path
            }
            $status = if ($found) { '[OK]' } else { '[--]' }
            $report += '  ' + $status + ' ' + $c.Name
        }
        $report += ''

        # --- VSCodium Extensions ---
        $report += '--- VSCODIUM EXTENSIONS ---'
        $extDir = "$env:USERPROFILE\.vscode-oss\extensions"
        if (Test-Path $extDir) {
            $pythonExt = Get-ChildItem $extDir -Directory -Filter 'ms-python.python-*' -ErrorAction SilentlyContinue
            $debugpyExt = Get-ChildItem $extDir -Directory -Filter 'ms-python.debugpy-*' -ErrorAction SilentlyContinue
            $envsExt = Get-ChildItem $extDir -Directory -Filter 'ms-python.vscode-python-envs-*' -ErrorAction SilentlyContinue
            $report += '  ' + $(if($pythonExt){'[OK]'}else{'[--]'}) + ' ms-python.python'
            $report += '  ' + $(if($debugpyExt){'[OK]'}else{'[--]'}) + ' ms-python.debugpy'
            $report += '  ' + $(if($envsExt){'[OK]'}else{'[--]'}) + ' ms-python.vscode-python-envs'
        } else {
            $report += '  [--] Extensions folder not found'
        }
        $report += ''

        # --- Shortcuts ---
        $report += '--- DESKTOP SHORTCUTS ---'
        $report += '  Created: ' + $script:shortcutSuccess
        $report += '  Skipped: ' + $script:shortcutSkipped
        $report += '  Errors: ' + $script:shortcutError
        $desktopShortcuts = Get-ChildItem "$env:PUBLIC\Desktop\*.lnk" -ErrorAction SilentlyContinue
        if ($desktopShortcuts) {
            foreach ($s in $desktopShortcuts) {
                $report += '  [OK] ' + $s.BaseName
            }
        }
        $report += ''

        # --- Activation ---
        $report += '--- ACTIVATION ---'
        $licStatus = cscript //nologo "$env:SystemRoot\System32\slmgr.vbs" /xpr 2>&1 | Out-String
        if ($licStatus -match 'permanently|activated') {
            $report += '  [OK] Windows activated'
        } else {
            $report += '  [!]  Windows: not activated'
        }
        # Office
        $officeOspp = @(
            "$env:ProgramFiles\Microsoft Office\root\vfs\ProgramFilesCommonX64\Microsoft Shared\Office16\OSPP.VBS",
            "$env:ProgramFiles\Microsoft Office\root\vfs\ProgramFilesCommonX86\Microsoft Shared\Office16\OSPP.VBS",
            "${env:CommonProgramFiles}\Microsoft Shared\Office16\OSPP.VBS",
            "${env:ProgramFiles(x86)}\Microsoft Office\root\vfs\ProgramFilesCommonX86\Microsoft Shared\Office16\OSPP.VBS"
        )
        $osppFound = $officeOspp | Where-Object { Test-Path $_ } | Select-Object -First 1
        if ($osppFound) {
            $officeStatus = cscript //nologo "$osppFound" /dstatus 2>&1 | Out-String
            if ($officeStatus -match 'LICENSED') {
                $report += '  [OK] Office activated'
            } else {
                $report += '  [!]  Office: not activated'
            }
        } else {
            $report += '  [--] Office not installed'
        }
        $report += ''

        # --- Windows Update ---
        $report += '--- WINDOWS UPDATE ---'
        $wuLog = Get-ChildItem "$env:TEMP\WindowsUpdate_*.log" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if ($wuLog) {
            $wuContent = Get-Content $wuLog.FullName -Tail 5
            foreach ($line in $wuContent) {
                $report += '  ' + $line
            }
        } else {
            $report += '  Update log not found'
        }
        $report += ''

        $report += '--- LOGS ---'
        $report += '  Main log: ' + $logPath
        if ($wuLog) { $report += '  Update log: ' + $wuLog.FullName }
        $report += ''
        $report += '============================================================'

        $report | Out-File -FilePath $reportPath -Encoding UTF8 -Force
        Write-Log "Installation report saved to desktop: $reportPath" -Level "SUCCESS"
    } catch {
        Write-Log "Could not generate report: $_" -Level "WARNING"
    }

    # Create activation shortcut on desktop (2 files: .cmd launcher + .ps1 script)
    try {
        $desktopPath = "$env:PUBLIC\Desktop"
        
        $helperPs1 = @'
Start-Process powershell.exe -Verb RunAs -ArgumentList '-NoProfile','-ExecutionPolicy','Bypass','-Command','irm https://get.activated.win | iex'
'@
        Set-Content -Path "$desktopPath\activate_helper.ps1" -Value $helperPs1 -Encoding UTF8 -Force
        
        $launcherCmd = @'
@echo off
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0activate_helper.ps1"
'@
        Set-Content -Path "$desktopPath\Activate_Windows_Office.cmd" -Value $launcherCmd -Encoding ASCII -Force
        
        Write-Log "Activation shortcut created on desktop (cmd + ps1)" -Level "SUCCESS"
    } catch {
        Write-Log "Could not create activation shortcut: $_" -Level "WARNING"
    }

    # Wait for Windows Update to finish before rebooting
    if ($updateProcess -and -not $updateProcess.HasExited) {
        Write-Host ""
        Write-Host "Waiting for Windows Update to finish..." -ForegroundColor Yellow
        Write-Log "Waiting for Windows Update process (PID: $($updateProcess.Id)) to finish..."
        $updateProcess.WaitForExit()
        Write-Log "Windows Update process finished" -Level "SUCCESS"
    }

    # Auto-reboot
    Write-Host ""
    Write-Host "============================================" -ForegroundColor Yellow
    Write-Host "  ALL PROCESSES COMPLETED"                   -ForegroundColor Yellow
    Write-Host "  COMPUTER WILL REBOOT IN 30 SECONDS"       -ForegroundColor Yellow
    Write-Host "  After reboot, run Activate_Windows_Office" -ForegroundColor Yellow
    Write-Host "  shortcut on the desktop."                  -ForegroundColor Yellow
    Write-Host "============================================" -ForegroundColor Yellow
    Write-Host ""
    Write-Log "All processes completed. Scheduling reboot in 30 seconds..." -Level "INFO"
    Start-Sleep -Seconds 30
    Restart-Computer -Force
}
catch {
    Write-Host ""
    Write-Host "CRITICAL ERROR OCCURRED!" -ForegroundColor Red
    Write-Host "Error: $_" -ForegroundColor Red
    Write-Log "[ERROR] CRITICAL ERROR: $_"
    Start-Sleep -Seconds 10
}
