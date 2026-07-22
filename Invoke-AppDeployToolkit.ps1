<#
.SYNOPSIS
    Microsoft 365 Apps for Enterprise - PSADT v4 Deployment Script
.DESCRIPTION
    Deploys M365 Apps via ODT with pre-install scrub using OfficeScrubberAIO.cmd (abbodi1406).
    OfficeScrubberAIO is a single self-contained script -- no companion bin\ folder required.
    It extracts its own VBS cabinet at runtime and accepts unattended switches via CLI arguments.
    Configuration XML is sourced from a publicly accessible Azure Blob Storage URL.

    Unattended scrubber arguments:
      /A  = Scrub ALL Office versions
      /C  = Scrub Click-to-Run (C2R) only  <-- used here
      /P  = Scrub Office UWP Apps           <-- used here

.PARAMETER DeploymentType
    Install or Uninstall.
.PARAMETER DeployMode
    Interactive, Silent, or NonInteractive.
.PARAMETER XmlInstallUrl
    Public URL for the ODT install configuration XML hosted in Azure Blob Storage.
.PARAMETER XmlUninstallUrl
    Public URL for the ODT uninstall configuration XML hosted in Azure Blob Storage.
.NOTES
    Toolkit:    PSAppDeployToolkit v4.x
    Author:     Chris McFall
    Intune:     Win32 App deployment
    Date:       14/05/2026

    Package structure:
      M365Apps_Deploy\
      ├── Invoke-AppDeployToolkit.exe
      ├── Invoke-AppDeployToolkit.ps1
      ├── PSAppDeployToolkit\
      │   └── (PSADT v4 module files - do not modify)
      └── Files\
          ├── setup.exe                  <- ODT setup binary
          └── OfficeScrubberAIO.cmd      <- abbodi1406 AIO scrubber (single file, self-contained)

    Intune Install command:
      Invoke-AppDeployToolkit.exe -DeploymentType Install -DeployMode Silent -XmlInstallUrl "https://<sa>.blob.core.windows.net/<container>/M365Apps-Configure.xml"

    Intune Uninstall command:
      Invoke-AppDeployToolkit.exe -DeploymentType Uninstall -DeployMode Silent -XmlUninstallUrl "https://<sa>.blob.core.windows.net/<container>/M365Apps-Remove.xml"

    Local test (as admin/SYSTEM, calling .ps1 directly):
      & ".\Invoke-AppDeployToolkit.ps1" -DeploymentType Install -DeployMode Silent -XmlInstallUrl "https://..."

    Detection rule (Intune) -- must match $PackageName below:
      File exists: %ProgramData%\Microsoft\Microsoft365AppsEnterprise\Microsoft_Microsoft365AppsEnterprise_1.0.1_x64_EN_01.ps1.tag
#>

[CmdletBinding()]
Param (
    [Parameter(Mandatory = $false)]
    [ValidateSet('Install', 'Uninstall')]
    [String]$DeploymentType = 'Install',

    [Parameter(Mandatory = $false)]
    [ValidateSet('Interactive', 'Silent', 'NonInteractive')]
    [String]$DeployMode = 'NonInteractive',

    [Parameter(Mandatory = $false)]
    [Switch]$AllowRebootPassThru,

    [Parameter(Mandatory = $false)]
    [Switch]$TerminalServerMode,

    [Parameter(Mandatory = $false)]
    [Switch]$DisableLogging,

    ## Public URL for the ODT install XML in Azure Blob Storage
    [Parameter(Mandatory = $false)]
    [String]$XmlInstallUrl = '',

    ## Public URL for the ODT uninstall XML in Azure Blob Storage
    [Parameter(Mandatory = $false)]
    [String]$XmlUninstallUrl = ''
)

##*===============================================
##* APP / PACKAGE VARIABLES
##*===============================================
## These match the client packaging convention and drive log file naming,
## tag file placement, and the PSADT session metadata.
[String]$Publisher   = 'Microsoft'
[String]$DisplayName = 'Microsoft 365 Apps for Enterprise'
[String]$AppName     = 'Microsoft365AppsEnterprise'
[String]$PackageName = 'Microsoft_Microsoft365AppsEnterprise_1.0.1_x64_EN_01'
[String]$Version     = '1.0.1'

##*===============================================
##* CLIENT LOG / TAG PATHS
##*===============================================
## Log format:  datetime - Section - Message - Severity
## Matches the Write-Log convention used across client packaging scripts.
[String]$ScriptName  = $MyInvocation.MyCommand.Name
[String]$DataFolder  = "$([Environment]::GetEnvironmentVariable('ProgramData'))\$Publisher\$AppName"
[String]$LogFolder   = "$([Environment]::GetEnvironmentVariable('WINDIR'))\Temp"
[String]$LogFile     = "$LogFolder\$($PackageName)_$($ScriptName).log"

##*===============================================
##* CLIENT LOGGING FUNCTION
##*===============================================
Function Write-ClientLog {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory)]
        [string]$TextBlock,
        [string]$Section = 'General',
        [ValidateSet('DEBUG','INFO','WARN','ERROR')]
        [string]$Severity = 'INFO'
    )
    $timestamp = Get-Date
    $output    = "$timestamp - $Section - $TextBlock - $Severity"
    Add-Content -Path $LogFile -Value $output
    Write-Host $output
}

##*===============================================
##* TAG FILE FUNCTIONS
##*===============================================
Function New-TagFile {
    $section = 'Create TagFile'
    Write-ClientLog -Section $section -TextBlock "Creating detection folder - $DataFolder"
    If (-not (Test-Path $DataFolder)) {
        Try {
            New-Item -Path $DataFolder -ItemType Directory -ErrorAction Stop | Out-Null
            Write-ClientLog -Section $section -TextBlock "   >> Success"
        } Catch {
            Write-ClientLog -Section $section -TextBlock "   >> Failed - '$_'" -Severity ERROR
        }
    } Else {
        Write-ClientLog -Section $section -TextBlock "   >> Folder already exists" -Severity WARN
    }
    $tagPath = "$DataFolder\$PackageName.ps1.tag"
    Write-ClientLog -Section $section -TextBlock "Creating tag file - $tagPath"
    Try {
        Set-Content -Path $tagPath -Value 'Installed' -ErrorAction Stop
        Write-ClientLog -Section $section -TextBlock "   >> Success"
    } Catch {
        Write-ClientLog -Section $section -TextBlock "   >> Failed - '$_'" -Severity ERROR
    }
}

Function Remove-TagFile {
    $section = 'Remove TagFile'
    $tagPath = "$DataFolder\$PackageName.ps1.tag"
    Write-ClientLog -Section $section -TextBlock "Removing tag file - $tagPath"
    If (Test-Path $tagPath) {
        Try {
            Remove-Item -Path $tagPath -Force -ErrorAction Stop
            Write-ClientLog -Section $section -TextBlock "   >> Success"
        } Catch {
            Write-ClientLog -Section $section -TextBlock "   >> Failed - '$_'" -Severity ERROR
        }
    } Else {
        Write-ClientLog -Section $section -TextBlock "   >> Tag file not found - skipping" -Severity WARN
    }
}

##*===============================================
##* UWP OFFICE REMOVAL FUNCTION
##*===============================================
Function Remove-OfficeUWPAllUsers {
    ## Removes the specified UWP/Appx packages registered under ALL user profiles on the device.
    ## Running as SYSTEM via Intune means Get-AppxPackage only sees the SYSTEM context --
    ## per-user registrations (including the Autopilot OOBE user) require enumerating each
    ## SID explicitly. We handle three populations:
    ##   1. Currently loaded user hives  (active/recent sessions)
    ##   2. Unloaded user hives on disk  (previous users, OOBE user)
    ##   3. Provisioned packages         (prevents registration for future new users)
    ##
    ## -PackageNames controls what is targeted, so the same engine serves two callers:
    ##   * The full Office UWP stub list (default) -- only when a scrub/install is happening.
    ##   * Just "new Outlook for Windows" (Microsoft.OutlookForWindows, binary OLK.exe, under
    ##     %ProgramFiles%\WindowsApps\Microsoft.OutlookForWindows_*) -- called unconditionally
    ##     before the pre-flight check so it is removed even on healthy classic-Office machines.
    Param (
        [Parameter(Mandatory = $false)]
        [String[]]$PackageNames = @(
            'Microsoft.MicrosoftOfficeHub'
            'Microsoft.Office.Desktop'
            'Microsoft.Office.Desktop.Access'
            'Microsoft.Office.Desktop.Excel'
            'Microsoft.Office.Desktop.Outlook'
            'Microsoft.Office.Desktop.PowerPoint'
            'Microsoft.Office.Desktop.Word'
            'Microsoft.OutlookForWindows'           ## "new Outlook" (OLK.exe)
        )
    )

    $section        = 'Remove UWP Office'
    $packageNames   = $PackageNames
    $newOutlookInScope = ($packageNames -contains 'Microsoft.OutlookForWindows')

    Write-ClientLog -Section $section -TextBlock "Targeting [$($packageNames.Count)] package(s): $($packageNames -join ', ')"
    Write-ADTLogEntry -Message "Starting per-user Office UWP removal across all profiles."
    Write-ClientLog -Section $section -TextBlock "_____________________________________________________________________"
    Write-ClientLog -Section $section -TextBlock "Starting Office UWP removal for all user profiles"

    ## Stop any running new-Outlook (OLK.exe) instances first -- a live process can block the
    ## Appx removal below. OLK.exe runs per-user; as SYSTEM, Stop-Process by name terminates
    ## all sessions' instances. Non-fatal if nothing is running. Only when new Outlook is in scope.
    If ($newOutlookInScope) {
        Try {
            $olkProcs = Get-Process -Name 'OLK' -ErrorAction SilentlyContinue
            If ($olkProcs) {
                Write-ClientLog -Section $section -TextBlock "Stopping [$($olkProcs.Count)] running OLK.exe (new Outlook) process(es)"
                $olkProcs | Stop-Process -Force -ErrorAction SilentlyContinue
            } Else {
                Write-ClientLog -Section $section -TextBlock "No running OLK.exe (new Outlook) processes found"
            }
        } Catch {
            Write-ClientLog -Section $section -TextBlock "   >> Failed to stop OLK.exe - continuing: $($_.Exception.Message)" -Severity WARN
        }
    }

    ## Build list of all user SIDs from the registry profile list
    $profileListPath = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList'
    $allSIDs = (Get-ChildItem -Path $profileListPath -ErrorAction SilentlyContinue).PSChildName |
        Where-Object { $_ -match '^S-1-5-21-' }  ## Filter to real user SIDs only, exclude SYSTEM/service accounts

    Write-ClientLog -Section $section -TextBlock "Found [$($allSIDs.Count)] user profile SID(s) to process"

    ForEach ($sid in $allSIDs) {
        $profilePath = (Get-ItemProperty -Path "$profileListPath\$sid" -ErrorAction SilentlyContinue).ProfileImagePath
        Write-ClientLog -Section $section -TextBlock "Processing SID [$sid] - Profile [$profilePath]"

        ## Check if this hive is already loaded
        $hiveLoaded  = Test-Path -Path "Registry::HKEY_USERS\$sid"
        $hiveMounted = $false

        If (-not $hiveLoaded) {
            ## Hive not loaded -- mount it temporarily from the user's NTUSER.DAT
            $ntUserDat = Join-Path -Path $profilePath -ChildPath 'NTUSER.DAT'
            If (Test-Path -LiteralPath $ntUserDat) {
                Try {
                    $regLoad = & reg.exe load "HKU\$sid" "$ntUserDat" 2>&1
                    If ($LASTEXITCODE -eq 0) {
                        $hiveMounted = $true
                        Write-ClientLog -Section $section -TextBlock "   >> Mounted hive for SID [$sid]"
                    } Else {
                        Write-ClientLog -Section $section -TextBlock "   >> Failed to mount hive for SID [$sid] - reg.exe exit [$LASTEXITCODE] output [$regLoad] - skipping" -Severity WARN
                        Continue
                    }
                } Catch {
                    Write-ClientLog -Section $section -TextBlock "   >> Exception mounting hive for SID [$sid] - '$_' - skipping" -Severity WARN
                    Continue
                }
            } Else {
                Write-ClientLog -Section $section -TextBlock "   >> NTUSER.DAT not found at [$ntUserDat] - skipping" -Severity WARN
                Continue
            }
        } Else {
            Write-ClientLog -Section $section -TextBlock "   >> Hive already loaded for SID [$sid]"
        }

        ## Remove each Office UWP package for this user
        ForEach ($packageName in $packageNames) {
            Try {
                $packages = Get-AppxPackage -User $sid -Name "$packageName*" -ErrorAction SilentlyContinue
                If ($packages) {
                    ForEach ($pkg in $packages) {
                        Write-ClientLog -Section $section -TextBlock "   >> Removing [$($pkg.PackageFullName)] for SID [$sid]"
                        Remove-AppxPackage -Package $pkg.PackageFullName -User $sid -ErrorAction Stop
                        Write-ClientLog -Section $section -TextBlock "   >> Success"
                    }
                } Else {
                    Write-ClientLog -Section $section -TextBlock "   >> No packages matching [$packageName] found for SID [$sid] - skipping"
                }
            } Catch {
                Write-ClientLog -Section $section -TextBlock "   >> Failed to remove [$packageName] for SID [$sid] - '$_'" -Severity ERROR
            }
        }

        ## Unmount hive if we mounted it -- GC collect first to release any open handles
        If ($hiveMounted) {
            Try {
                [GC]::Collect()
                [GC]::WaitForPendingFinalizers()
                Start-Sleep -Seconds 3
                $regUnload = & reg.exe unload "HKU\$sid" 2>&1
                If ($LASTEXITCODE -eq 0) {
                    Write-ClientLog -Section $section -TextBlock "   >> Unloaded hive for SID [$sid]"
                } Else {
                    Write-ClientLog -Section $section -TextBlock "   >> Failed to unload hive for SID [$sid] - reg.exe exit [$LASTEXITCODE] output [$regUnload]" -Severity WARN
                }
            } Catch {
                Write-ClientLog -Section $section -TextBlock "   >> Exception unloading hive for SID [$sid] - '$_'" -Severity WARN
            }
        }
    }

    ## Remove provisioned packages -- prevents registration for any future new user profiles
    Write-ClientLog -Section $section -TextBlock "Removing provisioned Office UWP packages"
    ForEach ($packageName in $packageNames) {
        Try {
            $provPkgs = Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue |
                Where-Object { $_.DisplayName -like "$packageName*" }
            If ($provPkgs) {
                ForEach ($pkg in $provPkgs) {
                    Write-ClientLog -Section $section -TextBlock "   >> Removing provisioned [$($pkg.PackageName)]"
                    Remove-AppxProvisionedPackage -Online -PackageName $pkg.PackageName -ErrorAction Stop | Out-Null
                    Write-ClientLog -Section $section -TextBlock "   >> Success"
                }
            }
        } Catch {
            Write-ClientLog -Section $section -TextBlock "   >> Failed to remove provisioned [$packageName] - '$_'" -Severity ERROR
        }
    }

    ## Verify new Outlook is gone. Removing the Appx package deregisters it, but the staged
    ## payload folder under WindowsApps is owned by TrustedInstaller and may linger until the
    ## next servicing pass -- so a leftover folder is informational, not a hard failure. We key
    ## the check on OLK.exe still being present, which is the meaningful "still installed" signal.
    If ($newOutlookInScope) {
      Try {
        $olkDirs = Get-ChildItem -Path (Join-Path -Path $env:ProgramFiles -ChildPath 'WindowsApps') `
            -Filter 'Microsoft.OutlookForWindows_*_x64__*' -Directory -ErrorAction SilentlyContinue
        $olkExeFound = $false
        ForEach ($dir in $olkDirs) {
            If (Test-Path -LiteralPath (Join-Path -Path $dir.FullName -ChildPath 'OLK.exe')) {
                $olkExeFound = $true
                Write-ClientLog -Section $section -TextBlock "   >> New Outlook payload still present (OLK.exe) at [$($dir.FullName)] - folder may clear on next servicing pass" -Severity WARN
            }
        }
        If (-not $olkExeFound) {
            Write-ClientLog -Section $section -TextBlock "Verification: no OLK.exe (new Outlook) payload remaining under WindowsApps"
        }
      } Catch {
        Write-ClientLog -Section $section -TextBlock "   >> New Outlook verification check failed - continuing: $($_.Exception.Message)" -Severity WARN
      }
    }

    Write-ADTLogEntry -Message "Office UWP removal across all user profiles completed."
    Write-ClientLog -Section $section -TextBlock "Office UWP removal completed"
    Write-ClientLog -Section $section -TextBlock "_____________________________________________________________________"
}

##*===============================================
##* LOG PACKAGE INFO
##*===============================================
If (-not (Test-Path $LogFolder)) { New-Item -Path $LogFolder -ItemType Directory -Force | Out-Null }
Write-ClientLog -Section 'Package Information' -TextBlock "_____________________________________________________________________"
Write-ClientLog -Section 'Package Information' -TextBlock "ScriptName      = $ScriptName"
Write-ClientLog -Section 'Package Information' -TextBlock "Publisher       = $Publisher"
Write-ClientLog -Section 'Package Information' -TextBlock "DisplayName     = $DisplayName"
Write-ClientLog -Section 'Package Information' -TextBlock "AppName         = $AppName"
Write-ClientLog -Section 'Package Information' -TextBlock "PackageName     = $PackageName"
Write-ClientLog -Section 'Package Information' -TextBlock "Version         = $Version"
Write-ClientLog -Section 'Package Information' -TextBlock "DeploymentType  = $DeploymentType"
Write-ClientLog -Section 'Package Information' -TextBlock "XmlInstallUrl   = $XmlInstallUrl"
Write-ClientLog -Section 'Package Information' -TextBlock "XmlUninstallUrl = $XmlUninstallUrl"
Write-ClientLog -Section 'Package Information' -TextBlock "_____________________________________________________________________"

## Temp paths for downloaded XMLs
[String]$xmlInstallTempPath   = "$env:TEMP\M365Apps-Configure.xml"
[String]$xmlUninstallTempPath = "$env:TEMP\M365Apps-Remove.xml"

## Force TLS 1.2 for the blob downloads. Under Windows PowerShell 5.1 running as SYSTEM the
## default SecurityProtocol can negotiate a protocol Azure Blob rejects, producing intermittent
## "Could not create SSL/TLS secure channel" failures. Set it process-wide up front.
Try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
} Catch {
    Write-ClientLog -Section 'Package Information' -TextBlock "Could not set TLS 1.2 - continuing: $($_.Exception.Message)" -Severity WARN
}

##*===============================================
##* PSADT v4 MODULE IMPORT
##*===============================================
Try {
    $modulePath = "$PSScriptRoot\PSAppDeployToolkit\PSAppDeployToolkit.psd1"
    If (-not (Test-Path -LiteralPath $modulePath -PathType Leaf)) {
        Throw "PSAppDeployToolkit module not found at [$modulePath]."
    }
    Import-Module -Name $modulePath -Force -ErrorAction Stop
}
Catch {
    Write-Error -Message "Failed to import PSAppDeployToolkit module: $($_.Exception.Message)"
    Exit 60008
}

##*===============================================
##* OPEN ADT SESSION
##*===============================================
$adtSession = @{
    ## Application metadata -- displayed in the UI header and PSADT logs
    AppVendor                   = $Publisher
    AppName                     = $DisplayName
    AppVersion                  = $Version
    AppArch                     = 'x64'
    AppLang                     = 'EN'
    AppRevision                 = '01'
    AppScriptDate               = [DateTime]'2026-05-14'
    AppScriptAuthor             = 'Chris McFall'
    DeployAppScriptFriendlyName = $AppName
    InstallName                 = $AppName
    InstallTitle                = $DisplayName
    ## DeploymentType, DeployMode, AllowRebootPassThru, TerminalServerMode, DisableLogging
    ## are passed via @PSBoundParameters and must not be duplicated here
}

Try {
    ## Remove our custom parameters before splatting -- Open-ADTSession does not accept them
    $adtParams = $PSBoundParameters
    $null = $adtParams.Remove('XmlInstallUrl')
    $null = $adtParams.Remove('XmlUninstallUrl')
    Open-ADTSession -SessionState $ExecutionContext.SessionState @adtSession @adtParams
    ## Get the live session object so we can read populated properties like DirFiles
    $adtSession = Get-ADTSession
}
Catch {
    Write-Error -Message "Failed to open ADT session: $($_.Exception.Message)"
    Exit 60008
}

## Resolve the payload folder from the live session (DirFiles is only populated if a Files\
## folder exists next to the script). Fall back to $PSScriptRoot\Files to guarantee a valid
## path either way.
$dirFiles = $adtSession.DirFiles
If ([String]::IsNullOrWhiteSpace($dirFiles)) {
    $dirFiles = Join-Path -Path $PSScriptRoot -ChildPath 'Files'
}

##*===============================================
##* INSTALLATION
##*===============================================
If ($DeploymentType -ine 'Uninstall') {

    ##*-------------------------------------------
    ##* PRE-INSTALLATION
    ##*-------------------------------------------

    ## Validate URL was supplied
    If ([String]::IsNullOrWhiteSpace($XmlInstallUrl)) {
        Write-ADTLogEntry -Message "No XmlInstallUrl parameter was provided. Cannot continue with installation." -Severity 3
        Write-ClientLog -Section 'Pre-Installation' -TextBlock "No XmlInstallUrl provided - aborting" -Severity ERROR
        Close-ADTSession -ExitCode 60001
    }

    ##*-------------------------------------------
    ##* REMOVE NEW OUTLOOK (UNCONDITIONAL)
    ##*-------------------------------------------
    ## Runs BEFORE the pre-flight check so it executes even on machines with a healthy classic
    ## Office/Outlook install (which would otherwise fast-exit and never reach the scrub phase).
    ## Scoped to new Outlook only -- it does NOT touch classic Office and does NOT influence the
    ## pre-flight health decision. Non-fatal: a failure here must not block the install.
    Try {
        Remove-OfficeUWPAllUsers -PackageNames 'Microsoft.OutlookForWindows'
    } Catch {
        Write-ADTLogEntry -Message "New Outlook removal threw an unexpected exception - continuing. Error: $($_.Exception.Message)" -Severity 2
        Write-ClientLog -Section 'Pre-Installation' -TextBlock "New Outlook removal exception - continuing: $($_.Exception.Message)" -Severity WARN
    }

    ##*-------------------------------------------
    ##* PRE-FLIGHT HEALTH CHECK
    ##*-------------------------------------------
    ## Before doing anything destructive, check whether a healthy M365 Apps installation
    ## already exists. Two-stage evaluation:
    ##
    ##   Stage 1 - Tag file check (fast exit):
    ##     If the tag file is already present this package has run before and Office was
    ##     confirmed healthy or installed. Exit 0 immediately without evaluating anything else.
    ##
    ##   Stage 2 - Health checks (only reached if tag file is absent):
    ##     Checks 2-5 are all evaluated. If all four pass, Office is considered healthy --
    ##     write the tag file and exit 0 with no reinstall. If any single check fails,
    ##     log the reason and proceed with the full scrub and install.
    ##
    ##   Check 1. Tag file present          -> fast exit, already managed by this package
    ##   Check 2. ProductReleaseIds contains O365ProPlusRetail  -> correct product registered
    ##   Check 3. VersionToReport is a parseable version        -> C2R engine has coherent build
    ##   Check 4. CDNBaseUrl is populated                       -> machine on an update channel
    ##   Check 5. Core Office executables exist on disk         -> binaries actually present

    $preFlightSection = 'Pre-Flight Check'
    $c2rRegPath       = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
    $tagPath          = "$DataFolder\$PackageName.ps1.tag"
    $preFlightReasons = @()

    Write-ADTLogEntry -Message "Starting pre-flight health check for existing M365 Apps installation."
    Write-ClientLog -Section $preFlightSection -TextBlock "_____________________________________________________________________"
    Write-ClientLog -Section $preFlightSection -TextBlock "Starting pre-flight health check"

    ## Check 1: Tag file already present
    If (Test-Path -LiteralPath $tagPath) {
        Write-ADTLogEntry -Message "Pre-flight: Tag file already present at [$tagPath]. Installation previously completed by this package. Exiting clean."
        Write-ClientLog -Section $preFlightSection -TextBlock "CHECK 1 PASS: Tag file already present - no action required"
        Write-ClientLog -Section $preFlightSection -TextBlock "_____________________________________________________________________"
        Close-ADTSession -ExitCode 0
    } Else {
        Write-ClientLog -Section $preFlightSection -TextBlock "CHECK 1: No tag file found - continuing checks"
    }

    ## Check 2: ProductReleaseIds contains O365ProPlusRetail
    $c2rConfig     = Get-ItemProperty -Path $c2rRegPath -ErrorAction SilentlyContinue
    $releaseIds    = $c2rConfig.ProductReleaseIds
    If (-not [String]::IsNullOrWhiteSpace($releaseIds) -and $releaseIds -match 'O365ProPlusRetail') {
        Write-ClientLog -Section $preFlightSection -TextBlock "CHECK 2 PASS: ProductReleaseIds = [$releaseIds]"
    } Else {
        $preFlightReasons += "CHECK 2 FAIL: ProductReleaseIds missing or does not contain O365ProPlusRetail (found: [$releaseIds])"
        Write-ClientLog -Section $preFlightSection -TextBlock $preFlightReasons[-1] -Severity WARN
    }

    ## Check 3: VersionToReport is a parseable version
    $versionToReport = $c2rConfig.VersionToReport
    $parsedVersion   = $null
    If (-not [String]::IsNullOrWhiteSpace($versionToReport) -and [Version]::TryParse($versionToReport, [ref]$parsedVersion)) {
        Write-ClientLog -Section $preFlightSection -TextBlock "CHECK 3 PASS: VersionToReport = [$versionToReport]"
    } Else {
        $preFlightReasons += "CHECK 3 FAIL: VersionToReport missing or not a valid version (found: [$versionToReport])"
        Write-ClientLog -Section $preFlightSection -TextBlock $preFlightReasons[-1] -Severity WARN
    }

    ## Check 4: CDNBaseUrl is populated (machine is connected to an update channel)
    $cdnBaseUrl = $c2rConfig.CDNBaseUrl
    If (-not [String]::IsNullOrWhiteSpace($cdnBaseUrl)) {
        Write-ClientLog -Section $preFlightSection -TextBlock "CHECK 4 PASS: CDNBaseUrl = [$cdnBaseUrl]"
    } Else {
        $preFlightReasons += "CHECK 4 FAIL: CDNBaseUrl is empty - machine may not be connected to an update channel"
        Write-ClientLog -Section $preFlightSection -TextBlock $preFlightReasons[-1] -Severity WARN
    }

    ## Check 5: Core Office executables exist on disk
    ## InstallationPath from the registry points to the C2R engine root, not the Office
    ## binaries root, and can vary. We instead check both known install locations directly
    ## to avoid 32/64-bit registry redirection issues when running as SYSTEM via Intune.
    $officeExes      = @('WINWORD.EXE', 'EXCEL.EXE', 'OUTLOOK.EXE')
    $programFilesX86 = [Environment]::GetFolderPath('ProgramFilesX86')
    $officeBinPaths  = @(
        "$env:ProgramFiles\Microsoft Office\Root\Office16"        ## x64 standard
        "$programFilesX86\Microsoft Office\Root\Office16"         ## x86 on 64-bit OS
        "$env:ProgramFiles\Microsoft Office\Office16"              ## alternative x64 layout
        "$programFilesX86\Microsoft Office\Office16"               ## alternative x86 layout
    )
    $missingExes     = @()
    $foundBinPath    = $null

    ## Find whichever bin path actually exists on this machine
    ForEach ($binPath in $officeBinPaths) {
        If (Test-Path -Path $binPath) {
            $foundBinPath = $binPath
            Break
        }
    }

    If ($foundBinPath) {
        ForEach ($exe in $officeExes) {
            $exePath = Join-Path -Path $foundBinPath -ChildPath $exe
            If (-not (Test-Path -LiteralPath $exePath)) {
                $missingExes += $exe
            }
        }
        If ($missingExes.Count -eq 0) {
            Write-ClientLog -Section $preFlightSection -TextBlock "CHECK 5 PASS: All core Office executables present at [$foundBinPath]"
        } Else {
            $preFlightReasons += "CHECK 5 FAIL: Missing executables [$($missingExes -join ', ')] under [$foundBinPath]"
            Write-ClientLog -Section $preFlightSection -TextBlock $preFlightReasons[-1] -Severity WARN
        }
    } Else {
        $preFlightReasons += "CHECK 5 FAIL: No Office binary path found. Checked: [$($officeBinPaths -join ' | ')]"
        Write-ClientLog -Section $preFlightSection -TextBlock $preFlightReasons[-1] -Severity WARN
    }

    ## Evaluate overall pre-flight result
    ## Checks 2-5 must all pass to consider Office healthy and skip reinstall.
    ## Check 1 (tag file) is handled separately above as an immediate fast-exit.
    If ($preFlightReasons.Count -eq 0) {
        ## All four health checks passed -- healthy install detected, no reinstall needed
        Write-ADTLogEntry -Message "Pre-flight PASSED: All health checks passed. M365 Apps healthy (Version: $versionToReport). Writing tag file and exiting clean."
        Write-ClientLog -Section $preFlightSection -TextBlock "Pre-flight PASSED: All 4 health checks passed - Version [$versionToReport] Channel [$cdnBaseUrl]"
        Write-ClientLog -Section $preFlightSection -TextBlock "Writing tag file for future pre-flight detection and exiting"
        Write-ClientLog -Section $preFlightSection -TextBlock "_____________________________________________________________________"
        New-TagFile
        Close-ADTSession -ExitCode 0
    } Else {
        ## One or more of checks 2-5 failed -- proceed with full scrub and install
        Write-ADTLogEntry -Message "Pre-flight FAILED: $($preFlightReasons.Count) of 4 health check(s) did not pass. Proceeding with scrub and install."
        Write-ClientLog -Section $preFlightSection -TextBlock "Pre-flight FAILED: $($preFlightReasons.Count) of 4 health check(s) failed - proceeding with full install"
        ForEach ($reason in $preFlightReasons) {
            Write-ClientLog -Section $preFlightSection -TextBlock "  -- $reason" -Severity WARN
        }
        Write-ClientLog -Section $preFlightSection -TextBlock "_____________________________________________________________________"
    }

    Show-ADTInstallationProgress -StatusMessage "Preparing to install Microsoft 365 Apps. Removing any existing Office installations..."

    ## Step 1: Remove Office UWP packages across all user profiles including the Autopilot
    ## OOBE user SID, which has already claimed stub registrations before Intune runs.
    ## This must run before the scrubber so /P has a clean provisioned layer to work against.
    ## Wrapped in try/catch -- a failure here is non-fatal, scrub + install will still proceed.
    Try {
        Remove-OfficeUWPAllUsers
    } Catch {
        Write-ADTLogEntry -Message "Remove-OfficeUWPAllUsers threw an unexpected exception - continuing. Error: $($_.Exception.Message)" -Severity 2
        Write-ClientLog -Section 'Pre-Installation' -TextBlock "Remove-OfficeUWPAllUsers exception - continuing: $($_.Exception.Message)" -Severity WARN
    }

    ## Step 2: Run OfficeScrubberAIO.cmd with /C /P to silently remove:
    ##   /C = All Click-to-Run (C2R) installations
    ##   /P = All Office UWP apps (present by default on Surface devices in multiple languages)
    ## The AIO variant is fully self-contained -- it carries its own VBS cabinet embedded in the
    ## script body and extracts it at runtime. No companion files needed.
    ## Running as SYSTEM under Intune satisfies the elevation requirement automatically.
    $scrubberPath = Join-Path -Path $dirFiles -ChildPath 'OfficeScrubberAIO.cmd'

    If (Test-Path -LiteralPath $scrubberPath) {
        Write-ADTLogEntry -Message "Running OfficeScrubberAIO.cmd /C /P (unattended C2R + UWP scrub) from [$scrubberPath]."
    Write-ClientLog -Section 'Pre-Installation' -TextBlock "Running OfficeScrubberAIO.cmd /C /P (C2R + UWP) from [$scrubberPath]"

        ## Wrapped in try/catch -- an unexpected scrubber exit code (one not in SuccessExitCodes)
        ## makes Start-ADTProcess throw. Treat that as non-fatal: log it and continue to the
        ## install rather than aborting, on the basis that a clean ODT install is still worth
        ## attempting even if the scrub was imperfect.
        Try {
            Start-ADTProcess -FilePath 'cmd.exe' `
                -ArgumentList "/c `"$scrubberPath`" /C /P" `
                -CreateNoWindow `
                -WaitForChildProcesses `
                -SuccessExitCodes 0,1,2,3,5 `
                -RebootExitCodes 3010 `
                -PassThru | Out-Null

            Write-ADTLogEntry -Message "OfficeScrubberAIO.cmd completed."
            Write-ClientLog -Section 'Pre-Installation' -TextBlock "OfficeScrubberAIO.cmd completed"
        } Catch {
            Write-ADTLogEntry -Message "OfficeScrubberAIO.cmd returned an unexpected exit code - continuing to install. Error: $($_.Exception.Message)" -Severity 2
            Write-ClientLog -Section 'Pre-Installation' -TextBlock "OfficeScrubberAIO.cmd unexpected exit code - continuing to install: $($_.Exception.Message)" -Severity WARN
        }
    }
    Else {
        Write-ADTLogEntry -Message "OfficeScrubberAIO.cmd not found at [$scrubberPath]. Skipping pre-install scrub phase." -Severity 2
        Write-ClientLog -Section 'Pre-Installation' -TextBlock "OfficeScrubberAIO.cmd not found - skipping scrub" -Severity WARN
    }

    ##*-------------------------------------------
    ##* INSTALLATION
    ##*-------------------------------------------

    Show-ADTInstallationProgress -StatusMessage "Installing Microsoft 365 Apps. This may take 10-20 minutes. Please wait..."

    ## Download ODT configuration XML.
    ## ODT setup.exe requires a local filesystem path -- URLs are not supported directly.
    $xmlInstallTempPath = "$env:TEMP\M365Apps-Configure.xml"
    Write-ADTLogEntry -Message "Downloading ODT install XML from [$XmlInstallUrl] to [$xmlInstallTempPath]."
    Write-ClientLog -Section 'Installation' -TextBlock "Downloading ODT install XML from [$XmlInstallUrl]"

    $webClient = New-Object System.Net.WebClient
    Try {
        $webClient.DownloadFile($XmlInstallUrl, $xmlInstallTempPath)
        Write-ADTLogEntry -Message "ODT install XML downloaded successfully."
        Write-ClientLog -Section 'Installation' -TextBlock "ODT install XML downloaded successfully"
    }
    Catch {
        Write-ADTLogEntry -Message "Failed to download install configuration XML. Error: $($_.Exception.Message)" -Severity 3
        Write-ClientLog -Section 'Installation' -TextBlock "Failed to download install XML - $($_.Exception.Message)" -Severity ERROR
        Close-ADTSession -ExitCode 60001
    }
    Finally {
        $webClient.Dispose()
    }

    ## Validate the downloaded file is present and non-empty
    If (-not (Test-Path -LiteralPath $xmlInstallTempPath) -or (Get-Item -LiteralPath $xmlInstallTempPath).Length -eq 0) {
        Write-ADTLogEntry -Message "Downloaded install XML is missing or empty. Aborting installation." -Severity 3
        Close-ADTSession -ExitCode 60001
    }

    ## Validate the file is well-formed XML before handing it to ODT. A truncated download or an
    ## error page served in place of the blob would otherwise be passed straight to setup.exe.
    Try {
        [xml]$null = Get-Content -LiteralPath $xmlInstallTempPath -Raw
    }
    Catch {
        Write-ADTLogEntry -Message "Downloaded install XML is not valid XML. Aborting installation. Error: $($_.Exception.Message)" -Severity 3
        Write-ClientLog -Section 'Installation' -TextBlock "Downloaded install XML is malformed - aborting" -Severity ERROR
        Close-ADTSession -ExitCode 60001
    }

    ## Run ODT setup.exe with the downloaded configuration XML
    ## ODT setup.exe is a thin bootstrapper -- it launches OfficeClickToRun.exe and
    ## exits immediately. Start-ADTProcess with -WaitForChildProcesses does not reliably
    ## track the C2R child across session boundaries when running as SYSTEM.
    ## Instead we launch setup.exe directly, then poll until OfficeClickToRun.exe has
    ## started and fully exited before continuing.
    Write-ADTLogEntry -Message "Launching ODT setup.exe /configure."
    Write-ClientLog -Section 'Installation' -TextBlock "Launching ODT setup.exe /configure"
    $odtProcess = Start-Process -FilePath (Join-Path -Path $dirFiles -ChildPath 'setup.exe') `
        -ArgumentList "/configure `"$xmlInstallTempPath`"" `
        -WindowStyle 'Hidden' `
        -PassThru

    ## Wait for setup.exe bootstrapper to exit
    $odtProcess | Wait-Process
    Write-ADTLogEntry -Message "setup.exe bootstrapper exited. Polling for installation completion via registry."
    Write-ClientLog -Section 'Installation' -TextBlock "setup.exe bootstrapper exited - polling registry for completion"

    ## ODT hands off to the Click-to-Run service (ClickToRunSvc) which runs as a persistent
    ## background service -- it cannot be used as a completion signal. Instead we poll the
    ## C2R configuration registry key for ProductReleaseIds, which ODT only writes once the
    ## installation has fully completed. Poll for up to 90 minutes.
    $regPath        = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
    $maxWaitSeconds = 5400
    $pollInterval   = 15
    $elapsed        = 0

    While ($elapsed -lt $maxWaitSeconds) {
        $releaseIds = (Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue).ProductReleaseIds
        If (-not [String]::IsNullOrWhiteSpace($releaseIds)) {
            Write-ADTLogEntry -Message "ProductReleaseIds detected in registry [$releaseIds] -- installation complete."
            Write-ClientLog -Section 'Installation' -TextBlock "ProductReleaseIds detected [$releaseIds] -- installation complete"
            Break
        }
        Write-ADTLogEntry -Message "Waiting for ODT to complete... [$elapsed/$maxWaitSeconds seconds elapsed]"
        Start-Sleep -Seconds $pollInterval
        $elapsed += $pollInterval
    }

    If ($elapsed -ge $maxWaitSeconds) {
        Write-ADTLogEntry -Message "Timed out waiting for ProductReleaseIds registry key after [$maxWaitSeconds] seconds." -Severity 3
        Write-ClientLog -Section 'Installation' -TextBlock "Timed out waiting for ODT installation after [$maxWaitSeconds] seconds" -Severity ERROR
        Close-ADTSession -ExitCode 60002
    }

    ## Clean up temp XML
    If (Test-Path -LiteralPath $xmlInstallTempPath) {
        Remove-Item -LiteralPath $xmlInstallTempPath -Force -ErrorAction SilentlyContinue
        Write-ADTLogEntry -Message "Temporary install XML removed."
        Write-ClientLog -Section 'Installation' -TextBlock "Temporary install XML removed"
    }

    ##*-------------------------------------------
    ##* POST-INSTALLATION
    ##*-------------------------------------------

    New-TagFile
    Write-ADTLogEntry -Message "Microsoft 365 Apps installation phase completed."
    Write-ClientLog -Section 'Post-Installation' -TextBlock "Installation complete. Tag file written." 
}

##*===============================================
##* UNINSTALLATION
##*===============================================
ElseIf ($DeploymentType -ieq 'Uninstall') {

    ##*-------------------------------------------
    ##* PRE-UNINSTALLATION
    ##*-------------------------------------------

    ## Validate URL was supplied
    If ([String]::IsNullOrWhiteSpace($XmlUninstallUrl)) {
        Write-ADTLogEntry -Message "No XmlUninstallUrl parameter was provided. Cannot continue with uninstallation." -Severity 3
        Close-ADTSession -ExitCode 60001
    }

    Show-ADTInstallationProgress -StatusMessage "Removing Microsoft 365 Apps..."

    ##*-------------------------------------------
    ##* UNINSTALLATION
    ##*-------------------------------------------

    $xmlUninstallTempPath = "$env:TEMP\M365Apps-Remove.xml"
    Write-ADTLogEntry -Message "Downloading ODT uninstall XML from [$XmlUninstallUrl] to [$xmlUninstallTempPath]."

    $webClient = New-Object System.Net.WebClient
    Try {
        $webClient.DownloadFile($XmlUninstallUrl, $xmlUninstallTempPath)
        Write-ADTLogEntry -Message "ODT uninstall XML downloaded successfully."
    }
    Catch {
        Write-ADTLogEntry -Message "Failed to download uninstall configuration XML. Error: $($_.Exception.Message)" -Severity 3
        Close-ADTSession -ExitCode 60001
    }
    Finally {
        $webClient.Dispose()
    }

    If (-not (Test-Path -LiteralPath $xmlUninstallTempPath) -or (Get-Item -LiteralPath $xmlUninstallTempPath).Length -eq 0) {
        Write-ADTLogEntry -Message "Downloaded uninstall XML is missing or empty. Aborting." -Severity 3
        Close-ADTSession -ExitCode 60001
    }

    ## Validate the file is well-formed XML before handing it to ODT.
    Try {
        [xml]$null = Get-Content -LiteralPath $xmlUninstallTempPath -Raw
    }
    Catch {
        Write-ADTLogEntry -Message "Downloaded uninstall XML is not valid XML. Aborting. Error: $($_.Exception.Message)" -Severity 3
        Close-ADTSession -ExitCode 60001
    }

    Write-ADTLogEntry -Message "Launching ODT setup.exe /configure (uninstall)."
    $odtProcess = Start-Process -FilePath (Join-Path -Path $dirFiles -ChildPath 'setup.exe') `
        -ArgumentList "/configure `"$xmlUninstallTempPath`"" `
        -WindowStyle 'Hidden' `
        -PassThru

    $odtProcess | Wait-Process
    Write-ADTLogEntry -Message "setup.exe bootstrapper exited. Polling for uninstallation completion via registry."

    ## Poll for completion. ODT frequently leaves the ClickToRun and even the Configuration
    ## key behind after a successful removal, so waiting for the whole key to vanish can false-
    ## timeout on an otherwise-successful uninstall. Instead, treat Office as removed when EITHER
    ## the Configuration key is gone OR its ProductReleaseIds value has been cleared (ODT empties
    ## this as part of teardown). If neither condition is already true, keep polling.
    $regPath        = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
    $maxWaitSeconds = 5400
    $pollInterval   = 15
    $elapsed        = 0

    While ($elapsed -lt $maxWaitSeconds) {
        If (-not (Test-Path -Path $regPath)) {
            Write-ADTLogEntry -Message "C2R configuration registry key removed -- uninstallation complete."
            Break
        }
        $releaseIds = (Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue).ProductReleaseIds
        If ([String]::IsNullOrWhiteSpace($releaseIds)) {
            Write-ADTLogEntry -Message "C2R ProductReleaseIds cleared -- uninstallation complete (config key remains but Office is removed)."
            Break
        }
        Write-ADTLogEntry -Message "Waiting for ODT uninstall to complete... [$elapsed/$maxWaitSeconds seconds elapsed]"
        Start-Sleep -Seconds $pollInterval
        $elapsed += $pollInterval
    }

    If ($elapsed -ge $maxWaitSeconds) {
        Write-ADTLogEntry -Message "Timed out waiting for C2R removal after [$maxWaitSeconds] seconds." -Severity 3
        Close-ADTSession -ExitCode 60002
    }

    If (Test-Path -LiteralPath $xmlUninstallTempPath) {
        Remove-Item -LiteralPath $xmlUninstallTempPath -Force -ErrorAction SilentlyContinue
        Write-ADTLogEntry -Message "Temporary uninstall XML removed."
    }

    ## Mirror the install-side cleanup: ODT only removes the C2R product, leaving Office UWP
    ## stub/provisioned packages behind. Remove them across all profiles so an uninstall fully
    ## cleans the device. Non-fatal -- a failure here should not fail the uninstall.
    Try {
        Remove-OfficeUWPAllUsers
    } Catch {
        Write-ADTLogEntry -Message "Remove-OfficeUWPAllUsers threw during uninstall - continuing. Error: $($_.Exception.Message)" -Severity 2
        Write-ClientLog -Section 'Post-Uninstallation' -TextBlock "Remove-OfficeUWPAllUsers exception during uninstall - continuing: $($_.Exception.Message)" -Severity WARN
    }

    ##*-------------------------------------------
    ##* POST-UNINSTALLATION
    ##*-------------------------------------------

    Remove-TagFile
    Write-ADTLogEntry -Message "Microsoft 365 Apps uninstallation phase completed."
    Write-ClientLog -Section 'Post-Uninstallation' -TextBlock "Uninstallation complete. Tag file removed." 
}

##*===============================================
##* CLOSE SESSION
##*===============================================
Close-ADTSession
