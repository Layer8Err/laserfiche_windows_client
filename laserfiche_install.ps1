# Install/Update Laserfiche Client components

$LFInstallerName = "LfWebOffice110.exe"
$LFDownloadURI = "https://lfxstatic.com/dist/WA/latest/LfWebOffice110.exe"
$LFVersionURI = "https://raw.githubusercontent.com/Layer8Err/laserfiche_windows_client/main/VER_GEN/current_version.json"
$LFTempRoot = "$env:WINDIR\temp\LFInstaller"

function Download-Laserfiche ($Path="") {
    ## Download LF installer
    $DownloadName = $LFInstallerName
    $DownloadURI = $LFDownloadURI
    $rootFolder = $PSScriptRoot
    if ($rootFolder.Length -lt 1){
        $rootFolder = $pwd.Path
        if ($rootFolder -match "Microsoft.PowerShell"){
            $rootFolder = $rootFolder.Split('::')[2]
        }
    }
    if ($Path -ne ""){
        if (Test-Path $Path){
            $rootFolder = $Path
        }
    }
    $tgtFilePath = $rootFolder + "/" + $DownloadName
    $client= New-Object System.Net.WebClient
    $client.DownloadFile($DownloadURI, $tgtFilePath)
}

function Check-LFRequired ($Verbose=$false) {
    # Check if Laserfiche is installed and whether or not it is the current version
    $installedSoftware = ( Get-Package -Provider Programs -IncludeWindowsInstaller | Select-Object * )
    $versions = ConvertFrom-Json((New-Object System.Net.WebClient).DownloadString($LFVersionURI)) # Get version info available on GitHub
    $LFOfficeVer = $versions."Laserfiche Office Integration"
    $LFWebtoolsVer = $versions."Laserfiche Webtools Agent"
    $CurrentLFOfficeVer = ($installedSoftware | Where-Object Name -Match "Laserfiche Office Integration").Version
    $CurrentLFWebtoolsVer = ($installedSoftware | Where-Object Name -Match "Laserfiche Webtools Agent").Version
    $LFReqs = [PSCustomObject]@{
        "Latest_OfficeIntegration_Ver" = $LFOfficeVer
        "Current_OfficeIntegration_Ver" = $CurrentLFOfficeVer
        "Upgrade_OfficeIntegration" = $false
        "Install_OfficeIntegration" = $false
        "Latest_WebtoolsAgent_Ver" = $LFWebtoolsVer
        "Current_WebtoolsAgent_Ver" = $CurrentLFWebtoolsVer
        "Upgrade_Webtools" = $false
        "Install_Webtools" = $false
    }
    if ($LFOfficeVer -ne $CurrentLFOfficeVer){
        if ($CurrentLFOfficeVer.Length -eq 0){
            if ($Verbose){
                Write-Warning "Laserfiche Office Integration is not installed"
            }
            $LFReqs.Install_OfficeIntegration = $true
        } else {
            if ($Verbose){
                Write-Warning "Laserfiche Office Integration is version $CurrentLFOfficeVer, but version $LFOfficeVer exists"
            }
            $LFReqs.Upgrade_OfficeIntegration = $true
        }
    } else {
        if ($Verbose){
            Write-Host "Laserfiche Office Integration is already up-to-date (version: $CurrentLFOfficeVer)."
        }
    }
    if ($LFOfficeVer -ne $CurrentLFOfficeVer){
        if ($CurrentLFOfficeVer.Length -eq 0){
            if ($Verbose){
                Write-Warning "Laserfiche Webtools Agent is not installed"
            }
            $LFReqs.Install_Webtools = $true
        } else {
            if ($Verbose){
                Write-Warning "Laserfiche Webtools Agent is version $CurrentLFWebtoolsVer, but version $LFWebtoolsVer exists"
            }
            $LFReqs.Upgrade_Webtools = $true
        }
    } else {
        if ($Verbose){
            Write-Host "Laserfiche Webtools Agent is already up-to-date     (version: $CurrentLFWebtoolsVer)."
        }
    }
    return $LFReqs
}

function Extract-Laserfiche () {
    # Extract a downloaded LfWebOffice110.exe
    # $Installer : should be full path to LfWebOffice110.exe installer
    # $Path : should be the target path that we want to extract Laserfiche to
    Param (
        [Parameter(Mandatory = $true)] [string]$Installer,
        [ValidateNotNullOrEmpty()] [string]$Path = "$env:LOCALAPPDATA\Temp\LFInstaller"
    )
    if (!(Test-Path $Path)){
        New-Item -Path $Path -ItemType Directory
    }
    function Get-7ZipPath($Verbose=$false) {
        # Determine where/if 7-Zip is installed
        $installedSoftware = ( Get-Package -Provider Programs -IncludeWindowsInstaller | Select-Object * )
        $sz = "Required"
        if ((($installedSoftware | Where-Object Name -match "7-Zip").Name).Length -ne 0){
            if ($Verbose){
                Write-Host "7-Zip is currently installed"
            }
            if (Test-Path "$env:ProgramFiles\7-Zip"){
                if (Test-Path "$env:ProgramFiles\7-Zip\7z.exe"){
                    $sz = "$env:ProgramFiles\7-Zip\7z.exe"
                } else {
                    Write-Error "7-Zip folder exists, but is missing 7z.exe"
                }
            } elseif (Test-Path "${env:ProgramFiles(x86)}\7-Zip"){
                if (Test-Path "${env:ProgramFiles(x86)}\7-Zip\7z.exe"){
                    $sz = "${env:ProgramFiles(x86)}\7-Zip\7z.exe"
                } else {
                    Write-Error "7-Zip folder exists, but is missing 7z.exe"
                }
            }
        } else {
            Write-Error "7-Zip has not been installed and is required"
        }
        return $sz
    }
    $sz = Get-7ZipPath
    Start-Process -FilePath $sz -ArgumentList "x -aoa -o`"$Path`" $Installer" -NoNewWindow -Wait
}

function Install-Laserfiche () {
    # Install or Upgrade Laserfiche
    Param (
        [Parameter(Mandatory = $true)] [string]$InstallerRoot,
        [ValidateNotNullOrEmpty()] $LFreqs
    )
    $installedSoftware = ( Get-Package -Provider Programs -IncludeWindowsInstaller | Select-Object * )

    function Wait-Msiexec ($MaxWait=300) {
        $startTime = Get-Date
        $forceEndTime = (Get-Date).AddSeconds($MaxWait)
        while (((Get-Process -Name msiexec -ErrorAction:SilentlyContinue).Length -ne 0) -and ((Get-Date) -lt $forceEndTime)){
            Start-Sleep -Seconds 5 # Backoff 5 seconds if an installer is still running
        }
        if ((Get-Process -Name msiexec -ErrorAction:SilentlyContinue).Length -ne 0) {
            # Assume that msiexec.exe is crashed/hung
            # Write-Warning "Waiting since $startTime, but MaxWait exceeded $MaxWait seconds. Terminating msiexec."
            Get-Process -Name msiexec | Stop-Process -Force
        }
    }
    function Install-PreReqs () {
        # Install programs required by Laserfiche
        function Check-Installed ($Name) {
            # Check if a specific software package is already installed
            $newName = $Name
            if ($Name -match '\++'){
                $newName = $Name.Replace('++', '\++')
            }
            $isInstalled = $false
            $installedSoftware | ForEach-Object {
                $checkName = $_.Name
                if ($checkName -match $newName) {
                    $isInstalled = $true
                }
            }
            return $isInstalled
        }
        
        # MSXML 6.0 should be in $env:WINDIR\System32\Msxml6.dll
        # if (Test-Path "$InstallerRoot\Support\msxml6_x86.msi"){
        #     Write-Host " * Installing MSXML 6.0 Parser (x86) SP1..."
        #     msiexec.exe /i ($InstallerRoot + "\Support\msxml6_x86.msi") /qn
        # }
        if (Test-Path "$InstallerRoot\Support\msxml6_x64.msi"){
            if (!(Test-Path "$env:WINDIR\System32\msxml6.dll")){
                Write-Host " * Installing MSXML 6.0 Parser (x64) SP1..." -ForegroundColor Cyan
                Wait-Msiexec -MaxWait 5
                msiexec.exe /i ($InstallerRoot + "\Support\msxml6_x64.msi") /qn
            }
        }
        if (!(Check-Installed -Name 'Microsoft Visual C++ 2015-2019 Redistributable (x86)*')){
            Write-Host " * Installing Microsoft Visual C++ 2019 Redistributable (x86) - 14.28.29913.0..." -ForegroundColor Cyan
            Start-Process -FilePath ($InstallerRoot + "\Support\MSVC2019\VC_redist.x86.exe") -ArgumentList "/install /quiet /norestart" -Wait
        }
        if (!(Check-Installed -Name 'Microsoft Visual C++ 2015-2019 Redistributable (x64)*')){
            Write-Host " * Installing Microsoft Visual C++ 2019 Redistributable (x64) - 14.28.29913.0..." -ForegroundColor Cyan
            Start-Process -FilePath ($InstallerRoot + "\Support\MSVC2019\VC_redist.x64.exe") -ArgumentList "/install /quiet /norestart" -Wait
        }
        if (!(Check-Installed -Name 'Microsoft Edge WebView2 Runtime')){
            Write-Host " * Installing Microsoft Edge Web View 2 Runtime (x64)..." -ForegroundColor Cyan
            Start-Process -FilePath ($InstallerRoot + "\Support\MicrosoftEdgeWebView2RuntimeInstallerX64.exe") -ArgumentList "/silent /install" -Wait
        }
    }
    function Install-SetupLf () {
        # Use SetupLf.exe to install Laserfiche
        $logFolder = $InstallerRoot + "\LFInstall_Log"
        $instArgs = "/silent /-noui /-iacceptlicenseagreement -log $logFolder INSTALLLEVEL=300"
        $killList = @(
            'ChromeNativeMessaging',
            'Laserfiche.OfficeMonitor',
            'Laserfiche Webtools Agent',
            'SetupLf',
            'msiexec',
            'EXCEL',
            'WINWORD',
            'POWERPNT'
        ) # Programs that cannot be running while installing Laserfiche
        $installerPath = $InstallerRoot + "\ClientWeb\SetupLf.exe"
        if (Test-Path $installerPath){
            $killList | ForEach-Object {
                Stop-Process -Name $_ -Force -ErrorAction:SilentlyContinue
            }
            $instWorkingDir = $InstallerRoot + '\ClientWeb'
            Start-Process -FilePath $installerPath -ArgumentList $instArgs -WorkingDirectory $instWorkingDir -NoNewWindow -Wait
        } else {
            Write-Error "SetupLf.exe could not be found at $installerPath"
        }
    }
    Write-Host "Installing Laserfiche..." -ForegroundColor Cyan
    Install-PreReqs
    Wait-Msiexec # Wait for any running install processes to finish
    Install-SetupLf
}

function Update-Laserfiche () {
    # Update Laserfiche
    Param (
        [Parameter(Mandatory = $true)] [string]$InstallerRoot,
        [ValidateNotNullOrEmpty()] $LFreqs
    )
    $installedSoftware = ( Get-Package -Provider Programs -IncludeWindowsInstaller | Select-Object * )

    function Wait-Msiexec ($MaxWait=300) {
        $startTime = Get-Date
        $forceEndTime = (Get-Date).AddSeconds($MaxWait)
        while (((Get-Process -Name msiexec -ErrorAction:SilentlyContinue).Length -ne 0) -and ((Get-Date) -lt $forceEndTime)){
            Start-Sleep -Seconds 5 # Backoff 5 seconds if an installer is still running
        }
        if ((Get-Process -Name msiexec -ErrorAction:SilentlyContinue).Length -ne 0) {
            # Assume that msiexec.exe is crashed/hung
            # Write-Warning "Waiting since $startTime, but MaxWait exceeded $MaxWait seconds. Terminating msiexec."
            Get-Process -Name msiexec | Stop-Process -Force
        }
    }
    function Install-PreReqs () {
        # Install programs required by Laserfiche
        function Check-Installed ($Name) {
            # Check if a specific software package is already installed
            $newName = $Name
            if ($Name -match '\++'){
                $newName = $Name.Replace('++', '\++')
            }
            $isInstalled = $false
            $installedSoftware | ForEach-Object {
                $checkName = $_.Name
                if ($checkName -match $newName) {
                    $isInstalled = $true
                }
            }
            return $isInstalled
        }
        
        # MSXML 6.0 should be in $env:WINDIR\System32\Msxml6.dll
        # if (Test-Path "$InstallerRoot\Support\msxml6_x86.msi"){
        #     Write-Host " * Installing MSXML 6.0 Parser (x86) SP1..."
        #     msiexec.exe /i ($InstallerRoot + "\Support\msxml6_x86.msi") /qn
        # }
        if (Test-Path "$InstallerRoot\Support\msxml6_x64.msi"){
            if (!(Test-Path "$env:WINDIR\System32\msxml6.dll")){
                Write-Host " * Installing MSXML 6.0 Parser (x64) SP1..." -ForegroundColor Cyan
                Wait-Msiexec -MaxWait 5
                msiexec.exe /i ($InstallerRoot + "\Support\msxml6_x64.msi") /qn
            }
        }
        #TODO: Add check and install C++ 2012
        # ECHO  * Installing Microsoft Visual C++ 2012 Redistributable (x86) - 11.0.61030 ...
        # vc_redist-2012_x86.exe /q
        # ECHO  * Installing Microsoft Visual C++ 2012 Redistributable (x64) - 11.0.61030 ...
        # vc_redist-2012_x64.exe /q
        if (!(Check-Installed -Name 'Microsoft Visual C++ 2015-2019 Redistributable (x86)*')){
            Write-Host " * Installing Microsoft Visual C++ 2019 Redistributable (x86) - 14.28.29913.0..." -ForegroundColor Cyan
            Start-Process -FilePath ($InstallerRoot + "\Support\MSVC2019\VC_redist.x86.exe") -ArgumentList "/install /quiet /norestart" -Wait
        }
        if (!(Check-Installed -Name 'Microsoft Visual C++ 2015-2019 Redistributable (x64)*')){
            Write-Host " * Installing Microsoft Visual C++ 2019 Redistributable (x64) - 14.28.29913.0..." -ForegroundColor Cyan
            Start-Process -FilePath ($InstallerRoot + "\Support\MSVC2019\VC_redist.x64.exe") -ArgumentList "/install /quiet /norestart" -Wait
        }
        if (!(Check-Installed -Name 'Microsoft Edge WebView2 Runtime')){
            Write-Host " * Installing Microsoft Edge Web View 2 Runtime (x64)..." -ForegroundColor Cyan
            Start-Process -FilePath ($InstallerRoot + "\Support\MicrosoftEdgeWebView2RuntimeInstallerX64.exe") -ArgumentList "/silent /install" -Wait
        }
    }
    function Update-Lf () {
        # Use lfoffice-x64_en.msi and lfwebtools.msi to update an existing Laserfiche install
        $killList = @(
            'ChromeNativeMessaging',
            'Laserfiche.OfficeMonitor',
            'Laserfiche Webtools Agent',
            'SetupLf',
            'msiexec',
            'EXCEL',
            'WINWORD',
            'POWERPNT'
        ) # Programs that cannot be running while installing Laserfiche
        $lfoffice = $InstallerRoot + "\ClientWeb\lfoffice-x64_en.msi"
        $lfwebtools = $InstallerRoot + "\ClientWeb\lfwebtools.msi"
        if ((Test-Path $lfoffice) -and (Test-Path $lfwebtools)){
            $killList | ForEach-Object {
                Stop-Process -Name $_ -Force -ErrorAction:SilentlyContinue
            }
            $lfofficeArgs = 'REBOOT=ReallySuppress INSTALLDIR="' + $env:ProgramFiles + '\Laserfiche\Client\" INSTALLDIR32="' + ${env:ProgramFiles(x86)} + '\Laserfiche\Client\" ADDLOCAL="LFOffice" REINSTALL=ALL REINSTALLMODE=vomus'
            $lfwebtoolsArgs = 'REBOOT=ReallySuppress INSTALLDIR="' + ${env:ProgramFiles(x86)} + '\Laserfiche\Webtools Agent\" ADDLOCAL="LFWebtools" REINSTALL=ALL REINSTALLMODE=vomus'
            $lfofficeInst = "/qn /i $lfoffice $lfofficeArgs"
            $lfwebtoolsInst = "/qn /i $lfwebtools $lfwebtoolsArgs"
            # msiexec.exe /qn /i $lfoffice $lfofficeArgs
            Write-Host " * Updating Laserfiche Office Integration..."
            Start-Process msiexec.exe -WorkingDirectory ($InstallerRoot + "\ClientWeb") -ArgumentList $lfofficeInst -Wait
            Wait-Msiexec
            Write-Host " * Updating Laserfiche Webtools Agent..."
            Start-Process msiexec.exe -WorkingDirectory ($InstallerRoot + "\ClientWeb") -ArgumentList $lfwebtoolsInst -Wait
        } else {
            Write-Error "Could not find files required to update Laserfiche"
        }
    }
    Write-Host "Updating Laserfiche..." -ForegroundColor Cyan
    Install-PreReqs
    Wait-Msiexec # Wait for any running install processes to finish
    Update-Lf
}

function Remove-Software ($Name) {
    # $installedPrograms = Get-CimInstance Win32_Product
    $installedSoftware = ( Get-Package -Provider Programs -IncludeWindowsInstaller | Select-Object * )
    $targetForRemoval = $installedSoftware | Where-Object Name -match $Name
    if ($targetForRemoval.Length -ne 0) {
        $targetForRemoval | ForEach-Object {
            $xmlTxt = $_.SwidTagText
            [xml]$xml = $xmlTxt
            $UninstallString = $xml.SoftwareIdentity.Meta.UninstallString
            if ($UninstallString -ne $Null) {
                if ($UninstallString.ToLower() -match "msiexec") {
                    if ($UninstallString.ToLower() -match "/i{"){
                        $UninstallString = $UninstallString.Replace("/I{", "/X{")
                        $UninstallString = $UninstallString.Replace("/i{", "/X{")
                    }
                    $UninstallArgs = $UninstallString.Replace("msiexec.exe ", "")
                    $UninstallArgs = $UninstallArgs.Replace("MsiExec.exe ", "")
                    $UninstallArgs += " /qn /norestart"
                    if ((Get-Process -Name msiexec -ErrorAction:SilentlyContinue).Length -gt 0){
                        Stop-Process -Name msiexec -Force # Force kill any running (hung?) msiexec processes
                    }
                    Write-Host "Beginning MsiExec uninstall of $Name..." -ForegroundColor Cyan
                    Write-Host "  msiexec.exe $UninstallArgs" -ForegroundColor White
                    Start-Process "msiexec.exe" -arg "$UninstallArgs" -Wait
                    Start-Sleep -Seconds 5
                } else {
                    Write-Warning "Unknown uninstall method required for uninstall string:"
                    Write-Host "$UninstallString" -ForegroundColor Red
                }
                if ($UninstallString.ToLower() -match "msiexec") {
                    While ((Get-Process -Name msiexec -ErrorAction:SilentlyContinue).Length -gt 0){
                        Start-Sleep -Seconds 5
                    }
                }
            } else {
                Write-Host $_.SwidTagText
                Write-Host $UninstallString
            }
        }
    } else {
        Write-Warning "Existing $Name install not found."
    }
}

function Uninstall-Laserfiche () {
    Remove-Software -Name "Laserfiche Webtools Agent"
    Start-Sleep -Seconds 10
    While ((Get-Process -Name msiexec -ErrorAction:SilentlyContinue).Length -gt 0){
        Start-Sleep -Seconds 10 # Backoff while waiting for "Laserfiche Webtools Agent 10" to uninstall
    }
    Remove-Software -Name "Laserfiche Office Integration"
}

$LFInfo = Check-LFRequired

if ($LFInfo.Install_Webtools -or $LFInfo.Install_OfficeIntegration){
    Write-Host "Laserfiche is not installed... Beginning Install process..." -ForegroundColor Yellow
    if (Test-Path $LFTempRoot){
        Remove-Item -Recurse -Path $LFTempRoot -Force -ErrorAction:SilentlyContinue
        if (!(Test-Path $LFTempRoot)){
            New-Item -Path $LFTempRoot -ItemType Directory | Out-Null
        }
    } else {
        New-Item -Path $LFTempRoot -ItemType Directory | Out-Null
    }
    Write-Host "Downloading latest Laserfiche Office Installer..." -ForegroundColor Cyan
    Download-Laserfiche -Path $LFTempRoot
    Write-Host "Extracting Laserfiche Office Installer for silent install..." -ForegroundColor Cyan
    Extract-Laserfiche -Installer "$LFTempRoot\$LFInstallerName" -Path $LFTempRoot
    Install-Laserfiche -InstallerRoot $LFTempRoot -LFreqs $LFInfo
} elseif ($LFInfo.Upgrade_OfficeIntegration -or $LFInfo.Upgrade_Webtools){
    Write-Host "Laserfiche is not current version.. Beginning Update process..." -ForegroundColor Yellow
    $inPlaceUpgrade = $true
    $LFOMajorVersion = $LFInfo.Latest_OfficeIntegration_Ver.Split('.')[0]
    $LFOCurentMajorVersion = $LFInfo.Current_OfficeIntegration_Ver.Split('.')[0]
    if ($LFOMajorVersion -ne $LFOCurentMajorVersion){ $inPlaceUpgrade = $false }
    if ($inPlaceUpgrade) {
        if (Test-Path $LFTempRoot){
        Remove-Item -Recurse -Path $LFTempRoot -Force -ErrorAction:SilentlyContinue
        if (!(Test-Path $LFTempRoot)){
            New-Item -Path $LFTempRoot -ItemType Directory | Out-Null
        }
        } else {
            New-Item -Path $LFTempRoot -ItemType Directory | Out-Null
        }
        if ($LFInfo.Upgrade_OfficeIntegration){
            Write-Host "Current Laserfiche Office Integration version: $($LFInfo.Current_OfficeIntegration_Ver) [version $($LFInfo.Latest_OfficeIntegration_Ver) available]" -ForegroundColor White
        }
        if ($LFInfo.Upgrade_Webtools){
            Write-Host "Current Laserfiche Webtools Agent version: $($LFInfo.Current_WebtoolsAgent_Ver) [version $($LFInfo.Latest_WebtoolsAgent_Ver) available]" -ForegroundColor White
        }
        Write-Host "Downloading latest Laserfiche Office Installer..." -ForegroundColor Cyan
        Download-Laserfiche -Path $LFTempRoot
        Write-Host "Extracting Laserfiche Office Installer for silent install..." -ForegroundColor Cyan
        Extract-Laserfiche -Installer "$LFTempRoot\$LFInstallerName" -Path $LFTempRoot
        Update-Laserfiche -InstallerRoot $LFTempRoot -LFreqs $LFInfo
    } else {
        Write-Warning "Current Laserfiche version is too old for upgrade... Beginning clean install..."
        Write-Host "Uninstalling old Laserfiche Office components..." -ForegroundColor Cyan
        Uninstall-Laserfiche
        While ((Get-Process -Name msiexec -ErrorAction:SilentlyContinue).Length -gt 0){
            Start-Sleep -Seconds 10 # Backoff while waiting for msiexec processes to finish
        }
        Write-Host "Downloading latest Laserfiche Office Installer..." -ForegroundColor Cyan
        Download-Laserfiche -Path $LFTempRoot
        Write-Host "Extracting Laserfiche Office Installer for silent install..." -ForegroundColor Cyan
        Extract-Laserfiche -Installer "$LFTempRoot\$LFInstallerName" -Path $LFTempRoot
        Install-Laserfiche -InstallerRoot $LFTempRoot -LFreqs $LFInfo
    }
    
} else {
    Write-Host "Laserfiche appears to be installed and up-to-date" -ForegroundColor Green
    Write-Host " * Installed Laserfiche Office Integration version: $($LFInfo.Current_OfficeIntegration_Ver) [latest version $($LFInfo.Latest_OfficeIntegration_Ver)]" -ForegroundColor Cyan
    Write-Host " * Installed Laserfiche Webtools Agent version: $($LFInfo.Current_WebtoolsAgent_Ver) [latest version: $($LFInfo.Latest_WebtoolsAgent_Ver)]" -ForegroundColor Cyan
}
