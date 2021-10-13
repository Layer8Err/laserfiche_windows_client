# Install/Update Laserfiche Client components

$LFInstallerName = "LfWebOffice110.exe"
$LFDownloadURI = "https://lfxstatic.com/dist/WA/latest/LfWebOffice110.exe"
$LFVersionURI = "https://raw.githubusercontent.com/Layer8Err/laserfiche_windows_client/dev/VER_GEN/current_version.json" # TODO: Change /dev/ to /main/ when ready
$LFTempRoot = "$env:WINDIR/temp/LFInstaller"

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
    $versions = ConvertFrom-Json((Invoke-WebRequest $LFVersionURI).Content) # Get version info available on GitHub
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
    Start-Process -FilePath $sz -ArgumentList "x -o`"$Path`" $Installer" -NoNewWindow -Wait
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
            Write-Warning "Waiting since $startTime, but MaxWait exceeded $MaxWait seconds. Terminating msiexec."
            Stop-Process -Name msiexec -Force -ErrorAction:SilentlyContinue
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
        
        Write-Host "Installing Laserfiche pre-requisites..."
        if (Test-Path "$InstallerRoot\Support\msxml6_x86.msi"){
            Write-Host " * Installing MSXML 6.0 Parser (x86) SP1..."
            msiexec.exe /i ($InstallerRoot + "\Support\msxml6_x86.msi") /qn
        }
        if (Test-Path "$InstallerRoot\Support\msxml6_x64.msi"){
            Write-Host " * Installing MSXML 6.0 Parser (x64) SP1..."
            Wait-Msiexec
            msiexec.exe /i ($InstallerRoot + "\Support\msxml6_x64.msi") /qn
        }
        if (!(Check-Installed -Name 'Microsoft Visual C++ 2015-2019 Redistributable (x86)*')){
            Write-Host " * Installing Microsoft Visual C++ 2019 Redistributable (x86) - 14.28.29913.0..."
            Start-Process -FilePath ($InstallerRoot + "\Support\MSVC2019\VC_redist.x86.exe") -ArgumentList "/install /quiet /norestart" -Wait
        }
        if (!(Check-Installed -Name 'Microsoft Visual C++ 2015-2019 Redistributable (x64)*')){
            Write-Host " * Installing Microsoft Visual C++ 2019 Redistributable (x64) - 14.28.29913.0..."
            Start-Process -FilePath ($InstallerRoot + "\Support\MSVC2019\VC_redist.x64.exe") -ArgumentList "/install /quiet /norestart" -Wait
        }
        if (!(Check-Installed -Name 'Microsoft Edge WebView2 Runtime')){
            Write-Host " * Installing Microsoft Edge Web View 2 Runtime (x64)..."
            Start-Process -FilePath ($InstallerRoot + "\Support\MicrosoftEdgeWebView2RuntimeInstallerX64.exe") -ArgumentList "/silent /install" -Wait
        }
    }
    function Install-SetupLf () {
        # Use SetupLf.exe to install Laserfiche
        $logFolder = $InstallerRoot + "\LFInstall_Log"
        $instArgs = "/silent /-noui /-iacceptlicenseagreement -log $logFolder INSTALLLEVEL=300"
        $killList = @(
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
            Write-Host "Installing Laserfiche from $installerPath"
            $instWorkingDir = $InstallerRoot + '\ClientWeb'
            Start-Process -FilePath $installerPath -ArgumentList $instArgs -WorkingDirectory $instWorkingDir -NoNewWindow -Wait
        } else {
            Write-Error "SetupLf.exe could not be found at $installerPath"
        }
    }
    Install-PreReqs
    Wait-Msiexec # Wait for any running install processes to finish
    Install-SetupLf
}

$LFInfo = Check-LFRequired

if ($LFInfo.Install_Webtools -or $LFInfo.Install_OfficeIntegration){
    Write-Host "Laserfiche is not installed... Beginning Install process..."
    if (Test-Path $LFTempRoot){
        Remove-Item -Recurse -Path $LFTempRoot -Force
        New-Item -Path $LFTempRoot -ItemType Directory | Out-Null
    } else {
        New-Item -Path $LFTempRoot -ItemType Directory | Out-Null
    }
    Download-Laserfiche -Path $LFTempRoot
    Extract-Laserfiche -Installer "$LFTempRoot\$LFInstallerName" -Path $LFTempRoot
    Install-Laserfiche -InstallerRoot $LFTempRoot -LFreqs $LFInfo
}