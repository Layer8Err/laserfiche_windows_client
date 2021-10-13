# Install/Update Laserfiche Client components

$LFInstallerName = "LfWebOffice110.exe"
$LFDownloadURI = "https://lfxstatic.com/dist/WA/latest/LfWebOffice110.exe"
$LFVersionURI = "https://raw.githubusercontent.com/Layer8Err/laserfiche_windows_client/dev/VER_GEN/current_version.json" # TODO: Change /dev/ to /main/ when ready


function Download-Laserfiche ($Path="") {
    ## Download LF installer
    $DownloadName = "LfWebOffice110.exe"
    $DownloadURI = "https://lfxstatic.com/dist/WA/latest/" + $DownloadName
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

function Check-LFRequired ($Verbose=$true) {
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
        if ($CurrentLFOfficeVer -eq 0){
            if ($Verbose){
                Write-Warning "Laserfiche Office Integration is not installed"
            }
            $LFReqs.Install_OfficeIntegration = $true
        } else {
            if ($Verbose){
                Write-Warning "Laserfiche Office Integration is version $CurrentLFOfficeVer, but version $LFOfficeVer exists"
            }
            $LFReqs.Upgrade_OfficeIntegration = $false
        }
    } else {
        if ($Verbose){
            Write-Host "Laserfiche Office Integration is already up-to-date (version: $CurrentLFOfficeVer)."
        }
    }
    if ($LFOfficeVer -ne $CurrentLFOfficeVer){
        if ($CurrentLFOfficeVer -eq 0){
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

