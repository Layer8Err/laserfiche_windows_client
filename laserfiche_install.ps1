# Install/Update Laserfiche Client components

$LFInstallerName = "LfWebOffice110.exe"
$LFDownloadURI = "https://lfxstatic.com/dist/WA/latest/LfWebOffice110.exe"
$LFOfficeVer = ""
$LFWebtoolsVer = ""


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

