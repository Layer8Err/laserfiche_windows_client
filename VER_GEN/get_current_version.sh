#!/bin/bash
# Get the current version of the Laserfiche client components
# Extract LfWebOffice110.exe
# requires 7-Zip (p7zip will not work)
CWD=$( pwd )
WORKINGDIR=${CWD}
LFOfficeInstallerPath=${CWD}/LfWebOffice110.exe
LFDownloadURI="https://lfxstatic.com/dist/WA/latest/LfWebOffice110.exe"


# Get 7zip that works with LfWebOffice PE
Install7Zip() {
    DLURI="https://www.7-zip.org/a/7z2103-linux-x64.tar.xz"
    if [[ -f "/usr/bin/7zz" ]]; then
        echo "7-Zip has already been installed"
    else
        echo "Installing 7-Zip..."
        curl -A "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/94.0.4606.81 Safari/537.36 Edg/94.0.992.47" \
            -o /tmp/7z2103-linux-x64.tar.xz ${DLURI}
        cd /tmp
        tar -xvf /tmp/7z2103-linux-x64.tar.xz -C /tmp
        sudo cp /tmp/7zz /usr/bin/.
        sudo chmod +x /usr/bin/7zz
    fi
}

DownloadLF() {
    curl -A "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/94.0.4606.81 Safari/537.36 Edg/94.0.992.47" \
        -o ${LFOfficeInstallerPath} ${LFDownloadURI}
}

if [[ -f "/usr/bin/7zz" ]]; then
    echo "Found 7-Zip!"
else
    echo "7-Zip is not installed. Downloading and Installing..."
    Install7Zip
fi

if [[ -d "${WORKINGDIR}/LfWebOffice" ]]; then
    echo "${WORKINGDIR}/LfWebOffice already exists"
else
    echo "Creating ${WORKINGDIR}/LfWebOffice..."
    mkdir ${WORKINGDIR}/LfWebOffice
fi

if [[ -f "${LFOfficeInstallerPath}" ]]; then
    echo "Removing old ${LFOfficeInstallerPath}"
    rm ${LFOfficeInstallerPath}
    echo "Downloading latest Laserfiche installer..."
    DownloadLF
else
    echo "Downloading latest Laserfiche installer..."
    DownloadLF
fi

if [[ -f "${LFOfficeInstallerPath}" ]]; then
    7zz x ${LFOfficeInstallerPath} -o${WORKINGDIR}/LfWebOffice
fi
# 7zr x ${CWD}/LfWebOffice110.exe -r -y -aoa -o${CWD}/LfWebOffice
# docker run -it -v /mnt/c/temp/dev/tmp:/data crazymax/7zip /bin/sh << won't work, we need 7zz not p7zip

#####
# Get Laserfiche file versions from MSI installers
if [[ -d "${WORKINGDIR}/LfWebOffice/ClientWeb" ]]; then
    cd ${WORKINGDIR}/LfWebOffice/ClientWeb

    # Get Laserfiche Office Integration version
    LFOfficeLongVer=$( strings lfoffice-x64_en.msi | grep "Laserfiche Office Integration 11" | head -n 1 )
    LFOfficeMainVer=$( echo $LFOfficeLongVer | awk -F 'Laserfiche Office Integration' '{print $2}' | awk -F ' build ' '{print $1}' )
    LFOfficeBuildVer=$( echo $LFOfficeLongVer | awk -F 'Laserfiche Office Integration' '{print $2}' | awk -F ' build ' '{print $2}' )
    LFOfficeVer=$( echo "${LFOfficeMainVer}.${LFOfficeBuildVer}" | xargs )

    # Get Laserfiche Webtools Agent version
    LFWebtoolsVer=$( strings lfwebtools.msi | grep "Laserfiche Webtools Agent.exe" | grep "PluginContainerExe" | grep "Laserfiche Webtools Agent.exe" | awk -F 'Laserfiche Webtools Agent.exe' '{print $2}' | awk -F 'PluginContainerExeFile' '{print $1}' | xargs )

    #####
    # Build current_version.json with version info
    echo -n '{"Laserfiche Office Integration":"'${LFOfficeVer}'","Laserfiche Webtools Agent":"'${LFWebtoolsVer}'"}' > ${WORKINGDIR}/current_version.json
else
    echo "ERROR: No extracted Laserfiche files detected"
fi

#####
# Cleanup leftover download file
if [[ -f "${LFOfficeInstallerPath}" ]]; then
    echo "Removing ${LFOfficeInstallerPath}..."
    rm ${LFOfficeInstallerPath}
fi
if [[ -d "${WORKINGDIR}/LfWebOffice" ]]; then
    echo "Removing leftover extracted files..."
    rm -r "${WORKINGDIR}/LfWebOffice"
fi

cd ${CWD}
