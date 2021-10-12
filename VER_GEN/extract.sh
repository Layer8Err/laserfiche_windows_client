#!/bin/bash
# Extract LfWebOffice110.exe
# requires 7-Zip (p7zip will not work)

# Get 7zip that works with this PE
DLURI="https://www.7-zip.org/a/7z2103-linux-x64.tar.xz"
wget -c ${DLURI}
tar -xvf 7z2103-linux-x64.tar.xz
sudo cp 7zz /bin/.
sudo chmod +x /bin/7zz

CWD=$( pwd )

if [[ -d "${CWD}/LfWebOffice110" ]]; then
    echo "${CWD}/LfWebOffice110 already exists"
else
    echo "Creating ${CWD}/LfWebOffice110..."
    mkdir ${CWD}/LfWebOffice110
fi

7zz x ${CWD}/LfWebOffice110.exe -o${CWD}/LfWebOffice110
# 7zr x ${CWD}/LfWebOffice110.exe -r -y -aoa -o${CWD}/LfWebOffice110
# docker run -it -v /mnt/c/temp/dev/tmp:/data crazymax/7zip /bin/sh

# Get archive info
# 7z l ${CWD}/LfWebOffice110.exe

# Get Laserfiche Office Integration version
LFOfficeLongVer=$( strings lfoffice-x64_en.msi | grep "Laserfiche Office Integration 11" | head -n 1 )

# Get Laserfiche Webtools Agent version
LFWebtoolsVer=$( strings lfwebtools.msi | grep "Laserfiche Webtools Agent.exe" | grep "PluginContainerExe" | grep "Laserfiche Webtools Agent.exe" | awk -F 'Laserfiche Webtools Agent.exe' '{print $2}' | awk -F 'PluginContainerExeFile' '{print $1}' )
