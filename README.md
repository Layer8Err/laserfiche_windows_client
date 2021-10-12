# Laserfiche Windows Client
Laserfiche Cloud Office Client Updater/Installer

This is a set of files and scripts used for keeping the
Laserfiche Client for Windows up-to-date.

## Installed Components
The `laserfiche_install.ps1` script will install/update the following Laserfiche components:

* Laserfiche Office Integration
* Laserfiche Webtools Agent

## Why does this repo exist?
Currently the downloaded "`LfWebOffice110.exe`" does not allow
for silent updates or install. If this program does allow for silent updates/install, please submit a pull-request as that would make things a lot easier.

In order to keep the Laserfiche client components up-to-date we need to accomplish the following tasks:

* Determine the currently installed version of Laserfiche
* Determine the latest available version of Laserfiche
* Update or install to make sure that the installed version remains current.

# Clone
You can clone this repo with the command:

```bash
git clonegit@github.com:Layer8Err/laserfiche_windows_client.git
```