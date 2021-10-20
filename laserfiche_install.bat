@ECHO OFF
TITLE Installing Laserfiche
CD "%~dp0"
powershell -ExecutionPolicy RemoteSigned -Command .\laserfiche_install.ps1