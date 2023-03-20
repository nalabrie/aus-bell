@ECHO off

@REM make version file
git rev-parse --short HEAD > VERSION

@REM package application using 7zip
7z a dist.zip config.json main.py README.md requirements.txt setup.bat start.bat VERSION

pause
