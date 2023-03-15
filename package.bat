@ECHO off

@REM package application using 7zip
7z a dist.zip main.py README.md requirements.txt setup.bat
