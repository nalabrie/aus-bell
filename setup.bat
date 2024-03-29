@ECHO off

@REM install/update pip packages
echo [31minstalling required packages with pip[0m
pip install -U --upgrade-strategy eager -r requirements.txt

@REM @REM download ffmpeg 7zip archive to RAM
echo [31mdownloading "ffmpeg-release-full.7z"[0m
curl -# -L "https://www.gyan.dev/ffmpeg/builds/ffmpeg-release-full.7z" -o %TEMP%\ffmpeg.7z

@REM @REM extract ffmpeg.exe and ffplay.exe and delete archive
echo [31mextracting ffmpeg.exe[0m
7z e %TEMP%\ffmpeg.7z ffmpeg.exe ffplay.exe -r -aoa
del %TEMP%\ffmpeg.7z

pause
