@echo off
setlocal
title Video to GIF - 5x speed
echo Video to GIF - default: 5x speed, 800 px, 8 fps
echo Color quality: 256 colors per frame, smooth color transitions
echo Double-click to select videos, or drag video files onto this launcher.
echo Output folder: %~dp0assets
echo.
py -3 "%~dp0tools\video_to_gif.py" %*
set "gif_exit_code=%errorlevel%"
echo.
pause
exit /b %gif_exit_code%
