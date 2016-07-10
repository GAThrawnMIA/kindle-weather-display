REM @echo off

del ..\public\weather-script-result.png

cscript /nologo weather-script.vbs
..\inkscape\inkscape weather-script-output.svg --export-png=weather-script-output.png
..\pngcrush -blacken -c 4 -reduce weather-script-output.png ..\public\weather-script-result.png

del weather-script-output.svg
del weather-script-output.png