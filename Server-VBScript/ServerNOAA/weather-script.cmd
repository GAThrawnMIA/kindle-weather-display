REM @echo off
cscript /nologo weather-script.vbs

inkscape\inkscape weather-script-output.svg --export-png=weather-script-output.png
pngcrush -c 4 -reduce weather-script-output.png weather-script-result.png
del weather-script-output.png