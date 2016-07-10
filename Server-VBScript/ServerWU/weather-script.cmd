@echo off
pushd \\Server\Users\James\Sites\kindle-weather\ServerWU

del ..\public\weather-script-resultWU.png

cscript /nologo weather-script.vbs
..\inkscape\inkscape weather-script-output.svg --export-png=weather-script-output.png
..\pngcrush -c 4 -reduce weather-script-output.png ..\public\weather-script-resultWU.png

del weather-script-output.svg
del weather-script-output.png

popd