From:
http://www.mpetroff.net/archives/2012/09/14/kindle-weather-display/

# Windows version of the Kindle Weather Server Scripts (using VBScript)

Rewrite of [Matthew Petroff's Python server scripts](http://www.mpetroff.net/archives/2012/09/14/kindle-weather-display/) in VBScript for a Windows server.

* **[Server-VBScript](https://github.com/GAThrawnMIA/kindle-weather-display/tree/master/Server-VBScript)** - Three folders of server scripts, written in VBScript and batch files for a Windows server, just use them from one folder depending on which weather provider you want to use:
 * [ServerMet](https://github.com/GAThrawnMIA/kindle-weather-display/tree/master/Server-VBScript/ServerMet) - Her Majesty's Government's [Meteorological Office](http://www.metoffice.gov.uk/datapoint) (The Met Office) forecasts. They **only cover the UK** (this is what I'm using on my Kindle).
 * [ServerNOAA](https://github.com/GAThrawnMIA/kindle-weather-display/tree/master/Server-VBScript/ServerNOAA) - the US Government's [NOAA](http://graphical.weather.gov/)'s National Weather Service. The major downside (for me) is that they **only provide forecasts for the USA**.
 * [ServerWU](https://github.com/GAThrawnMIA/kindle-weather-display/tree/master/Server-VBScript/ServerWU) - [Weather Underground](http://www.wunderground.com/?apiref=f0020bb946bdd10a), run from a US University, they **provide worldwide coverage**, and **in multiple languages**.
* **[kindle](https://github.com/GAThrawnMIA/kindle-weather-display/tree/master/kindle)** - scripts (and dummy image) to be loaded onto the Kindle
* **[server](https://github.com/GAThrawnMIA/kindle-weather-display/tree/master/server)** - Python and shell scripts for a Linux server


---

See here for more info: [GAThrawn: Kindle Weather Display (from Windows)](http://gathrawn.jard.co.uk/2013/06/kindle-weather-display-from-windows.html)
