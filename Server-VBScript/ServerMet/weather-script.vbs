'===================================================
' Kindle Weather (for Windows servers)
' Met Office version
'
' James Blatchford
' v1.0 4/2/2013
'
' Based on the original Python script by Matthew Petroff
' http://www.mpetroff.net/archives/2012/09/14/kindle-weather-display/
'===================================================

option explicit

'Fetch data
Dim objXMLDoc, objFSO, objRoot, arrLocation, arrPeriods
Dim SVGTemplateFileName, SVGOutputFileName, objSVGTemplate, inputSVG, objSVGOutputFile, objSVGOutput, arrDate
Dim daysOfWeek, dayNum, ordSuffix
Dim Highs(4), Lows(4), Icons(4), Dates(4), arrWeatherTypes(30)

Set objXMLDoc = CreateObject("MSXML.DOMDocument") 
objXMLDoc.async = False 
'Site list  http://datapoint.metoffice.gov.uk/public/data/val/wxfcs/all/xml/sitelist?key=API_KEY
If objXMLDoc.load("http://datapoint.metoffice.gov.uk/public/data/val/wxfcs/all/xml/350928?res=daily&key=APIKEY") Then
	Set objRoot = objXMLDoc.documentElement
	Set arrLocation = objRoot.getElementsByTagName("Location")
	Set arrPeriods = arrLocation(0).childNodes
	
	Dates(1) = arrPeriods(0).getAttribute("value") '&" "& arrPeriods(0).getAttribute("type")
	Highs(1) = arrPeriods(0).childNodes(0).getAttribute("Dm")
	Lows(1) = arrPeriods(0).childNodes(1).getAttribute("Nm")
	Icons(1) = arrPeriods(0).childNodes(0).getAttribute("W")
	
	Dates(2) = arrPeriods(1).getAttribute("value") '&" "& arrPeriods(0).getAttribute("type")
	Highs(2) = arrPeriods(1).childNodes(0).getAttribute("Dm")
	Lows(2) = arrPeriods(1).childNodes(1).getAttribute("Nm")
	Icons(2) = arrPeriods(1).childNodes(0).getAttribute("W")
	
	Dates(3) = arrPeriods(2).getAttribute("value") '&" "& arrPeriods(0).getAttribute("type")
	Highs(3) = arrPeriods(2).childNodes(0).getAttribute("Dm")
	Lows(3) = arrPeriods(2).childNodes(1).getAttribute("Nm")
	Icons(3) = arrPeriods(2).childNodes(0).getAttribute("W")
	
	Dates(4) = arrPeriods(3).getAttribute("value") '&" "& arrPeriods(0).getAttribute("type")
	Highs(4) = arrPeriods(3).childNodes(0).getAttribute("Dm")
	Lows(4) = arrPeriods(3).childNodes(1).getAttribute("Nm")
	Icons(4) = arrPeriods(3).childNodes(0).getAttribute("W")
Else
	'Failed to load file
	Wscript.Quit
End If
Set objXMLDoc = Nothing

'Weather types from http://www.metoffice.gov.uk/datapoint/support/code-definitions
'NA	Not available
arrWeatherTypes(0) = "Clear night"
arrWeatherTypes(1) = "Sunny day"	'*
arrWeatherTypes(2) = "Partly cloudy (night)"	'*
arrWeatherTypes(3) = "Partly cloudy (day)"	'*
arrWeatherTypes(4) = "Not used"
arrWeatherTypes(5) = "Mist"
arrWeatherTypes(6) = "Fog"	'*
arrWeatherTypes(7) = "Cloudy"	'*
arrWeatherTypes(8) = "Overcast"	'*
arrWeatherTypes(9) = "Light rain shower (night)"	'*
arrWeatherTypes(10) = "Light rain shower (day)"	'*
arrWeatherTypes(11) = "Drizzle"	'*
arrWeatherTypes(12) = "Light rain"	'*
arrWeatherTypes(13) = "Heavy rain shower (night)"	'*
arrWeatherTypes(14) = "Heavy rain shower (day)"	'*
arrWeatherTypes(15) = "Heavy rain"	'*
arrWeatherTypes(16) = "Sleet shower (night)"	'*
arrWeatherTypes(17) = "Sleet shower (day)"	'*
arrWeatherTypes(18) = "Sleet"	'*
arrWeatherTypes(19) = "Hail shower (night)"	'*
arrWeatherTypes(20) = "Hail shower (day)"	'*
arrWeatherTypes(21) = "Hail"	'*
arrWeatherTypes(22) = "Light snow shower (night)"	'*
arrWeatherTypes(23) = "Light snow shower (day)"	'*
arrWeatherTypes(24) = "Light snow"	'*
arrWeatherTypes(25) = "Heavy snow shower (night)"	'*
arrWeatherTypes(26) = "Heavy snow shower (day)"	'*
arrWeatherTypes(27) = "Heavy snow"	'*
arrWeatherTypes(28) = "Thunder shower (night)"	'*
arrWeatherTypes(29) = "Thunder shower (day)"	'*
arrWeatherTypes(30) = "Thunder"	'*
Dim i
'Remove spaces and bracketed words from weather icons
For i=1 To UBound(Icons)
	Icons(i) = Replace(Split(arrWeatherTypes(Icons(i)),"(")(0)," ","")
Next


'Insert day of week and date
daysOfWeek = Array("","Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun") 'zero based arrays
For i=1 To UBound(Dates)
	Dates(i) = Replace(Dates(i),"Z","")
	dayNum = FormatNumber(Split(Dates(i),"-")(2),0,vbFalse)
	Select Case dayNum
		Case 1, 21, 31
			ordSuffix = "st"
		Case 2, 22
			ordSuffix = "nd"
		Case 3, 23
			ordSuffix = "rd"
		Case else
			ordSuffix = "th"
	End select
	Dates(i) = daysOfWeek(WeekDay(Dates(i),vbMonday)) & " " & dayNum & ordSuffix
Next

'Debugging data output
'Dim j
'For Each j in Highs
'	wscript.echo j
'Next
'For Each j in Lows
'	wscript.echo j
'Next
'For Each j in Icons
'	wscript.echo j
'Next
'For Each j in Dates
'	wscript.echo j
'Next

'
' Preprocess SVG
'
'Open SVG to process
Const ForReading = 1
Const ForWriting = 2
SVGTemplateFileName = "weather-script-preprocess.svg"
SVGOutputFileName = "weather-script-output.svg"
Set objFSO = CreateObject("Scripting.FileSystemObject")
If objFSO.FileExists(SVGTemplateFileName) Then
	Set objSVGTemplate = objFSO.OpenTextFile(SVGTemplateFileName,ForReading)
Else
	'File Not Found
	Wscript.Quit
End If
inputSVG = objSVGTemplate.ReadAll
objSVGTemplate.Close

' Insert icons and temperatures
inputSVG = Replace (inputSVG,"ICON_ONE",icons(1))
inputSVG = Replace (inputSVG,"ICON_TWO",icons(2))
inputSVG = Replace (inputSVG,"ICON_THREE",icons(3))
inputSVG = Replace (inputSVG,"ICON_FOUR",icons(4))
inputSVG = Replace (inputSVG,"HIGH_ONE",Highs(1))
inputSVG = Replace (inputSVG,"HIGH_TWO",Highs(2))
inputSVG = Replace (inputSVG,"HIGH_THREE",Highs(3))
inputSVG = Replace (inputSVG,"HIGH_FOUR",Highs(4))
inputSVG = Replace (inputSVG,"LOW_ONE",Lows(1))
inputSVG = Replace (inputSVG,"LOW_TWO",Lows(2))
inputSVG = Replace (inputSVG,"LOW_THREE",Lows(3))
inputSVG = Replace (inputSVG,"LOW_FOUR",Lows(4))
inputSVG = Replace (inputSVG,"DAY_ONE",Dates(1))
inputSVG = Replace (inputSVG,"DAY_TWO",Dates(2))
inputSVG = Replace (inputSVG,"DAY_THREE",Dates(3))
inputSVG = Replace (inputSVG,"DAY_FOUR",Dates(4))

'Write Output
If Not objFSO.FileExists(SVGOutputFileName) Then
	Set objSVGOutputFile = objFSO.CreateTextFile(SVGOutputFileName)
	objSVGOutputFile.Close
End If
Set objSVGOutput = objFSO.OpenTextFile(SVGOutputFileName,ForWriting)
objSVGOutput.WriteLine InputSVG
objSVGOutput.Close
