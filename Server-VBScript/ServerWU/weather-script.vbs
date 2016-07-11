'===================================================
' Kindle Weather (for Windows servers)
' Weather Underground version
'
' James Blatchford
' v1.0 30/1/2013	Initial release
' v1.1 27/2/2014	Fix for dates overflowing available space after WUnderground changed the weekday field to be full name, and added a new weekday_short field
'
' Based on the original Python script by Matthew Petroff
' http://www.mpetroff.net/archives/2012/09/14/kindle-weather-display/
'===================================================

option explicit

'Fetch data
Dim objXMLDoc, objFSO, objRoot, arrForecast, objForecastDays, arrForecastDays
Dim SVGTemplateFileName, SVGOutputFileName, objSVGTemplate, inputSVG, objSVGOutputFile, objSVGOutput, arrDate
Dim Highs(4), Lows(4), Icons(4), Dates(4)

Set objXMLDoc = CreateObject("MSXML.DOMDocument") 
objXMLDoc.async = False 
' EXAMPLE: http://api.wunderground.com/api/APIKEY/forecast/q/United Kingdom/London.xml") Then
If objXMLDoc.load("http://api.wunderground.com/api/APIKEY/forecast/q/COUNTRY/CITY.xml") Then
	Set objRoot = objXMLDoc.documentElement
	Set arrForecast = objRoot.getElementsByTagName("simpleforecast")
	Set objForecastDays = arrForecast(0).firstChild
	Set arrForecastDays = objForecastDays.childNodes

	Set arrDate = arrForecastDays(0).getElementsByTagName("date")(0).childNodes
	Dates(1) = arrDate(14).text & " " & arrDate(3).text & " " & arrDate(12).text
	Highs(1) = arrForecastDays(0).selectSingleNode("high/celsius").text
	Lows(1) = arrForecastDays(0).selectSingleNode("low/celsius").text
	Icons(1) = arrForecastDays(0).selectSingleNode("icon").text
	
	Set arrDate = arrForecastDays(1).getElementsByTagName("date")(0).childNodes
	Dates(2) = arrDate(12).text & " " & arrDate(3).text & " " & arrDate(13).text
	Highs(2) = arrForecastDays(1).selectSingleNode("high/celsius").text
	Lows(2) = arrForecastDays(1).selectSingleNode("low/celsius").text
	Icons(2) = arrForecastDays(1).selectSingleNode("icon").text
	
	Set arrDate = arrForecastDays(2).getElementsByTagName("date")(0).childNodes
	Dates(3) = arrDate(12).text & " " & arrDate(3).text & " " & arrDate(13).text
	Highs(3) = arrForecastDays(2).selectSingleNode("high/celsius").text
	Lows(3) = arrForecastDays(2).selectSingleNode("low/celsius").text
	Icons(3) = arrForecastDays(2).selectSingleNode("icon").text
	
	Set arrDate = arrForecastDays(3).getElementsByTagName("date")(0).childNodes
	Dates(4) = arrDate(12).text & " " & arrDate(3).text & " " & arrDate(13).text
	Highs(4) = arrForecastDays(3).selectSingleNode("high/celsius").text
	Lows(4) = arrForecastDays(3).selectSingleNode("low/celsius").text
	Icons(4) = arrForecastDays(3).selectSingleNode("icon").text
Else
	'Failed to load file
	Wscript.Quit
End If
Set objXMLDoc = Nothing

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
