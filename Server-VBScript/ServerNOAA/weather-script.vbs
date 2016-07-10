'===================================================
' Kindle Weather (for Windows servers)
'
' James Blatchford
' v1.0 29/1/2013
'
' Based on the original Python script by Matthew Petroff
' http://www.mpetroff.net/archives/2012/09/14/kindle-weather-display/
'===================================================

option explicit

'Fetch data
Dim objXMLDoc, objFSO, Root, xmlDayOne
Dim SVGTemplateFileName, SVGOutputFileName, objSVGTemplate, inputSVG, objSVGOutput, objSVGOutputFile, daysOfWeek
Dim Highs(4), Lows(4), Icons(4)

Set objXMLDoc = CreateObject("Microsoft.XMLDOM") 
objXMLDoc.async = False 
If objXMLDoc.load("http://graphical.weather.gov/xml/SOAP_server/ndfdSOAPclientByDay.php?whichClient=NDFDgenByDay&lat=39.3286&lon=-76.6169&format=24+hourly&numDays=4&Unit=e") Then
	Dim TempList, Elem, values, i, xmlIcons, xmlDays
	Set Root = objXMLDoc.documentElement 
	
	'Parse temperatures
	Set TempList = Root.getElementsByTagName("temperature")
	For Each Elem In TempList
		If Elem.getAttribute("type") = "maximum" Then
			Set values = Elem.getElementsByTagName("value")
			For i=1 To values.length
				Highs(i) = values(i-1).text
			Next
		End If
		If Elem.getAttribute("type") = "minimum" Then
			Set values = Elem.getElementsByTagName("value")
			For i=1 To values.length
				Lows(i) = values(i-1).text
			Next
		End If
	Next
	
	'Parse Icons
	Set xmlIcons = Root.getElementsByTagName("icon-link")
	For i=1 to xmlIcons.length
		Dim urlParts, IconName, IconFileName
		urlParts = Split(xmlIcons(i-1).text,"/")
		IconName = Split(urlParts(6),".")
		IconFileName = IconName(0)
		While IsNumeric(Right(IconFileName,1))
			IconFileName = (Left(IconFileName,Len(IconFileName)-1))
		Wend
		Icons(i) = IconFileName
	Next
	
	'Parse Dates
	Set xmlDays = Root.getElementsByTagName("start-valid-time")
	xmlDayOne = Left(xmlDays(0).text,10)
Else
	'Failed to load file
	Wscript.Quit
End If
Set objXMLDoc = Nothing

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

'Insert Days of week
daysOfWeek = Array("","Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday") 'zero based arrays
InputSVG = Replace (inputSVG, "DAY_THREE",daysOfWeek(WeekDay(DateAdd("d",2,xmlDayOne),vbMonday)))
InputSVG = Replace (inputSVG, "DAY_FOUR",daysOfWeek(WeekDay(DateAdd("d",3,xmlDayOne),vbMonday)))

'Write Output
If Not objFSO.FileExists(SVGOutputFileName) Then
	Set objSVGOutputFile = objFSO.CreateTextFile(SVGOutputFileName)
	objSVGOutputFile.Close
End If
Set objSVGOutput = objFSO.OpenTextFile(SVGOutputFileName,ForWriting)
objSVGOutput.WriteLine InputSVG
objSVGOutput.Close
