sub objXML()
	'Dim xmlDocument, compValue, sysValue, xmlNewNode
	Dim x,y,xw,yh,arrFileLines()
	ForReading = 1
strIP = "config\ipadress.ini"
objStr1 = "<TABLE border='1'>" & vbCrLf & "<TR>" & vbCrLf & "" & vbCrLf & "<TH><FONT SIZE='' COLOR='#0000FF'>Узел:</FONT></TH>" & vbCrLf & "<TH><FONT SIZE='' COLOR='#0000FF'>IP Adress:</FONT></TH>" & vbCrLf & "</TR>" & vbCrLf
objStr3 = "</TABLE>"
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objIP = objFSO.OpenTextFile(strIP, ForReading)
	n = 0
	Do Until objIP.AtEndOfStream
	Redim Preserve arrFileLines(n)
	arrFileLines(n) = objIP.ReadLine
	arr = Split(arrFileLines(n), "|")
	text0 = arr(0)
	text1 = arr(1)
	objStr2 = "<TR>"& vbCrLf & "<TD><input name='t" & n+1 &"' size='15' value='" & text1 & "'></TD>" & vbCrLf &"<TD><input name='CompName"& n+1 &"' size='15' value='" & text0 & "'></TD>"& vbCrLf &"</TR> "& vbCrLf 
	objStr_new = objStr_new  & objStr2
	n = n + 1
	Loop
	objIP.Close
	Set objIP = Nothing
	objWin.InnerHTML = objStr1 & objStr_new & objStr3
	objSIP.value = arrLength(arrFileLines)
	xw = 310
	yh = 400
	x = (window.screen.width - xw) / 2
	y = (window.screen.height - yh) / 2
	If x < 0 Then x = 0
	If y < 0 Then y = 0
	window.resizeTo xw,yh
	window.moveTo x,y
end sub
Function arrLength(vArray)
	ItemCount = 0
		For ItemIndex = 0 To UBound(vArray)
			If Not(vArray(ItemIndex)) = Empty Then
				ItemCount = ItemCount + 1
			End If
		Next
	arrLength = ItemCount
End Function
sub XMLsave_onclick()
	Dim arrFileLines()
	ForReading = 1
	strIP = "config\ipadress.ini"
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objIP = objFSO.OpenTextFile(strIP, ForReading)
	n = 0
	Do Until objIP.AtEndOfStream
	Redim Preserve arrFileLines(n)
	arrFileLines(n) = objIP.ReadLine
	arr = Split(arrFileLines(n), "|")
	text0 = arr(0)
	text1 = arr(1)
	
	n = n + 1
	Loop
	objIP.Close
	Set objIP = Nothing
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	set c2f=objFSO.OpenTextFile("config\ipadress.ini",2,-1)
	for i = 1 to objSIP.value
	objText = document.getElementByID("CompName" & i).value & "|" & document.getElementByID("t" & i).value & "|" & "ON-OFF" & "|" & "-1"
	c2f.WriteLine objText
	next
	c2f.close
end sub
Sub ForEach
	If pin0.Checked then ForEachTrue
		If Not pin0.Checked then ForEachFalse
End Sub
Sub ForEachTrue
 For Each checkbox In CheckboxOption
  checkbox.Checked = True 
 Next
 End Sub
Sub ForEachFalse
	For Each checkbox In CheckboxOption
		checkbox.Checked = False
	Next
 End Sub
Sub WindowOnLoad

	objXML

End Sub
sub CloseButton_onclick()
XMLsave_onclick()
WindowOnLoad
end sub
