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
	text2 = arr(2)
	text3 = arr(3)
	call PING_S(text1, text0, text2, text3)
	objStr2 = text0 & "|" & text1 & "|" & text2 & "|" & text3
	
	objStr_new =  objStr_new  & objStr2 & vbCrLf
	n = n + 1
	
	Loop
	objIP.Close

	Set objIP = Nothing
	set c2f=objFSO.OpenTextFile("config\ipadress.ini",2,-1)
	c2f.Write objStr_new
	c2f.close

Function PING_S(t, CompName, Dostup, time0)
	if CompName <> "" then
		Set objPing = GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery("select * from Win32_PingStatus where address = '" & CompName & "'")
		For Each objStatus In objPing
			If IsNull(objStatus.StatusCode) or objStatus.StatusCode<>0 Then
				Dostup = "OFF"
				time0 = "-1"
			Else
				Dostup = "ON"
				time0 = objStatus.ResponseTime 'nTIME
			end if
		Next
		
	end if

End Function