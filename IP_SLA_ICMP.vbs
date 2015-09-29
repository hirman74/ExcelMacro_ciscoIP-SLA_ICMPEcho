Dim objRead, StringCMD
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim objShell : Set objShell = CreateObject("WScript.Shell")
Dim sessionName, lossNum, objRead1, objRead2, strLine1, strLine2, completeNum
	sessionName = 0
	lossNum = 0	
Dim livePing, avePING, memoryUSED, memoryFREE, liveMemory, aveMemory, cpuTotal, aveCPU, liveCPU, targetIP, targetSNMP, targetPolling, SLAprocessID, binDir, snmpTimedout, intervalMRTG, fromDate
	livePing = 0
	avePING = 0
	memoryUSED = 0
	memoryFREE = 0
	liveMemory = 0
	aveMemory = 0
	cpuTotal = 0
	aveCPU = 0
	liveCPU = 0
	targetIP = Wscript.Arguments.Item(0)
	'targetIP = "192.168.11.108"
	targetSNMP = Wscript.Arguments.Item(1)
	'targetSNMP = "private"
	targetPolling = Wscript.Arguments.Item(2)
	'targetPolling = 5
	SLAprocessID = Wscript.Arguments.Item(3)
	'SLAprocessID = "15"	
	snmpTimedout = 2
	intervalMRTG = Wscript.Arguments.Item(4)
	localFolder = Wscript.Arguments.Item(5)
	'intervalMRTG = 270 '5 Minutes minus 30 Seconds
	'ActiveSheetName = Wscript.Arguments.Item(6)
	fromDate = Now
	'binDir = ".\MIBS"
	binDir = localFolder & "\MIBS"
	
If len (hour(now)) > 1 Then
	strOutputPath = strOutputPath & "" & hour(now)
Else 
	strOutputPath = strOutputPath & "0" & hour(now)
End If 
If len (minute(now)) > 1 Then
	strOutputPath = strOutputPath & minute(now)
Else 
	strOutputPath = strOutputPath & "0" & minute(now)
End If	
If len (second(now)) > 1 Then
	strOutputPath = strOutputPath & second(now)
Else 
	strOutputPath = strOutputPath & "0" & second(now)
End If
	
On Error Resume Next

Dim collectValues(), n
	n = 0
	ReDim Preserve collectValues (2,n)
'Do Until Abs(DateDiff("s",fromDate,Now)) > Abs(intervalMRTG)
	StringCMD = "cmd /c " & localFolder & "\BIN\snmpwalk.exe -v2c -c " & targetSNMP & " -mAll -M" & binDir & " -r 5 -t " & snmpTimedout & " -Oqs " & targetIP & " rttMonLatestRttOperCompletionTime > " & localFolder & "\" & strOutputPath & "01.txt"
	objShell.Run StringCMD, 0 , True
		Set objRead = objFSO.OpenTextFile(localFolder & "\" & strOutputPath & "01.txt", ForReading, FALSE)
		Do Until objread.AtEndOfStream
				strLine = objRead.Readline
				If Instr (strLine, "rttMonLatestRttOperCompletionTime." & SLAprocessID ) > 0 Then 
					PING = Split(strLine)
					livePing = Abs(PING(1))
				End If
		Loop	
		objRead.close
		objFSO.DeleteFile (localFolder & "\" & strOutputPath & "01.txt")
			If avePING <> 0 and livePing <> 0 Then
				avePING = (livePing + avePING) / 2
			ElseIf avePING = 0 and livePing <> 0 Then
				avePING = livePing
			End If	
	StringCMD = "cmd /c " & localFolder & "\BIN\snmpwalk.exe -v2c -c " & targetSNMP & " -mAll -M" & binDir & " -r 5 -t " & snmpTimedout & " -Oqs " & targetIP & " ciscoMemoryPoolUsed.1 > " & localFolder & "\" & strOutputPath & "02.txt"
	objShell.Run StringCMD, 0 , True
	StringCMD = "cmd /c " & localFolder & "\BIN\snmpwalk.exe -v2c -c " & targetSNMP & " -mAll -M" & binDir & " -r 5 -t " & snmpTimedout & " -Oqs " & targetIP & " ciscoMemoryPoolFree.1 > " & localFolder & "\" & strOutputPath & "03.txt"
	objShell.Run StringCMD, 0 , True
		Set objRead = objFSO.OpenTextFile(localFolder & "\" & strOutputPath & "02.txt", ForReading, FALSE)
		Do Until objread.AtEndOfStream
				strLine = objRead.Readline
				If Instr (strLine, "ciscoMemoryPoolUsed.1" ) > 0 Then 
					memoryUSED = Split(strLine)
				End If
		Loop
		objRead.close
		Set objRead = objFSO.OpenTextFile(localFolder & "\" & strOutputPath & "03.txt", ForReading, FALSE)
		Do Until objread.AtEndOfStream
				strLine = objRead.Readline
				If Instr (strLine, "ciscoMemoryPoolFree.1" ) > 0 Then 
					memoryFREE = Split(strLine)
				End If
		Loop
		objRead.close
		objFSO.DeleteFile (localFolder & "\" & strOutputPath & "02.txt")
		objFSO.DeleteFile (localFolder & "\" & strOutputPath & "03.txt")	
			liveMemory = (Abs(memoryUSED(1)) / (Abs(memoryUSED(1)) + Abs(memoryFREE(1))) ) * 100
			If aveMemory <> 0 Then
				aveMemory = (liveMemory + aveMemory) / 2
			ElseIf aveMemory = 0 Then
				aveMemory = liveMemory
			End If	
	
	StringCMD = "cmd /c " & localFolder & "\BIN\snmpwalk.exe -v2c -c " & targetSNMP & " -mAll -M" & binDir & " -r 5 -t " & snmpTimedout & " -Oqs " & targetIP & " cpmCPUTotal5secRev > " & localFolder & "\" & strOutputPath & "04.txt"
	objShell.Run StringCMD, 0 , True		
		Set objRead = objFSO.OpenTextFile(localFolder & "\" & strOutputPath & "04.txt", ForReading, FALSE)
		Do Until objread.AtEndOfStream
				strLine = objRead.Readline
				If Instr (strLine, "cpmCPUTotal5secRev.1" ) > 0 Then 
					cpuTotal = Split(strLine)
					liveCPU = Abs(cpuTotal(1))
				End If
		Loop
		objRead.close
		objFSO.DeleteFile (localFolder & "\" & strOutputPath & "04.txt")
			If aveCPU <> 0 Then
				aveCPU = (liveCPU + aveCPU) / 2
			ElseIf aveCPU = 0 Then
				aveCPU = liveCPU
			End If
	StringCMD = "cmd /c " & localFolder & "\BIN\snmpwalk.exe -v2c -c " & targetSNMP & " -mAll -M" & binDir & " -r 5 -t " & snmpTimedout & " -Oqs " & targetIP & " rttMonStatsCaptureCompletions > " & localFolder & "\" & strOutputPath & "05.txt"
	objShell.Run StringCMD, 0 , True
	StringCMD = "cmd /c " & localFolder & "\BIN\snmpwalk.exe -v2c -c " & targetSNMP & " -mAll -M" & binDir & " -r 5 -t " & snmpTimedout & " -Oqs " & targetIP & " rttMonStatsCollectTimeouts > " & localFolder & "\" & strOutputPath & "06.txt"
	objShell.Run StringCMD, 0 , True		
			Set objRead1 = objFSO.OpenTextFile(localFolder & "\" & strOutputPath & "05.txt", ForReading, FALSE)
			Set objRead2 = objFSO.OpenTextFile(localFolder & "\" & strOutputPath & "06.txt", ForReading, FALSE)
			Do Until objread1.AtEndOfStream
					strLine1 = objRead1.Readline
					If Instr (strLine1, "rttMonStatsCaptureCompletions." & SLAprocessID ) > 0 Then 
						completeNum = Split(strLine1)
						sessionName = Replace (completeNum(0), "rttMonStatsCaptureCompletions." & SLAprocessID & "." , "")
						sessionName = Replace (sessionName, ".1.1.1", "")
					End If
			Loop
			objRead1.close
			Set objRead1 = objFSO.OpenTextFile(localFolder & "\" & strOutputPath & "05.txt", ForReading, FALSE)
			Do Until objread1.AtEndOfStream
					strLine1 = objRead1.Readline
					If Instr (strLine1, "rttMonStatsCaptureCompletions." & SLAprocessID & "." & sessionName ) > 0 Then 
						completeNum = Split(strLine1)
						Do Until objread2.AtEndOfStream
								strLine2 = objRead2.Readline
								If Instr (strLine2, "rttMonStatsCollectTimeouts." & SLAprocessID & "." & sessionName ) > 0 Then 
									lossNum = Split(strLine2)
								End If
						Loop
					End If
			Loop
			objRead1.close
			objRead2.close
			objFSO.DeleteFile (localFolder & "\" & strOutputPath & "05.txt")
			objFSO.DeleteFile (localFolder & "\" & strOutputPath & "06.txt")
	collectValues (0,n) = sessionName
	collectValues (1,n) = Abs(completeNum(1))
	collectValues (2,n) = Abs(lossNum(1))	
	n = Ubound(collectValues,2)+1
	ReDim Preserve collectValues (2,Ubound(collectValues,2)+1)			
	WScript.Sleep (targetPolling * 1000) 'sleep 
'Loop
Set objWrite = objFSO.OpenTextFile(localFolder & "\ping.txt", ForWriting, TRUE)
	objWrite.write avePING
	objWrite.close
	Set objWrite = objFSO.OpenTextFile(localFolder & "\memory.txt", ForWriting, TRUE)
	objWrite.write aveMemory
	objWrite.close
	Set objWrite = objFSO.OpenTextFile(localFolder & "\cpu.txt", ForWriting, TRUE)
	objWrite.write aveCPU
	objWrite.close
	Dim deltaCOMPLETE, deltaLOSS	
		deltaCOMPLETE = abs(collectValues(1,Ubound(collectValues,2)-1)) - abs(collectValues(1,0))
		deltaLOSS = abs(collectValues(2,Ubound(collectValues,2)-1)) - abs(collectValues(2,0))
		Set objWrite = objFSO.OpenTextFile(localFolder & "\success.txt", ForWriting, TRUE)
			If deltaCOMPLETE = 0 or deltaLOSS = 0 Then
			objWrite.write "100%"
			Else
			objWrite.write (deltaCOMPLETE /(deltaCOMPLETE + deltaLOSS) ) * 100
			End If
		objWrite.close

