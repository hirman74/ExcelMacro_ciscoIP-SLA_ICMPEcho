Sub IPSLAcollect()
    Dim sGroup As String
    Dim sDevice As String
    Dim sIP As String
    Dim sType As String
    Dim sError As String
    Dim sReportTime As Date
    
    Dim dataLine As String
    Dim wshShell As Object
    Set wshShell = CreateObject("Wscript.Shell")
    Dim ws As Worksheet
    Set ws = Worksheets("Sheet1") 

    Dim sFName0 As String
    Dim sFName1 As String
    Dim sFName2 As String
    Dim sFName3 As String
    Dim intFNumber As Integer
    Dim lRow As Long
    Dim lColumn As Long
    
    Dim targetIP As String
    Dim targetSNMP As String
    Dim targetPolling As Integer
    Dim SLAprocessID As Integer
    Dim intervalMRTG As Integer
    
    targetIP = Range("B2").Value
    targetSNMP = Range("B3").Value
    targetPolling = Range("B4").Value
    SLAprocessID = Range("B5").Value
    intervalMRTG = Range("B6").Value
    fromDate = Now
    Do Until Abs(DateDiff("s", fromDate, Now)) > Abs(intervalMRTG)
    lRow = 3
    lColumn = 5
    Do While ws.Cells(lRow, lColumn).Value <> ""
        'lColumn = lColumn + 1
        lRow = lRow + 1
    Loop
            ws.Cells(lRow, lColumn - 1).Value = Format(Now, "dd-mm-yyyy hh:mm:ss AM/PM")
        Set Proc = wshShell.Exec("C:\Windows\System32\wscript.exe " & """" & Application.ActiveWorkbook.Path & "\IP_SLA_ICMP.vbs" & """" & " " & targetIP & " " & targetSNMP & " " & targetPolling & " " & SLAprocessID & " " & intervalMRTG & " " & """" & Application.ActiveWorkbook.Path & """" & " " & """")
            Do While Proc.Status = 0
                Application.Wait Now + TimeValue("00:00:2")
                DoEvents
            Loop
            ws.Cells(lRow, lColumn + 4).Value = Format(Now, "dd-mm-yyyy hh:mm:ss AM/PM")
            'The full path of the text file that will be opened
            If FileFolderExists(Application.ActiveWorkbook.Path & "\ping.txt") And FileFolderExists(Application.ActiveWorkbook.Path & "\memory.txt") And FileFolderExists(Application.ActiveWorkbook.Path & "\cpu.txt") And FileFolderExists(Application.ActiveWorkbook.Path & "\success.txt") Then
                sFName0 = Application.ActiveWorkbook.Path & "\cpu.txt"
                sFName1 = Application.ActiveWorkbook.Path & "\memory.txt"
                sFName2 = Application.ActiveWorkbook.Path & "\ping.txt"
                sFName3 = Application.ActiveWorkbook.Path & "\success.txt"
               'Get an unused file number
                intFNumber = FreeFile
                    'Prepare text file for reading
                    Open sFName0 For Input As #intFNumber
        
                    'Loop until the end of file
                    Do While Not EOF(intFNumber)
                        'Read data from the text file
                        Line Input #intFNumber, dataLine
                        'tmp = Split(dataLine, ",")
                        'Write selected data to the worksheet
                        With ws
                            .Cells(lRow, lColumn).Value = dataLine
                        End With
                        'Address next row of worksheet
                    Loop
              
                    'Close the text file
                    Close #intFNumber
                    Open sFName1 For Input As #intFNumber
                    Do While Not EOF(intFNumber)
                        Line Input #intFNumber, dataLine
                        With ws
                            .Cells(lRow, lColumn + 1).Value = dataLine
                        End With
                    Loop
                    Close #intFNumber
                    Open sFName2 For Input As #intFNumber
                    Do While Not EOF(intFNumber)
                        Line Input #intFNumber, dataLine
                        With ws
                            .Cells(lRow, lColumn + 2).Value = dataLine
                        End With
                    Loop
                    Close #intFNumber
                    Open sFName3 For Input As #intFNumber
                    Do While Not EOF(intFNumber)
                        Line Input #intFNumber, dataLine
                        With ws
                            .Cells(lRow, lColumn + 3).Value = dataLine
                        End With
                    Loop
                    Close #intFNumber
                
                    Kill (Application.ActiveWorkbook.Path & "\cpu.txt")
                    Kill (Application.ActiveWorkbook.Path & "\memory.txt")
                    Kill (Application.ActiveWorkbook.Path & "\ping.txt")
                    Kill (Application.ActiveWorkbook.Path & "\success.txt")
             End If
    Loop
     
MsgBox "Done"
End Sub
Public Function FileFolderExists(strFullPath As String) As Boolean
    On Error GoTo EarlyExit
    If Not Dir(strFullPath, vbDirectory) = vbNullString Then FileFolderExists = True
EarlyExit:
    On Error GoTo 0
End Function





