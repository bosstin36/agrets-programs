Attribute VB_Name = "modfromClient"
'this modules defines client response
Private pckt_loginR As String
Private pckt_loginY As String
Private pckt_loginN As String

Public Sub readClientTable(file As String)
'On Error GoTo fileErr
Dim readLine As String
Dim readOption As String
Dim readAns As String

file = Replace(file, ".\", App.Path & "\")
Open file For Input As #3
While Not EOF(1)
Line Input #3, readLine
    If UBound(Split(readLine, "=")) <> 0 Then
        readOption = Split(readLine, "=")(0)
        readAns = Left(Split(readLine, "=")(1), 2)
        Select Case LCase(readOption)
        Case Is = "login_request":
        pckt_loginR = readAns
        Case Is = "login_success":
        pckt_loginY = readAns
        Case Is = "login_fail":
        pckt_loginN = readAns
        
        Case Else:
        logtext "Invalid packet " & readAns & " server may contain bugs caused by invalid client packet table"
        frmLog.Show
        End Select
    End If
Wend

Exit Sub
fileErr:
MsgBox "Error parsing client packet table, please check file path in configuration", vbCritical + vbOKOnly, "Fatal Error"
End Sub

Public Sub PP_client(packet As String, Index As Integer) ' parse packet for client
Dim pcktCommand As String
Dim pcktData As String
Dim gcode As String

pcktCommand = Left(packet, 4)
pcktData = Right(packet, Len(packet) - 4)

    pckt_loginR = "0001"

Select Case pcktCommand
    Case Is = pckt_loginR:
        Dim xuser As String
        Dim xpass As String
        Dim xPar As String ' parametres
        xuser = Split(pcktData, Chr(1))(0)
        xpass = Split(pcktData, Chr(1))(1)
        If chkAcc(xuser, xpass) Then
            gcode = xuser & "-" & Hour(Time) & "o2emu-only"
            If socketInfo(0).ready And socketInfo(1).ready Then
                xPar = "1 " & socketInfo(0).IP & " " & socketInfo(0).servPort
                xPar = xPar & " " & socketInfo(1).IP & " " & socketInfo(1).servPort
                frmMain.sckAcceptServer(socketInfo(0).Index).SendData encode(gcode) & Chr(2) & Chr(3) & xuser & Chr(2) & Chr(3) & accInfo(xuser)
                frmMain.sckAcceptServer(socketInfo(1).Index).SendData encode(gcode) & Chr(2) & Chr(3) & xuser & Chr(2) & Chr(3) & accInfo(xuser)
            ElseIf socketInfo(1).ready And Not socketInfo(0).ready Then
                xPar = "1 " & socketInfo(1).IP & " " & socketInfo(1).servPort
                xPar = xPar & " " & socketInfo(1).IP & " " & socketInfo(1).servPort
                frmMain.sckAcceptServer(socketInfo(1).Index).SendData encode(gcode) & Chr(2) & Chr(3) & xuser & Chr(2) & Chr(3) & accInfo(xuser)
            ElseIf socketInfo(0).ready And Not socketInfo(1).ready Then
                xPar = "1 " & socketInfo(0).IP & " " & socketInfo(0).servPort
                xPar = xPar & " " & socketInfo(0).IP & " " & socketInfo(0).servPort
                frmMain.sckAcceptServer(socketInfo(0).Index).SendData encode(gcode) & Chr(2) & Chr(3) & xuser & Chr(2) & Chr(3) & accInfo(xuser)
            Else:
                frmMain.sckAcceptClient(Index).SendData "1111" & Chr(3) & "System server not ready"
                Exit Sub
            End If
            xPar = encode(gcode) & " 202.75.43.42:1234 O2Jam " & xPar
            frmMain.sckAcceptClient(Index).SendData "11010" & Chr(3) & xPar
            
        ElseIf Not chkAcc(xuser, xpass) Then
            frmMain.sckAcceptClient(Index).SendData "11011"
        End If
        
End Select


End Sub

