Attribute VB_Name = "modtoHub"
Public pckt_AuthMe As String
Public pckt_Usr_Ins As String
Public pckt_Usr_Rem As String
Public pckt_Sys_Err As String
Public pckt_SyncState As String
Public hub_PcktTbl As String
Private Type newSocket
    IP As String
    state As Integer
    ready As Boolean
    servPort As Long
    Index As Integer
End Type
Private hubNumber As Integer
Public socketInfo(0 To 1) As newSocket

'defines hub-system connection
Public Sub readRespTable()
On Error GoTo errRead:
file = hub_PcktTbl
Dim readLine As String
Dim readAns As String
Dim readOption As String
Dim lcount As Integer

    Open file For Input As #4
    While Not EOF(4)
    Line Input #4, readLine
    If Not Left(readLine, 1) = "'" Then
    lcount = lcount + 1
    If UBound(Split(readLine, "=")) = 0 Then Exit Sub
    readOption = Split(readLine, "=")(0)
    readAns = Split(readLine, "=")(1)
        If Len(readAns) = 0 Then
            MsgBox "Error parsing account table at line " & lcount & " near " & readOption, vbOKOnly, "Error"
            Exit Sub
        ElseIf Len(readAns) = 1 Then
            readAns = "000" & readAns
        ElseIf Len(readAns) = 2 Then
            readAns = "00" & readAns
        ElseIf Len(readAns) = 3 Then
            readAns = "0" & readAns
        ElseIf Len(readAns) > 4 Then
                readAns = Right(readAns, 4)
                logtext file & " has packet table format error at line " & lcount & " near " & readOption & " packet length should be exactly 4"
        End If
        
        Select Case LCase(readOption)
        Case Is = "senduser"
        pckt_Usr_Ins = readAns
        Case Is = "closeuser"
        pckt_Usr_Rem = readAns
        Case Is = "hubauth"
        pckt_AuthMe = readAns
        Case Is = "statereport"
        pckt_SyncState = readAns
        End Select
        
    End If
    Wend
    Close #4
errRead:
    logtext "Error opening " & file
End Sub

Public Sub parseServer(packet As String, Index As Integer)
Dim pck_Command As String
Dim pck_Data As String

If Len(packet) > 4 Then
pck_Command = Left(packet, 4)
pck_Data = Right(packet, Len(packet) - 4)
ElseIf Len(packet) > 4 Then
    Exit Sub
End If

Select Case pck_Command
    Case Is = pckt_AuthMe:
    If Split(pck_Data, Chr(3))(0) = passKey Then
        'hubNumber = hubNumber + 1
        If hubNumber > 1 Then
            logtext "Number of supported hubServers has reached!"
            hubNumber = hubNumber - 1
            Exit Sub
        End If
        frmMain.sckAcceptServer(Index).SendData pckt_AuthMe & 1
        logtext "Hub server from IP " & frmMain.sckAcceptServer(Index).RemoteHostIP & " slot #" & Index & " is verified(passkey match)."
        frmMain.sckAcceptServer(Index).Tag = hubNumber
        socketInfo(hubNumber).IP = frmMain.sckAcceptServer(Index).RemoteHostIP
        socketInfo(hubNumber).state = 1
        socketInfo(hubNumber).Index = Index
        hubNumber = hubNumber + 1
    End If
    Case Is = pckt_Usr_Ins:
    Case Is = pckt_SyncState:
    'On Error GoTo errexec
    'If frmMain.sckAcceptServer(Index).Tag = "" Then Exit Sub
    'frmMain.sckAcceptServer(Index).Tag = "1" & Chr(3) & Split(pck_Data, Chr(3))(0)
    If socketInfo(frmMain.sckAcceptServer(Index).Tag).state = 1 Then
    Dim lstate As Integer
    lstate = Split(pck_Data, Chr(3))(1)
        Select Case lstate
        Case Is = 2
        socketInfo(frmMain.sckAcceptServer(Index).Tag).servPort = Split(pck_Data, Chr(3))(2)
        socketInfo(frmMain.sckAcceptServer(Index).Tag).ready = False
            If stateLog = 1 Then logtext "HubServer Report: HubServer on slot #" & Index & " has connected."
        Case Is = 3
            If stateLog = 1 Then logtext "HubServer Report: HubServer on slot #" & Index & " is up and running."
        socketInfo(frmMain.sckAcceptServer(Index).Tag).ready = True
        If hubNumber > 2 Then frmMain.sckServListen.Close
        End Select
    End If
End Select
Exit Sub

errexec:
frmMain.sckAcceptServer(Index).Close
Unload frmMain.sckAcceptServer(Index)
End Sub

Public Sub AuthClose(Index As Integer)
If socketInfo(frmMain.sckAcceptServer(Index).Tag).state = 3 Or socketInfo(frmMain.sckAcceptServer(Index).Tag).state = 2 Then
    hubNumber = hubNumber - 1
    socketInfo(frmMain.sckAcceptServer(Index).Tag).ready = False
End If
End Sub
