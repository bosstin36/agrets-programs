Attribute VB_Name = "modMain"
Public clientport As String
Public hubport As String
Public authIPS As String
Public accDbPath As String
Public maxclients As String
Public serverssec As Boolean
Public loadacc2ram As Boolean
Public passKey As String
Public stateLog As Integer



Sub main()
Dim Mstring As String
Dim Bstring As String
Dim Cans As String

'On Error GoTo goset
Dim sFile As String
    sFile = App.Path & "\SysSettings.ini"
Open sFile For Input As #1
    While Not EOF(1)
        Line Input #1, Mstring
        If UBound(Split(Mstring, "=")) >= 0 Then
        Bstring = LCase(Split(Mstring, "=")(0))
        Cans = Split(Mstring, "=")(1)
            Select Case Bstring
            Case Is = "clientport":
            clientport = Cans
            Case Is = "serverport":
            hubport = Cans
            Case Is = "authsips":
            authIPS = Cans
            Case Is = "maxclients"
            maxclients = Cans
            Case Is = "dbpath"
            accTablePath = Cans
            Case Is = "ipsecurity"
            serverssec = Cans
            Case Is = "spackettable"
            hub_PcktTbl = Cans
            readRespTable
            Case Is = "cpackettable"
            readClientTable Cans
            Case Is = "accounttable"
            accTablePath = Cans
            Case Is = "loadaccountstoram"
            loadacc2ram = Cans
            Case Is = "hubpasscode"
            passKey = Cans
            End Select
       ElseIf UBound(Split(Mstring, "=")) = 0 Then
            Select Case Mstring
            Case Is = "debugmode"
            SetDebugMode
            End Select
        End If
    Wend
    Close
If valid Then
    frmMain.sckServListen = hubport
    frmMain.sckLisClient = clientport
    frmMain.Show
    Load frmLog
    
Else:
    GoTo goset
End If
Exit Sub

goset:
restart = MsgBox("Error parsing settings file, would you like to recheck your settings?", vbYesNo + vbQuestion, "Parse error")
If restart = 6 Then
    frmSet.Show
ElseIf restart = 7 Then
End
End If
End Sub

Private Sub SetDebugMode()
frmMain.Width = frmMain.Width + 5760
frmMain.cmdDebug.Enabled = True
frmMain.cmdDebug.Visible = True
End Sub

Public Function valid() As Boolean
valid = False
If Len(clientport) <> 0 And _
    Len(hubport) <> 0 And _
    Len(maxclients) <> 0 And _
    Len(accTablePath) <> 0 And _
    initAccs Then
    valid = True
End If

End Function

Public Sub logtext(text As String)
frmLog.txtLog = frmLog.txtLog & Date & "(" & Time & ")" & vbTab & text & vbCrLf
End Sub

