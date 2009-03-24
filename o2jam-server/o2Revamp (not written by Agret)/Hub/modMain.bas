Attribute VB_Name = "modMain"
Public GameLisPort As String
Public ServerRunning As Boolean
Public currentSockNumber As Integer
Public maxsocket As Integer
Public SysServerIP As String
Public SysServerPort As Long
Public SysPasskey As String
Public ServState As Integer
Public pcktTable As String
Public players As cPlayers
Public Channels As cChannels
Public Rooms As cRooms


    Public pckt_AuthMe As String
    Public pckt_Usr_Ins As String
    Public pckt_Usr_Rem As String
    Public pckt_Sys_Err As String
    Public pckt_SyncState As String
    

'ServState eq serverstate:
'0 : connection timer not activated
'1 : server connected, waiting for authentication
'2 : server connected, ready to start
'3 : server connected, o2jam server started
'4 : connection to system server lost

Public Function srun() As Boolean ' check for empty strings
srun = False
If GameLisPort <> Empty & _
    maxsocket <> 0 & _
    Len(SysServerIP) <> 0 & _
    Len(SysServerPort) <> 0 & _
    Len(SysPasskey) <> 0 _
    Then
    
        srun = True
        'If pcktTable = "" Then srun = False
End If
End Function
Sub Main()
Dim i As Integer
Set Channels = New cChannels
Set players = New cPlayers

For i = 1 To 20
Channels.Add i
Set Channels(i).Rooms = New cRooms
Next i
'On Error GoTo errcheck
Open App.Path & "\CoreSettings.ini" For Input As #1 'start with loading server settings
While Not EOF(1)
    Line Input #1, slinein
        If UBound(Split(slinein, "=")) > 0 Then
        ans = Split(slinein, "=")(1)
            strTmp = Split(LCase(slinein), "=")(0)
            Select Case strTmp
            Case Is = "port":
            GameLisPort = ans 'Main hall socket's port
            Case Is = "maxsocket":
            maxsocket = ans ' maximum socket to load
            Case Is = "sysserverip":
            SysServerIP = ans
            Case Is = "sysserverport":
            SysServerPort = ans
            Case Is = "sysserverpckttbl"
            pcktTable = ans
            InitPcktTbl
            Case Is = "passkey"
            SysPasskey = ans
            End Select
            
        ElseIf UBound(Split(slinein, "=")) = 0 Then
        
            Select Case slinein 'Special Commands, not essential
            Case Is = "ShowDebug"
            LoadDebugWindow
            End Select
            
        End If
Wend
Close

If Not srun Then 'checks all value
MsgBox "Error opening config file '" & App.Path & "\CoreSettings.ini' please run config wizard and reload server.", vbOKOnly + vbCritical, "Parse error"
End
ElseIf srun Then
Load frmMain
frmMain.Show
Set players = New cPlayers
End If
Exit Sub
errcheck:

End Sub


Public Sub StartServer()
On Error GoTo sckError
Dim i As Integer
If frmMain.sckSysServer.State = 7 Then
frmMain.sckListen.LocalPort = GameLisPort
frmMain.sckListen.Listen
ServState = 3
frmMain.sckSysServer.SendData pckt_SyncState & chr(3) & ServState & chr(3) & frmMain.cmdStartStop.Value
frmMain.tmrSyncState.Interval = 100
End If
For i = 0 To 20
Channels.Add i
'Channels(i).enabled = True

Next i
Exit Sub

sckError:
logtext "Error connecting to account service " & SysServerIP & ":" & SysServerPort
frmLog.Show
End Sub

Public Sub StopServer()
frmMain.sckListen.Close
closeAllSocket
ServState = 2
frmMain.tmrSyncState.Interval = 100
'frmMain.tmrSyncState.Enabled = False
End Sub

Public Sub LoadDebugWindow()    'debug window consumes more memory, so don't use it when not needed
frmMain.cmdShowDebug.Visible = True
frmMain.cmdShowDebug.enabled = True
Load frmDebug
End Sub

Public Sub closeAllSocket()
On Error Resume Next
For i = 0 To currentSockNumber
    frmMain.sckControl(i).Close
Next i
End Sub

Public Sub debugText(text As String)
frmDebug.txtDebug = frmDebug.txtDebug & vbCrLf & text
End Sub

Public Sub logtext(text As String)
frmLog.txtLog = frmLog.txtLog & vbCrLf & Date & "(" & Time & ")" & " : " & text
End Sub

'defines hub-system connection
Public Sub InitPcktTbl()
file = pcktTable
debugText "Parsing hub-server response table"

On Error GoTo errRead:
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
    
    debugText "Finished parsing packet table with " & lcount & " lines"
Exit Sub
errRead:
    logtext "Error opening " & file
End Sub
    
Public Sub SysParse(packet As String)
Dim pcktCmd As String
Dim pcktData As String
pcktCmd = Left(packet, 4)
pcktData = Right(packet, Len(packet) - 4)
Select Case pcktCmd
    Case Is = pckt_AuthMe
        If Left(pcktData, 1) = "1" Then
        ServState = 2
        ElseIf Left(pcktData, 1) = 0 Then
        ServState = 5
        frmMain.cmdConSys.Caption = "Authenticate failed!"
        End If
    Case Is = pckt_Usr_Ins
    Dim ID As String
    Dim lvl As Integer
    Dim plName As String
    ID = Split(packet, chr(2) & chr(3))(0)
    lvl = Split(packet, chr(2) & chr(3))(3)
    plName = Split(packet, chr(2) & chr(3))(1)
        players.addUser ID, plName, lvl
End Select
End Sub
