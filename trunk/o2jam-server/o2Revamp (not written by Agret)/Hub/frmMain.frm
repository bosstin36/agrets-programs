VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "O2Emu Hub Server"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4860
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   4860
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdConSys 
      Caption         =   "Connect to System"
      Height          =   615
      Left            =   2400
      TabIndex        =   9
      Top             =   960
      Width           =   2175
   End
   Begin VB.Timer tmrSyncState 
      Left            =   3600
      Top             =   2160
   End
   Begin VB.PictureBox pichook 
      Height          =   555
      Left            =   2520
      ScaleHeight     =   495
      ScaleWidth      =   795
      TabIndex        =   7
      Top             =   2280
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSWinsockLib.Winsock sckSysServer 
      Left            =   840
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Unload"
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton cmdMinimize 
      Caption         =   "To Tray"
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton cmdShowDebug 
      Caption         =   "Load Debug Window"
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Frame frmTools 
      Caption         =   "Tools"
      Height          =   2415
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2175
      Begin VB.CommandButton cmdLog 
         Caption         =   "View Log"
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   1680
         Width           =   1815
      End
      Begin VB.CommandButton cmdCon 
         Caption         =   "Server Settings"
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   1815
      End
      Begin VB.CommandButton cmdQue 
         Caption         =   "Query System"
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1815
      End
   End
   Begin MSWinsockLib.Winsock sckControl 
      Index           =   0
      Left            =   840
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckListen 
      Left            =   240
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CheckBox cmdStartStop 
      Caption         =   "Start Server"
      Enabled         =   0   'False
      Height          =   615
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1680
      Width           =   2175
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim trayi As NOTIFYICONDATA

Private Sub cmdCon_Click()
frmSet.Show
End Sub


Private Sub cmdConSys_Click()
tmrSyncState.enabled = True
tmrSyncState.Interval = 100
If ServState = 3 Then
    tmrSyncState.enabled = False
    tmrSyncState.Interval = 0
    cmdConSys.Caption = "Connect"
    cmdStartStop.Value = False
    cmdStartStop.enabled = False
    ServState = 0
End If
End Sub

Private Sub cmdLog_Click()
frmLog.Show
End Sub

Private Sub cmdMinimize_Click()
    trayi.cbSize = Len(trayi)
    trayi.hWnd = pichook.hWnd 'Link the trayicon to this picturebox
    trayi.uId = 1&
    trayi.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    trayi.ucallbackMessage = WM_LBUTTONDOWN
    trayi.hIcon = frmMain.Icon
    trayi.szTip = "O2emu hub server" & chr$(0)
   Shell_NotifyIcon NIM_ADD, trayi
    'Create the icon
    Me.Hide

End Sub

Private Sub cmdQue_Click()
frmQue.Show
End Sub

Private Sub cmdQuit_Click()
Dim xconfirm As Integer
xconfirm = MsgBox("Confirm Exit?", vbOKCancel + vbQuestion, "Are you sure?")
If xconfirm = 1 Then End
End Sub

Private Sub cmdShowDebug_Click()
frmDebug.Show
End Sub

Private Sub cmdStartStop_Click()
If cmdStartStop.Value = 1 Then
    StartServer
ElseIf cmdStartStop.Value = 0 Then
    StopServer
End If
End Sub


Private Sub pichook_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Msg As String
    Msg = X / Screen.TwipsPerPixelX
    'If Not Msg = 512 Then MsgBox Msg
    If Msg = &H203 Then  'If the user dubbel-clicked on the icon
    trayi.cbSize = Len(trayi)
    trayi.hWnd = pichook.hWnd
    trayi.uId = 1&
    'Delete the icon
    Shell_NotifyIcon NIM_DELETE, trayi
    frmMain.Show
    ElseIf Msg = WM_RBUTTONDOWN Then
        Me.PopupMenu frmTest.cmdGlFile
    End If

End Sub

Private Sub sckControl_Close(Index As Integer)
On Error Resume Next
Unload sckControl(Index)
debugText "Freed socket " & Index & " or it has been disconnected"
End Sub

Private Sub sckControl_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim packet As String
sckControl(Index).GetData packet
MainParse packet, Index
End Sub

Private Sub sckListen_ConnectionRequest(ByVal requestID As Long)
Dim i As Integer
For i = 0 To currentSockNumber
    On Error GoTo useNum
    If Not (sckControl(i).State >= 3 And sckControl(i).State <= 7) And (i <= maxsocket) Then
    currentSockNumber = currentSockNumber + 1
    sckControl(i).Accept requestID
    End If
    
Next i
Exit Sub
On Error GoTo sckError
useNum:
If i >= maxsocket Then Exit Sub
    Load sckControl(i)
    
useSock:
    sckControl(i).Accept requestID
        currentSockNumber = currentSockNumber + 1
Exit Sub
sckError:

    
End Sub

Private Sub sckSysServer_Connect()
ServState = 1
logtext "Connecting to " & sckSysServer.RemoteHost & " at port " & sckSysServer.RemotePort & "..........."
'frmLog.Show
End Sub

Private Sub sckSysServer_DataArrival(ByVal bytesTotal As Long)
Dim packet As String
sckSysServer.GetData packet
SysParse packet
End Sub

Private Sub tmrSyncState_Timer()
On Error Resume Next

Select Case ServState
    
    Case Is = 0
    
        If sckSysServer.State <> sckConnected And sckSysServer.State <> sckConnecting Then
            sckSysServer.Connect SysServerIP, SysServerPort
            cmdConSys.Caption = "Connecting..."
            tmrSyncState.Interval = 3000
        ElseIf sckSysServer.State = sckConnected Then
            ServState = 1
            tmrSyncState.Interval = 1100
            cmdConSys.Caption = "Authenticating.."
        End If
        
    Case Is = 1
    
        If sckSysServer.State = sckConnected Then
        cmdConSys.Caption = "Authenticating.."
            sckSysServer.SendData pckt_AuthMe & SysPasskey & chr(3) & GameLisPort
            tmrSyncState.Interval = 5000
        ElseIf sckSysServer.State <> sckConnected Then
            tmrSyncState.Interval = 1100
            ServState = 4
            cmdConSys.Caption = "Connecting.."
        End If
        
    Case Is = 2
    
        If sckSysServer.State = sckConnected Then
            sckSysServer.SendData pckt_SyncState & chr(3) & ServState & chr(3) & GameLisPort
            tmrSyncState.Interval = 5000
            cmdConSys.Caption = "Ready!"
            cmdStartStop.enabled = True
        ElseIf sckSysServer.State <> sckConnected Then
            cmdStartStop.enabled = False
            ServState = 4
            'sckSysServer.SendData pckt_syncstate & chr(3) & currentSockNumber
            cmdConSys.Caption = "Connecting.."
        End If
        
    Case Is = 3
        
        If sckSysServer.State = sckConnected Then
            cmdStartStop.enabled = True
            sckSysServer.SendData pckt_SyncState & chr(3) & ServState & chr(3) & GameLisPort
            cmdConSys.Caption = "Push here to disconnect"
            tmrSyncState.Interval = 5000
        ElseIf sckSysServer.State <> sckConnected Then
            ServState = 4
            tmrSyncState.Interval = 2500
            cmdStartStop.enabled = False
            cmdConSys.Caption = "Connecting.."
        End If
    
    Case Is = 4
        
        If sckSysServer.State = sckConnected Then
            ServState = 1
            tmrSyncState.Interval = 1100
            cmdConSys.Caption = "Authenticating.."
        ElseIf sckSysServer.State <> sckConnected & sckSysServer.State <> sckConnecting Then
            sckSysServer.Connect SysServerIP, SysServerPort
            cmdStartStop.enabled = False
            tmrSyncState.Interval = 1100
        End If
        
End Select

End Sub
