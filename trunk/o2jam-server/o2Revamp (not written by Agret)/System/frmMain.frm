VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWinsck.OCX"
Begin VB.Form frmMain 
   Caption         =   "System Server Panel"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4080
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdShowLog 
      Caption         =   "View Log"
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton cmdDebug 
      Caption         =   "Debug Window"
      Enabled         =   0   'False
      Height          =   975
      Left            =   4200
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock sckServListen 
      Left            =   2160
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckAcceptServer 
      Index           =   0
      Left            =   1560
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckLisClient 
      Left            =   960
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckAcceptClient 
      Index           =   0
      Left            =   360
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CheckBox cmdStartInter 
      Caption         =   "Start Client Service"
      Height          =   975
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   360
      Width           =   1575
   End
   Begin VB.CheckBox cmdStartAcc 
      Caption         =   "Start Account Service"
      Height          =   975
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton cmdSet 
      Caption         =   "Settings"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton cmdQue 
      Caption         =   "Accounts setup"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   1575
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private currentsocknumber As Integer
Private serversocket As Integer
Private CurServState As Integer

Private Sub cmdSet_Click()
frmSet.Show
End Sub

Private Sub cmdShowLog_Click()
frmLog.Show
End Sub

Private Sub cmdStartAcc_Click()
On Error GoTo err
If cmdStartAcc.Value = 1 Then
Me.sckServListen.LocalPort = hubport
sckServListen.Listen
logtext "Account service started..."
ElseIf cmdStartAcc.Value = False Then
Me.sckServListen.Close
logtext "Account service halted"
End If
Exit Sub
err:
logtext "Account service error, try chaning hubservice port"
frmLog.Show
End Sub

Private Sub cmdStartInter_Click()
On Error GoTo err:
If cmdStartInter.Value = 1 Then
    sckLisClient.LocalPort = clientport
    sckLisClient.Listen
    logtext "LauncherClientServer service started"
ElseIf cmdStartInter.Value = 0 Then
    sckLisClient.Close
    closesocks
    logtext "LauncherClientServer halted"
End If
Exit Sub
err:
    logtext "Error creating socket!, try changing client service port"
    frmLog.Show
End Sub



Private Sub sckAcceptClient_Close(Index As Integer)
Unload sckAcceptClient(Index)
End Sub

Private Sub sckAcceptClient_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim packet As String
sckAcceptClient(Index).GetData packet
PP_client packet, Index
End Sub

Private Sub sckAcceptServer_Close(Index As Integer)
AuthClose Index
End Sub

Private Sub sckAcceptServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim packet As String
sckAcceptServer(Index).GetData packet
 parseServer packet, Index
End Sub

Private Sub sckLisClient_ConnectionRequest(ByVal requestID As Long)
On Error Resume Next
Dim Loaded As Boolean
    Loaded = False

    
If (sckAcceptClient.Count) <= currentsocknumber Then ' find recyclable sockets
    On Error GoTo newsock
    For i = 1 To currentsocknumber
        On Error Resume Next
        If sckAcceptClient(i).state <> sckConnected Then
        On Error GoTo socketerr
            Load sckAcceptClient(i)
            sckAcceptClient(i).Accept requestID
            logtext "Socket #" & i & " re-used"
            Loaded = True
            Exit Sub
            Exit For
        End If
    
    Next i
    
End If

If Loaded Then Exit Sub
newsock:
If currentsocknumber >= maxclients Then
    logtext "Maximum scoket reached!"
    Exit Sub
End If

currentsocknumber = currentsocknumber + 1
'On Error GoTo socketerr
    If Loaded <> True Then
    Load sckAcceptClient(currentsocknumber)
    logtext "Socket #" & currentsocknumber & " loaded"
    sckAcceptClient(currentsocknumber).Accept requestID
    End If
    
    
    
Exit Sub
'startsock:
'On Error GoTo socketerr
'    Load sckAcceptClient(i)
'    sckAcceptClient(i).Accept requestID
'    Exit Sub
    
socketerr:
    MsgBox "Eror creating scoket for " & requestID & " at this time, you may have discovered a bug in o2emu, " & vbCrLf _
     & "or settings has error", vbOKOnly, "Error starting socket"
    

End Sub

Private Sub sckServListen_ConnectionRequest(ByVal requestID As Long)
If serverssec Then
    For i = LBound(Split(authIPS, " , ")) To UBound(Split(authIPS, " , "))
    x = Split(authIPS, " , ")(i)
    If sckServListen.RemoteHostIP = x Then
        logtext sckServListen.RemoteHost & " will be authenticated as server, Security rule matched"
        GoTo loadsocket
    Else:
        logtext sckServListen.RemoteHost & " Will be dropped, security rule not match"
        frmLog.Show
        Exit Sub
    End If
    Next i
End If
logtext "Accepting " & sckServListen.RemoteHostIP & "(IP security not turned on) as server....."
loadsocket:
'On Error GoTo sckerror
Load Me.sckAcceptServer(sckAcceptServer.UBound + 1)
sckAcceptServer(sckAcceptServer.UBound).Accept requestID
'logtext sckServListen.RemoteHostIP & " is accepted and connected as hub server."
Exit Sub
sckerror:
logtext "Unable to initiate connection with IP:" & sckServListen.RemoteHostIP
frmLog.Show
End Sub
Private Sub closesocks()
On Error Resume Next
For i = sckAcceptClient.LBound To sckAcceptClient.UBound - 1
    sckAcceptClient(i).Close
Next i
End Sub

