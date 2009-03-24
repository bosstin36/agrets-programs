VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmSet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Server Settings"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   8640
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   6960
      TabIndex        =   7
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   5280
      TabIndex        =   6
      Top             =   4560
      Width           =   1455
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4335
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   7646
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Main"
      TabPicture(0)   =   "frmSet.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtAccPcktTbl"
      Tab(0).Control(1)=   "txtHubPcktTbl"
      Tab(0).Control(2)=   "txtmaxclient"
      Tab(0).Control(3)=   "txtHubPort"
      Tab(0).Control(4)=   "txtAccPort"
      Tab(0).Control(5)=   "Label7"
      Tab(0).Control(6)=   "Label6"
      Tab(0).Control(7)=   "Label3"
      Tab(0).Control(8)=   "Label2"
      Tab(0).Control(9)=   "Label1"
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Accounts"
      TabPicture(1)   =   "frmSet.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "chkUseRam"
      Tab(1).Control(1)=   "txtDBpath"
      Tab(1).Control(2)=   "Label4"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Hub Service"
      TabPicture(2)   =   "frmSet.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "List1"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame2"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "chkRepState"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      Begin VB.CheckBox chkRepState 
         Caption         =   "Log hub server status (annoying)"
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   3840
         Width           =   2895
      End
      Begin VB.TextBox txtAccPcktTbl 
         Height          =   285
         Left            =   -72600
         TabIndex        =   24
         Top             =   2160
         Width           =   3135
      End
      Begin VB.TextBox txtHubPcktTbl 
         Height          =   285
         Left            =   -72600
         TabIndex        =   21
         Top             =   1800
         Width           =   3135
      End
      Begin VB.Frame Frame2 
         Caption         =   "HubServer Connect Security"
         Height          =   975
         Left            =   240
         TabIndex        =   12
         Top             =   600
         Width           =   3255
         Begin VB.TextBox txtPassKey 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   840
            PasswordChar    =   "*"
            TabIndex        =   13
            Top             =   330
            Width           =   2055
         End
         Begin VB.Label Label5 
            Caption         =   "PassKey:"
            Height          =   375
            Left            =   120
            TabIndex        =   14
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.CheckBox chkUseRam 
         Caption         =   "Load Accounts to memory."
         Height          =   255
         Left            =   -74760
         TabIndex        =   11
         ToolTipText     =   "Uses more RAM, but reduces time and number of harddisk reads, use this option if less accounts is needed"
         Top             =   3840
         Width           =   3975
      End
      Begin VB.TextBox txtDBpath 
         Height          =   285
         Left            =   -73440
         TabIndex        =   9
         Top             =   600
         Width           =   3495
      End
      Begin VB.TextBox txtmaxclient 
         Height          =   285
         Left            =   -72840
         TabIndex        =   3
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox txtHubPort 
         Height          =   285
         Left            =   -72840
         TabIndex        =   2
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txtAccPort 
         Height          =   285
         Left            =   -72840
         TabIndex        =   0
         Top             =   720
         Width           =   1575
      End
      Begin VB.ListBox List1 
         Height          =   2010
         ItemData        =   "frmSet.frx":0054
         Left            =   3720
         List            =   "frmSet.frx":0056
         TabIndex        =   15
         Top             =   1440
         Width           =   3135
      End
      Begin VB.Frame Frame1 
         Caption         =   "Hub servers management"
         Height          =   3375
         Left            =   3600
         TabIndex        =   16
         Top             =   600
         Width           =   3735
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Add"
            Height          =   375
            Left            =   1560
            TabIndex        =   20
            Top             =   2880
            Width           =   855
         End
         Begin VB.CommandButton cmdRem 
            Caption         =   "Remove"
            Height          =   375
            Left            =   2520
            TabIndex        =   19
            Top             =   2880
            Width           =   855
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Allow Any IPs"
            Height          =   255
            Left            =   360
            TabIndex        =   18
            Top             =   240
            Width           =   2295
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Only allow these IPs"
            Height          =   375
            Left            =   360
            TabIndex        =   17
            Top             =   480
            Width           =   2175
         End
      End
      Begin VB.Label Label7 
         Caption         =   "Account server Packet Table"
         Height          =   255
         Left            =   -74880
         TabIndex        =   23
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Label Label6 
         Caption         =   "Hub server Packet Table"
         Height          =   255
         Left            =   -74760
         TabIndex        =   22
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label Label4 
         Caption         =   "Accounts DB path"
         Height          =   255
         Left            =   -74880
         TabIndex        =   10
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Maximum clients "
         Height          =   255
         Left            =   -74520
         TabIndex        =   8
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Hub service Port"
         Height          =   255
         Left            =   -74520
         TabIndex        =   5
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "LauncherClient service port"
         Height          =   255
         Left            =   -74880
         TabIndex        =   4
         Top             =   720
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private authIPsX As String
Private userQuit As Boolean

Private Sub cmdAdd_Click()
x = InputBox("Type a valid IP address, only IP is accepted", "Input IP")
    If UBound(Split(x, ".")) = 3 Then
    authIPsX = authIPsX & x & " , "
    ElseIf Len(x) > 0 Then
    MsgBox ("Input error!")
    End If
    listips
End Sub

Private Sub cmdOK_Click()
userQuit = True
stateLog = chkRepState
loadacc2ram = Me.chkUseRam
accTablePath = Me.txtDBpath
clientport = Me.txtAccPort
hubport = Me.txtHubPort
authIPS = authIPsX
maxclients = Me.txtmaxclient
serverssec = Option2.Value
passKey = txtPassKey.text
hub_PcktTbl = txtHubPcktTbl
readRespTable

Close All
If valid Then
    sFile = App.Path & "\SysSettings.ini"
    ssave = "authIPs=" & authIPS & vbCrLf & _
        "IPsecurity=" & Option2.Value & vbCrLf & _
        "ClientPort=" & clientport & vbCrLf & _
        "DBpath=" & txtDBpath & vbCrLf & _
        "ServerPort=" & hubport & vbCrLf & _
        "MaxClients=" & maxclients & vbCrLf & _
        "sPacketTable=" & hub_PcktTbl & vbCrLf & _
        "HubPassCode=" & passKey
Open sFile For Output As 1
Print #1, ssave
Close
frmMain.Show
Unload Me
Else:
    MsgBox "Please recheck settings, settings has error", vbInformation, "Settings error"
End If
End Sub

Private Sub cmdQuit_Click()
Dim a As Integer
a = MsgBox("Quit without saving?", vbYesNo + vbQuestion, "Confirm?")
If a = 6 Then
    frmMain.Show
    Unload Me
ElseIf a = 7 Then
End If
End Sub

Private Sub cmdRem_Click()
On Error GoTo errmsg
authIPsX = Replace(authIPsX, List1.List(List1.ListIndex) & " , ", Empty)
List1.RemoveItem (List1.ListIndex)
Exit Sub
errmsg:
MsgBox "Please select an item to remove", vbOKOnly + vbExclamation, "Error"
End Sub

Private Sub Form_Load()
chkRepState = stateLog
Me.chkUseRam.Value = loadacc2ram
authIPsX = authIPS
Me.txtAccPort = clientport
txtDBpath = accTablePath
txtmaxclient = maxclients
Me.txtHubPort = hubport
txtPassKey = passKey
txtHubPcktTbl = hub_PcktTbl
If serverssec Then
    Option1.Value = False
    Option2.Value = True
    Call Option2_Click
Else:
    Option1.Value = True
    Option2.Value = False
    Call Option1_Click
End If
    listips
End Sub

Private Sub Form_Resize()
SSTab1.Height = Me.Height - 1410

End Sub

Private Sub Form_Unload(Cancel As Integer)
If userQuit Then Exit Sub
Dim a As Integer
a = MsgBox("Quit without saving?", vbYesNo + vbQuestion, "Confirm?")
If a = 7 Then
    Cancel = 1
ElseIf a = 6 Then
End If

End Sub

Private Sub Option1_Click()
List1.Enabled = False
cmdAdd.Enabled = False
cmdRem.Enabled = False
End Sub

Private Sub Option2_Click()
List1.Enabled = True
cmdAdd.Enabled = True
cmdRem.Enabled = True
End Sub

Private Sub listips()
        List1.Clear
For i = LBound(Split(authIPsX, " , ")) To UBound(Split(authIPsX, " , ")) - 1
        curip = Split(authIPsX, " , ")(i)
        If UBound(Split(curip, ".")) = 3 Then
        List1.AddItem curip
        End If
Next i
End Sub

