VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmSet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   7305
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   495
      Left            =   4200
      TabIndex        =   6
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   5640
      TabIndex        =   5
      Top             =   6600
      Width           =   1455
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   11033
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Main"
      TabPicture(0)   =   "frmSet.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbl5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtAccIP"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtSysPort"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "frmSec"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtPassKey"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtTablePath"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Server"
      TabPicture(1)   =   "frmSet.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(1)=   "txtmaxsock"
      Tab(1).Control(2)=   "txtLisPort"
      Tab(1).Control(3)=   "lbl2"
      Tab(1).Control(4)=   "lbl1"
      Tab(1).ControlCount=   5
      Begin VB.Frame Frame1 
         Caption         =   "Channels"
         Height          =   4335
         Left            =   -74760
         TabIndex        =   21
         Top             =   1680
         Width           =   3495
         Begin VB.CommandButton cmdOpenAll 
            Caption         =   "Open All"
            Height          =   375
            Left            =   2520
            TabIndex        =   24
            Top             =   840
            Width           =   855
         End
         Begin VB.CommandButton cmdCloseAll 
            Caption         =   "Close All"
            Height          =   375
            Left            =   2520
            TabIndex        =   23
            Top             =   360
            Width           =   855
         End
         Begin VB.ListBox List1 
            Height          =   3210
            Left            =   240
            Style           =   1  'Checkbox
            TabIndex        =   22
            Top             =   360
            Width           =   2175
         End
      End
      Begin VB.TextBox txtTablePath 
         Height          =   285
         Left            =   2280
         TabIndex        =   20
         Top             =   1920
         Width           =   4215
      End
      Begin VB.TextBox txtPassKey 
         Height          =   285
         Left            =   2280
         TabIndex        =   18
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Frame frmSec 
         Caption         =   "InterServer Security"
         Height          =   2055
         Left            =   240
         TabIndex        =   11
         Top             =   3720
         Visible         =   0   'False
         Width           =   6255
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   3720
            TabIndex        =   16
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   1080
            TabIndex        =   14
            Top             =   600
            Width           =   1215
         End
         Begin VB.CheckBox chkencryt 
            Caption         =   "Encrytion enabled"
            Height          =   375
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label2 
            Caption         =   "Key2"
            Height          =   255
            Left            =   3120
            TabIndex        =   15
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Key1"
            Height          =   375
            Left            =   480
            TabIndex        =   13
            Top             =   600
            Width           =   1335
         End
      End
      Begin VB.TextBox txtSysPort 
         Height          =   285
         Left            =   2280
         TabIndex        =   10
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtAccIP 
         Height          =   285
         Left            =   2280
         TabIndex        =   8
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox txtmaxsock 
         Height          =   285
         Left            =   -73440
         TabIndex        =   4
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox txtLisPort 
         Height          =   285
         Left            =   -73440
         TabIndex        =   2
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "System packetTable File"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label lbl5 
         Caption         =   "System Server Passkey"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label lbl4 
         Caption         =   "System Server Port"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label lbl3 
         Caption         =   "System Server IP :"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label lbl2 
         Caption         =   "Max. Connections:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   3
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label lbl1 
         Caption         =   "Listen Port :"
         Height          =   375
         Left            =   -74400
         TabIndex        =   1
         Top             =   720
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Dim xconfirm As Integer
xconfirm = MsgBox("Are you sure to quit without saving?", vbYesNo + vbQuestion, "Confirm?")
If xconfirm = 6 Then Unload Me
End Sub

Private Sub cmdCloseAll_Click()
For i = 0 To 19
    List1.Selected(i) = False
Next i

End Sub

Private Sub cmdOk_Click()
maxsocket = txtmaxsock
GameLisPort = txtLisPort
SysServerIP = txtAccIP
SysServerPort = txtSysPort
SysPasskey = txtPassKey
pcktTable = txtTablePath
saveChannelList
InitPcktTbl
DoSave
Unload Me
End Sub

Private Sub cmdOpenAll_Click()
For i = 0 To 19
    List1.Selected(i) = 3
Next i
End Sub

Private Sub Form_Load()
txtAccIP = SysServerIP
txtSysPort = SysServerPort
txtmaxsock = maxsocket
txtLisPort = GameLisPort
txtPassKey = SysPasskey
txtTablePath = pcktTable
loadChannelList
End Sub

Private Sub DoSave()
txtout = "MaxSocket=" & maxsocket & vbCrLf & _
            "Port=" & GameLisPort & vbCrLf & _
            "SysServerIP=" & SysServerIP & vbCrLf & _
            "SysServerPort=" & SysServerPort & vbCrLf & _
            "PassKey=" & SysPasskey & vbCrLf & _
            "SysServerPcktTbl=" & pcktTable
Open App.Path & "\CoreSettings.ini" For Output As #1
    Print #1, txtout
Close
End Sub

Private Sub loadChannelList()
For i = 1 To Channels.count
    List1.AddItem "Channel " & i
    List1.Selected(i - 1) = Channels(i).enabled
Next i
End Sub

Private Sub saveChannelList()
For i = 1 To Channels.count
    Channels(i).enabled = List1.Selected(i - 1)
Next i
End Sub
