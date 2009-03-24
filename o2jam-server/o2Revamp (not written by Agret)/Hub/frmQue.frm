VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmQue 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hub Server Query"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   10635
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   600
      Top             =   4080
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "C&lose"
      Height          =   495
      Left            =   8280
      TabIndex        =   1
      Top             =   3960
      Width           =   2175
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   6800
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Server Statistics"
      TabPicture(0)   =   "frmQue.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblCurSockNum"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbl2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lbl3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Players"
      TabPicture(1)   =   "frmQue.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdEdit"
      Tab(1).Control(1)=   "txtID"
      Tab(1).Control(2)=   "cmdAdd"
      Tab(1).Control(3)=   "txtUsrName"
      Tab(1).Control(4)=   "chkRefresh"
      Tab(1).Control(5)=   "cmdRef"
      Tab(1).Control(6)=   "Frame1"
      Tab(1).Control(7)=   "Frame2"
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "Channels"
      TabPicture(2)   =   "frmQue.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit"
         Height          =   495
         Left            =   -70920
         TabIndex        =   16
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtID 
         Height          =   285
         Left            =   -68160
         TabIndex        =   12
         Top             =   1440
         Width           =   3255
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   495
         Left            =   -68160
         TabIndex        =   11
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox txtUsrName 
         Height          =   285
         Left            =   -68160
         TabIndex        =   10
         Top             =   1080
         Width           =   2895
      End
      Begin VB.CheckBox chkRefresh 
         Caption         =   "AutoRefresh"
         Height          =   375
         Left            =   -70920
         TabIndex        =   9
         Top             =   1920
         Width           =   1215
      End
      Begin VB.CommandButton cmdRef 
         Caption         =   "Re&fresh"
         Height          =   495
         Left            =   -70920
         TabIndex        =   8
         Top             =   720
         Width           =   855
      End
      Begin VB.Frame Frame1 
         Caption         =   "Players & ID"
         Height          =   3015
         Left            =   -74760
         TabIndex        =   6
         Top             =   600
         Width           =   3735
         Begin VB.ListBox List1 
            Height          =   2400
            Left            =   120
            TabIndex        =   7
            Top             =   360
            Width           =   3375
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Add Users Manually"
         Height          =   1815
         Left            =   -68760
         TabIndex        =   13
         Top             =   720
         Width           =   3975
         Begin VB.Label Label2 
            Caption         =   "ID"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   720
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Name"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.Label lbl3 
         Caption         =   "Number of authenticated players"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   1200
         Width           =   2775
      End
      Begin VB.Label lbl2 
         Caption         =   "Server started"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblCurSockNum 
         Height          =   255
         Left            =   2040
         TabIndex        =   3
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lbl1 
         Caption         =   "Number of sockets"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   480
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmQue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
If Frame2.Caption = "Edit User Info" Then
    players(List1.ListIndex + 1).ID = txtID
    players(List1.ListIndex + 1).usrName = txtUsrName
    Frame2.Caption = "Add Users Manually"
    cmdAdd.Caption = "Add"
    txtID = ""
    txtUsrName = ""
    Exit Sub
End If
If Len(txtID) = 0 And Len(txtUsrName) = 0 Then Exit Sub
players.addUser txtID.text, txtUsrName.text, 0
txtID = ""
txtUsrName = ""
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdEdit_Click()
If List1.ListIndex = -1 Then Exit Sub
Frame2.Caption = "Edit User Info"
cmdAdd.Caption = "Update"
Me.txtID = players(List1.ListIndex + 1).ID
Me.txtUsrName = players(List1.ListIndex + 1).usrName
End Sub

Private Sub cmdRef_Click()
Dim i As Integer
List1.Clear
If players.Count = 0 Then Exit Sub
For i = 1 To players.Count
    List1.AddItem players(i).usrName & " <[" & players(i).ID & "]>"
Next i
End Sub

Private Sub Form_Load()
On Error Resume Next
List1.Refresh
Dim i As Integer
On Error Resume Next
If players.Count = 0 Then Exit Sub
For i = 0 To players.Count
    List1.AddItem players(i).usrName & " <[" & players(i).ID & "]>"
Next i
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Timer1_Timer()
lblCurSockNum = frmMain.sckControl.Count
lbl2 = "Server Started =" & frmMain.cmdStartStop.Value
lbl3 = "Number of authenticated players : " & players.Count
List1.Refresh
If chkRefresh.Value = 1 Then
Dim i As Integer
List1.Clear
cmdEdit.enabled = False
    If players.Count = 0 Then Exit Sub
    For i = 1 To players.Count
    List1.AddItem players(i).usrName & " <[" & players(i).ID & "]>"
    Next i
ElseIf chkRefresh.Value = 0 Then
cmdEdit.enabled = True
End If
End Sub
