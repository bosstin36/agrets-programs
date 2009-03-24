VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   3840
   ClientLeft      =   165
   ClientTop       =   780
   ClientWidth     =   8865
   LinkTopic       =   "Form1"
   ScaleHeight     =   3840
   ScaleWidth      =   8865
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRIndex 
      Height          =   495
      Left            =   2760
      TabIndex        =   6
      Top             =   1920
      Width           =   495
   End
   Begin VB.TextBox txtIndex 
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox txtCha 
      Height          =   495
      Left            =   2040
      TabIndex        =   4
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Read"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox txtRName 
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   1200
      Width           =   5775
   End
   Begin VB.TextBox txtchannel 
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Write"
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   2160
      TabIndex        =   7
      Top             =   2640
      Width           =   4575
   End
   Begin VB.Menu cmdGlFile 
      Caption         =   "File"
      Begin VB.Menu cmdRes 
         Caption         =   "Restore"
      End
      Begin VB.Menu cmdState 
         Caption         =   "Status"
      End
      Begin VB.Menu cmdGlExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdLoad_Click()
Dim a As Integer
a = txtchannel.text
Channels(a).Rooms.Add txtRName.text
End Sub

Private Sub cmdRes_Click()
Dim trayi As NOTIFYICONDATA
    trayi.cbSize = Len(trayi)
    trayi.hWnd = frmMain.pichook.hWnd
    trayi.uId = 1&
    'Delete the icon
    Shell_NotifyIcon NIM_DELETE, trayi
    frmMain.Show
End Sub

Private Sub Command1_Click()
Dim a As Integer, b As Integer
a = txtCha.text
b = txtRIndex.text
Label1.Caption = Channels(a).Rooms.Item(b).rName
End Sub

