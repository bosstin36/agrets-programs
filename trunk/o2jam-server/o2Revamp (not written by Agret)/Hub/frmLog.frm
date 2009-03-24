VERSION 5.00
Begin VB.Form frmLog 
   Caption         =   "Log"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8130
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   8130
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtLog 
      BackColor       =   &H8000000F&
      Height          =   2415
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   3
      Top             =   0
      Width           =   7935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Back"
      Height          =   615
      Left            =   7200
      TabIndex        =   2
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton cmdSav 
      Caption         =   "Save"
      Height          =   615
      Left            =   6240
      TabIndex        =   1
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   855
   End
End
Attribute VB_Name = "frmLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClear_Click()
txtLog.text = Empty
End Sub

Private Sub Command1_Click()
Me.Hide
End Sub
