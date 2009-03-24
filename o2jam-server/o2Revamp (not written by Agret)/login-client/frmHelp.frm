VERSION 5.00
Begin VB.Form frmHelp 
   BackColor       =   &H00000000&
   Caption         =   "Help Mode"
   ClientHeight    =   3345
   ClientLeft      =   5610
   ClientTop       =   2580
   ClientWidth     =   4680
   ControlBox      =   0   'False
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form1"
   MousePointer    =   2  'Cross
   ScaleHeight     =   3345
   ScaleWidth      =   4680
   Begin VB.CommandButton helpOK 
      Caption         =   "Understand"
      Height          =   255
      Left            =   1320
      TabIndex        =   6
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Help"
      ForeColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   4215
      Begin VB.Label Label5 
         BackColor       =   &H00000000&
         Caption         =   "4) Question ask forum ( http://o2emu.chathome.net )"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         MousePointer    =   2  'Cross
         TabIndex        =   5
         Top             =   1440
         Width           =   3855
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         Caption         =   "3) Insert your Password ( ex : 12345 )"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   3015
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   "2) Insert your ID ( ex : o2user )"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   3015
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "1) Client on ""Servers"" button to choose a server"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   3855
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "O2emu Client Help"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   1125
      Left            =   2400
      Picture         =   "frmHelp.frx":0CCA
      Top             =   120
      Width           =   1500
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub helpOK_Click()
frmMain.Show
Unload Me
End Sub

