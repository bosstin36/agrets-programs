VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmLoad 
   BorderStyle     =   0  'None
   Caption         =   "Loading"
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   ScaleHeight     =   1800
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStop 
      Caption         =   "C&ancel"
      Height          =   615
      Left            =   4800
      TabIndex        =   1
      Top             =   840
      Width           =   1335
   End
   Begin MSComctlLib.ProgressBar bar1 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lblProgress 
      Height          =   615
      Left            =   480
      TabIndex        =   3
      Top             =   1080
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Loading..."
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   1695
   End
End
Attribute VB_Name = "frmLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdStop_Click()
frmMain.tmr_connect.Interval = 0
frmMain.tmr_connect.Enabled = False
frmMain.winsckConnect.Close
conState = 0
loadState False
End Sub
