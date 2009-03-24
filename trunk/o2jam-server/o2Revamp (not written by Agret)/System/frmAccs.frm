VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmAccs 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Accounts Setup:"
   ClientHeight    =   5250
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   7980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   7980
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid grdMain 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   8281
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "frmAccs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
grdMain.Rows = 3
Dim currLine As String
If initAccs Then
    Open accTablePath For Input As #1
    grdMain.AddItem Split(currLine, Chr(2) & Chr(3) & Chr(2))(0)
End Sub

