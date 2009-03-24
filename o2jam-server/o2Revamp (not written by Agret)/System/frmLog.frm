VERSION 5.00
Begin VB.Form frmLog 
   Caption         =   "Server log"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8265
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   8265
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdback 
      Caption         =   "Back"
      Height          =   495
      Left            =   7560
      TabIndex        =   3
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
      Height          =   615
      Left            =   7560
      TabIndex        =   2
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   615
      Left            =   7560
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox txtLog 
      Height          =   2895
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   7335
   End
End
Attribute VB_Name = "frmLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdback_Click()
Me.Hide
End Sub

Private Sub cmdClear_Click()
txtLog.text = Empty
End Sub

Private Sub cmdsave_Click()
Dim logpath As String
logpath = App.Path & "\log.txt"
Open logpath For Append As #1
Print #1, txtLog
Close 1
MsgBox "Log File appended to " & logpath, vbOKOnly + vbInformation, "Sucess!"
End Sub
