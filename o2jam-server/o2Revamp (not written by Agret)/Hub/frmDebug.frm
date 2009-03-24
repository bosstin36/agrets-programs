VERSION 5.00
Begin VB.Form frmDebug 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Debug Console"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   10500
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   9120
      TabIndex        =   2
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox txtIn 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3960
      Width           =   8895
   End
   Begin VB.TextBox txtDebug 
      Height          =   3735
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   10215
   End
End
Attribute VB_Name = "frmDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdApply_Click()
processInput txtIn
txtIn.text = ""

End Sub

Private Sub txtIn_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
processInput txtIn
txtIn.text = ""
End If
End Sub

Private Sub processInput(code As String)
Dim cmd As String
Dim rslt As String
On Error Resume Next
    cmd = Split(code)(0)
    'rslt = Split(code)(1)
    Select Case cmd
    Case Is = "StartServer"
        frmMain.sckListen.LocalPort = GameLisPort
        frmMain.sckListen.Listen
        frmMain.cmdStartStop.Value = 1
        Dim i As Integer
        For i = 1 To 20
            Channels.Add i
            Channels(i).enabled = True
        Next i
    Case Is = "StopServer"
        StopServer
    Case Is = "test"
        frmTest.Show
    Case Else:
        debugText "Unknown input <" & cmd & ">"
    End Select
End Sub

