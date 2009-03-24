VERSION 5.00
Begin VB.Form frmSet 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings"
   ClientHeight    =   2400
   ClientLeft      =   4620
   ClientTop       =   4170
   ClientWidth     =   6390
   Icon            =   "frmSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   2  'Cross
   ScaleHeight     =   2400
   ScaleWidth      =   6390
   Begin VB.TextBox txtXclient 
      Height          =   285
      Left            =   240
      TabIndex        =   8
      Top             =   1800
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00000000&
      Caption         =   "Other"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   240
      MaskColor       =   &H00000000&
      TabIndex        =   6
      Top             =   1320
      Width           =   2295
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00000000&
      Caption         =   "Default"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   240
      MaskColor       =   &H00000000&
      TabIndex        =   5
      Top             =   1080
      Value           =   -1  'True
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   240
      Width           =   3495
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   4800
      MaskColor       =   &H00FF0000&
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000008&
      Caption         =   "Customized client file name(.exe)"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   7
      Top             =   1560
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000008&
      Caption         =   "O2jam game client select"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   1770
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   2775
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000008&
      Caption         =   "O2JamPath"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   825
   End
End
Attribute VB_Name = "frmSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkMore_Click()
If chkMore.Value = 1 Then
    frmSet.Width = frmSet.Width + 2325
ElseIf chkMore.Value = 0 Then
        frmSet.Width = frmSet.Width - 2325
End If
End Sub

Private Sub cmdOK_Click()
frmMain.Show
GameDir = Text1
If Option1(1).Value And Len(txtXclient.Text) = 0 Then rundefclient = True

If Option1(1).Value = True Then
    rundefclient = False
    customclient = txtXclient.Text
ElseIf Option1(0).Value = True Then
    rundefclient = True
End If
Unload Me
End Sub

Private Sub Form_Load()
frmSet.Text1 = GameDir
If rundefclient = True Then
Label1(2).Visible = False
Option1(0).Value = True
ElseIf rundefclient = False Then
Option1(1).Value = True
Option1(0).Value = False
Label1(2).Visible = True
txtXclient = customclient
End If
End Sub

Private Sub Label1_Click(index As Integer)
MsgBox GetString(HKEY_CURRENT_USER, "Software\e-games\o2jam", "location")
End Sub

Private Sub Option1_Click(index As Integer)
If Option1(1).Value = True Then
    Label1(2).Visible = True
    txtXclient.Visible = True
ElseIf Option1(0).Value = True Then
    Label1(2).Visible = False
    txtXclient.Visible = False
End If
End Sub
