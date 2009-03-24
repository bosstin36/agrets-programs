VERSION 5.00
Begin VB.Form frmAddServers 
   BackColor       =   &H00000000&
   Caption         =   "Add a favourite server......"
   ClientHeight    =   2175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4095
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H80000009&
   LinkTopic       =   "Form1"
   ScaleHeight     =   2175
   ScaleWidth      =   4095
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPort 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2880
      TabIndex        =   8
      Top             =   1200
      Width           =   855
   End
   Begin VB.CheckBox cmdTar 
      BackColor       =   &H00000000&
      Caption         =   "Target Server Uses name authentication"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   3495
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "C&ancel"
      Height          =   495
      Left            =   1320
      TabIndex        =   5
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdSav 
      Caption         =   "&OK"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox txtIP 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Top             =   720
      Width           =   2175
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Port"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   2400
      TabIndex        =   7
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Server Host/IP"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Server Name"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmAddServers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
If isEdit Then
    txtName.Text = txtName.Tag
    txtPort.Text = txtPort.Tag
    txtIP.Text = txtIP.Tag
    Call cmdSav_Click
End If
Unload Me
End Sub

Private Sub cmdSav_Click()
IPlist = IPlist & Me.txtName & "|" & Me.txtIP & ":" & txtPort & ";"
refreshlist
Unload Me
End Sub

Private Sub Form_Load()
'Unload frmServers
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmServers.Show
End Sub
Private Sub refreshlist()
With frmServers
.List1.Clear
For i = LBound(Split(IPlist, "~")) To UBound(Split(IPlist, ";")) - 1
    sCol = Split(IPlist, ";")(i)
    xName = Split(sCol, "|")(0)
    xIP = Split(sCol, "|")(1)
    .List1.AddItem xName
Next i
End With
End Sub

