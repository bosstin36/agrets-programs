VERSION 5.00
Begin VB.Form frmServers 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Servers List"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   4620
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdUse 
      Caption         =   "&Use"
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton cmdRem 
      Caption         =   "&Remove"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   360
      Width           =   1095
   End
   Begin VB.Frame frm1 
      BackColor       =   &H00000000&
      Caption         =   "Servers Address/IP"
      ForeColor       =   &H00FFFFFF&
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
      Begin VB.ListBox List1 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   2400
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2775
      End
   End
End
Attribute VB_Name = "frmServers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
isEdit = False
frmAddServers.Show
End Sub

Private Sub cmdEdit_Click()
If List1.ListIndex = -1 Then Exit Sub
Dim xIP As String
Dim xSelect As String
Dim xPort As Long
isEdit = True
xSelect = Split(IPlist, ";")(List1.ListIndex)
xIP = Split(xSelect, "|")(1)
xPort = Split(xIP, ":")(1)
xIP = Replace(xIP, ":" & xPort, Empty)
frmAddServers.txtName = List1.List(List1.ListIndex)
frmAddServers.txtIP = xIP
frmAddServers.txtPort = xPort
With frmAddServers
.txtName.Tag = List1.List(List1.ListIndex)
.txtPort.Tag = xPort
.txtIP.Tag = xIP
End With
Call cmdRem_Click
frmAddServers.Show
End Sub

Private Sub cmdRem_Click()
If Not List1.ListIndex = -1 Then
Dim xCurrent As String
xCurrent = Split(IPlist, ";")(List1.ListIndex)
IPlist = Replace(IPlist, xCurrent & ";", Empty)
End If
refreshlist
End Sub

Private Sub cmdUse_Click()
If List1.ListIndex = -1 Then
MsgBox "Please select an item from list to use", vbInformation + vbOKOnly, "Select item"
Exit Sub
End If
Dim xIP As String
Dim xPort As Long
Dim xName As String
Dim xSelect As String
xSelect = Split(IPlist, ";")(List1.ListIndex)
useName = Split(xSelect, "|")(0)
xIP = Split(xSelect, "|")(1)
usePort = Split(xSelect, ":")(1)
useIP = Replace(xIP, ":" & xPort, Empty)
frmMain.Label2 = "Using " & useName & " server"
frmMain.Show
Unload Me
End Sub

Private Sub Form_Load()
Dim xIP As String
Dim xName As String
Dim sCol As String

List1.Clear
refreshlist
'frmMain.Hide
End Sub

Private Sub refreshlist()
On Error Resume Next
List1.Clear
For i = LBound(Split(IPlist, ";")) To UBound(Split(IPlist, ";")) - 1
    sCol = Split(IPlist, ";")(i)
    xName = Split(sCol, "|")(0)
    xIP = Split(sCol, "|")(1)
    List1.AddItem xName
Next i
End Sub
