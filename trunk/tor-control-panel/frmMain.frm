VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tor Control Panel"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3765
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   3765
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkStartHidden 
      Caption         =   "Start hidden"
      Height          =   255
      Left            =   2400
      TabIndex        =   6
      Top             =   720
      Width           =   1215
   End
   Begin VB.OptionButton chkNormal 
      Caption         =   "Normal"
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   720
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.OptionButton chkMinimized 
      Caption         =   "Minimized"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start Tor"
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "Show Tor"
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdHide 
      Caption         =   "Hide Tor"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblStatus 
      Caption         =   "Ready."
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   3495
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   3720
      Y1              =   1080
      Y2              =   1080
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Const SW_HIDE = 0
Private Const SW_SHOWNORMAL = 1
Private Const SW_NORMAL = 1
Private Const SW_SHOWMINIMIZED = 2
Private Const SW_SHOWMAXIMIZED = 3
Private Const SW_MAXIMIZE = 3
Private Const SW_SHOWNOACTIVATE = 4
Private Const SW_SHOW = 5
Private Const SW_MINIMIZE = 6
Private Const SW_SHOWMINNOACTIVE = 7
Private Const SW_SHOWNA = 8
Private Const SW_RESTORE = 9
Private Const SW_SHOWDEFAULT = 10
Private Const SW_MAX = 10

Private Sub cmdHide_Click()
    WindowHandle = 0
    WindowHandle = FindWindow(vbNullString, App.Path & "\tor.exe")
    If WindowHandle <> 0 Then
        ShowWindow WindowHandle, SW_HIDE
        lblStatus.Caption = "Tor hidden."
    Else
        lblStatus.Caption = "Tor not running."
    End If
End Sub

Private Sub cmdShow_Click()
    WindowHandle = 0
    WindowHandle = FindWindow(vbNullString, App.Path & "\tor.exe")
    If WindowHandle <> 0 Then
        If chkMinimized.Value = True Then
            ShowWindow WindowHandle, SW_SHOWMINNOACTIVE
        ElseIf chkNormal.Value = True Then
            ShowWindow WindowHandle, SW_SHOWNORMAL
        End If
        lblStatus.Caption = "Tor shown."
    Else
        lblStatus.Caption = "Tor not running."
    End If
End Sub

Private Sub cmdStart_Click()
    WindowHandle = 0
    WindowHandle = FindWindow(vbNullString, App.Path & "\tor.exe")
    If WindowHandle <> 0 Then
        ShowWindow WindowHandle, SW_SHOWNA
    Else
        On Error GoTo fuckoff
        If chkStartHidden.Value = 1 Then
            Shell App.Path & "\tor.exe", vbHide
        ElseIf chkStartHidden.Value = False Then
            Shell App.Path & "\tor.exe", vbNormalFocus
        End If
        lblStatus.Caption = "Tor started."
    End If
    Exit Sub
fuckoff:
    MsgBox "Error Launching Tor." & vbNewLine & vbNewLine & Err.Description & vbNewLine & vbNewLine & "Is Tor in the right directory?", vbCritical, Err.Description
    lblStatus.Caption = "Error launching Tor."
    Exit Sub
End Sub

Private Sub Form_Load()
    WindowHandle = 0
    WindowHandle = FindWindow(vbNullString, App.Path & "\tor.exe")
    If WindowHandle <> 0 Then
        lblStatus.Caption = "Tor running."
    Else
        lblStatus.Caption = "Tor not running."
    End If
End Sub
