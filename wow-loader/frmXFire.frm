VERSION 5.00
Begin VB.Form frmXFire 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WoW Loader - XFire Settings"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   4590
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraXFireSettings 
      Caption         =   "XFire Settings"
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.CommandButton cmdSaveClose 
         Caption         =   "Save Settings && Close"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   4440
         Width           =   4095
      End
      Begin VB.CommandButton cmdDelFake 
         Caption         =   "Delete fake folder entered above"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   3960
         Width           =   4095
      End
      Begin VB.TextBox txtFakeFolder 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Text            =   "C:\Games\WoW_XFire"
         Top             =   2160
         Width           =   4095
      End
      Begin VB.CheckBox chkEnableXFire 
         Caption         =   "Check to enable XFire supported mode"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4095
      End
      Begin VB.Label lblWarning 
         Caption         =   $"frmXFire.frx":0000
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         TabIndex        =   5
         Top             =   2520
         Width           =   4095
      End
      Begin VB.Label lblFakeFolder 
         BackStyle       =   0  'Transparent
         Caption         =   "Fake Folder:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1920
         Width           =   3975
      End
      Begin VB.Label lblXFireInfo 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmXFire.frx":0103
         Height          =   1455
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   4095
      End
   End
End
Attribute VB_Name = "frmXFire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'========================================================
'   API declaration
'========================================================
Private Declare Function ShellExecute _
    Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long

Private Sub cmdDelFake_Click()
    ShellExecute hwnd, "open", App.Path & "\junction.exe", " -d " & """" & txtFakeFolder.Text & """", vbNullString, 0
    MsgBox "Fake Folder Deleted.", vbInformation, "Done"
End Sub

Private Sub cmdSaveClose_Click()
    SaveSetting "WoW Loader", "Pref", "XFireFakeFolder", txtFakeFolder.Text
    If chkEnableXFire.Value = 1 Then
        SaveSetting "WoW Loader", "Pref", "UseXFire", "True"
    Else
        DeleteSetting "WoW Loader", "Pref", "UseXFire"
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    If Len(GetSetting("WoW Loader", "Pref", "XFireFakeFolder")) > 0 Then
        txtFakeFolder.Text = GetSetting("WoW Loader", "Pref", "XFireFakeFolder")
    End If
    
    If GetSetting("WoW Loader", "Pref", "UseXFire") = "True" Then
        chkEnableXFire.Value = 1
    Else
        chkEnableXFire.Value = 0
    End If
End Sub
