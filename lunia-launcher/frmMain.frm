VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lunia Launcher - By Agret (alias.zero2097@gmail.com)"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7455
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   7455
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3480
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   5040
      Width           =   2175
   End
   Begin VB.TextBox txtUser 
      Height          =   285
      Left            =   600
      TabIndex        =   3
      Top             =   5040
      Width           =   2175
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   4440
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Login"
      Height          =   375
      Left            =   5880
      TabIndex        =   1
      Top             =   5000
      Width           =   1335
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      ExtentX         =   12726
      ExtentY         =   8281
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Label lblPass 
      BackStyle       =   0  'Transparent
      Caption         =   "Pass:"
      Height          =   255
      Left            =   3000
      TabIndex        =   4
      Top             =   5055
      Width           =   975
   End
   Begin VB.Label lblUser 
      BackStyle       =   0  'Transparent
      Caption         =   "User:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   5055
      Width           =   975
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim StrHTML As String

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
    
Dim LuniaUser, LuniaPass As String
    
Private Sub cmdLaunch_Click()
    WebBrowser1.Navigate "http://lunia.ijji.com/common/launch.nhn?gameId=u_lunia&subId="
End Sub

Private Sub cmdLogin_Click()
    SaveSetting "LuniaLauncher", "LoginInfo", "User", txtUser.Text
    LuniaUser = txtUser.Text
    SaveSetting "LuniaLauncher", "LoginInfo", "Pass", txtPass.Text
    LuniaPass = txtPass.Text
    DoLogin
End Sub

Private Function DoLogin()
    WebBrowser1.Navigate "http://login.ijji.com/login.nhn?m=login&nextURL=http://lunia.ijji.com/common/launch.nhn?gameId=u_lunia&memberid=" & LuniaUser & "&password=" & LuniaPass & "&secure=false"
End Function

Private Sub txtPass_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdLogin_Click
End Sub

Private Sub txtUser_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdLogin_Click
End Sub

Private Sub Form_Load()
    WebBrowser1.Navigate "about:blank"
    
    If App.PrevInstance = True Then End
    
    txtUser.Text = GetSetting("LuniaLauncher", "LoginInfo", "User")
    txtPass.Text = GetSetting("LuniaLauncher", "LoginInfo", "Pass")
    
    If Command <> "" Then
        If LuniaUser = "" Then LuniaUser = GetBetween(Command, "User=", ";")
        If LuniaUser = "" Then LuniaUser = GetBetween(Command, "Username=", ";")
        If LuniaPass = "" Then LuniaPass = GetBetween(Command, "Pass=", ";")
        If LuniaPass = "" Then LuniaPass = GetBetween(Command, "Password=", ";")
        If LuniaUser <> "" And LuniaPass <> "" Or Command = "Login" Or Command = "DoLogin" Then
            If Command <> "Login" And Command <> "DoLogin" Then
                Me.Hide
            Else
                LuniaUser = GetSetting("LuniaLauncher", "LoginInfo", "User")
                LuniaPass = GetSetting("LuniaLauncher", "LoginInfo", "Pass")
            End If
            DoLogin
        End If
    End If
    
End Sub

Private Sub WebBrowser1_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    If URL = "http://drift.ijji.com/common/launch.nhn?gameId=u_lunia" Then
        ShellExecute hwnd, "open", GetBetween(ShowHtmlTextFromUrl(WebBrowser1), "launchScript" & """" & ":" & """", """" & ","), vbNullString, vbNullString, 1
        End
    End If
End Sub

Private Function ShowHtmlTextFromUrl(brw As WebBrowser) As String
    Dim URL As String
    URL = brw.LocationURL
    ShowHtmlTextFromUrl = Inet1.OpenURL(URL)
End Function

Function GetBetween(ByRef TextToParse As String, StartDelimiter As String, EndDelimiter As String) As String
    Dim lngStart As Long, lngend As Long
  
    lngStart = InStr(1, TextToParse, StartDelimiter)
    If lngStart Then
        lngStart = lngStart + Len(StartDelimiter)
        lngend = InStr(lngStart, TextToParse, EndDelimiter)
      
        If lngend Then
            GetBetween = Mid$(TextToParse, lngStart, lngend - lngStart)
        End If
    End If
End Function

