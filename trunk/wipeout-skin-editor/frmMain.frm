VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmMain 
   Caption         =   "Wipeout Pulse Hacked Skin Editor"
   ClientHeight    =   9825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10440
   LinkTopic       =   "Form1"
   ScaleHeight     =   9825
   ScaleWidth      =   10440
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraAdvanced 
      Caption         =   "Advanced Features"
      Height          =   1455
      Left            =   120
      TabIndex        =   11
      Top             =   8040
      Width           =   10095
      Begin VB.CommandButton cmdBackup 
         Caption         =   "Backup All Skins (Slots 1-5)"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame fraSkins 
      Caption         =   "Skin Editor"
      Height          =   7815
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   10095
      Begin VB.ComboBox cmbSkinSlot 
         Height          =   315
         ItemData        =   "frmMain.frx":0000
         Left            =   960
         List            =   "frmMain.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   500
         Width           =   1095
      End
      Begin VB.TextBox txtSkinSource 
         Height          =   285
         Left            =   4560
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   200
         Width           =   5175
      End
      Begin VB.ComboBox cmbSelectShip 
         Height          =   315
         ItemData        =   "frmMain.frx":0004
         Left            =   1200
         List            =   "frmMain.frx":002C
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   200
         Width           =   2175
      End
      Begin SHDocVwCtl.WebBrowser WebSkin0 
         Height          =   6855
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   9735
         ExtentX         =   17171
         ExtentY         =   12091
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   0
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
      Begin VB.Label Label1 
         Caption         =   "Skin Slot:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   550
         Width           =   6975
      End
      Begin VB.Label lblSkinSource 
         Caption         =   "Skin Source:"
         Height          =   255
         Left            =   3480
         TabIndex        =   7
         Top             =   240
         Width           =   6495
      End
      Begin VB.Label lblPickShip 
         BackStyle       =   0  'Transparent
         Caption         =   "Select a ship:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1695
      End
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   1440
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Frame fraLogin 
      Caption         =   "Login"
      Height          =   7815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10095
      Begin VB.CommandButton cmdImLoggedIn 
         Caption         =   "I'm Logged into my account now"
         Height          =   375
         Left            =   1920
         TabIndex        =   2
         Top             =   7320
         Width           =   6375
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   6975
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   9855
         ExtentX         =   17383
         ExtentY         =   12303
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
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function URLDownloadToFile Lib "urlmon" _
    Alias "URLDownloadToFileA" (ByVal pCaller As Long, _
    ByVal szURL As String, ByVal szFileName As String, _
    ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Private Sub cmdBackup_Click()
    cmdBackup.Enabled = False
    cmdBackup.Caption = "Please Wait (0/4)..."
    URLDownloadToFile 0, "http://www.wipeout-game.com/html/ship/Thumb0.png", App.Path & "\img\Thumb0.png", 0, 0
    URLDownloadToFile 0, "http://www.wipeout-game.com/html/ship/Tex0.png", App.Path & "\img\Tex0.png", 0, 0
    cmdBackup.Caption = "Please Wait (1/4)..."
    URLDownloadToFile 0, "http://www.wipeout-game.com/html/ship/Thumb1.png", App.Path & "\img\Thumb1.png", 0, 0
    URLDownloadToFile 0, "http://www.wipeout-game.com/html/ship/Tex1.png", App.Path & "\img\Tex1.png", 0, 0
    cmdBackup.Caption = "Please Wait (2/4)..."
    URLDownloadToFile 0, "http://www.wipeout-game.com/html/ship/Thumb2.png", App.Path & "\img\Thumb2.png", 0, 0
    URLDownloadToFile 0, "http://www.wipeout-game.com/html/ship/Tex2.png", App.Path & "\img\Tex2.png", 0, 0
    cmdBackup.Caption = "Please Wait (3/4)..."
    URLDownloadToFile 0, "http://www.wipeout-game.com/html/ship/Thumb3.png", App.Path & "\img\Thumb3.png", 0, 0
    URLDownloadToFile 0, "http://www.wipeout-game.com/html/ship/Tex3.png", App.Path & "\img\Tex3.png", 0, 0
    cmdBackup.Caption = "Please Wait (4/4)..."
    URLDownloadToFile 0, "http://www.wipeout-game.com/html/ship/Thumb4.png", App.Path & "\img\Thumb4.png", 0, 0
    URLDownloadToFile 0, "http://www.wipeout-game.com/html/ship/Tex4.png", App.Path & "\img\Tex4.png", 0, 0
    MsgBox "Backup Completed.", vbInformation, "Complete"
    cmdBackup.Enabled = True
    cmdBackup.Caption = "Backup All Skins"
    Open App.Path & "\img\Thumb0.htm" For Output As #1
        Print #1, "<body style=" & """" & "margin-left: 0px; margin-top: 0px;" & """" & ">"
        Print #1, "<img src=" & """" & "Thumb0.png" & """" & ">"
        Print #1, "</body>"
    Close #1
End Sub

Private Sub cmdImLoggedIn_Click()
    fraLogin.Visible = False
'    Unload WebBrowser1
    fraSkins.Visible = True
'    MsgBox "Mission Success?"
End Sub

Private Sub Form_Load()
    WebBrowser1.Navigate "https://store.playstation.com/external/index.vm?returnURL=https://www.wipeout-game.com/html/main/TicketLoginSubmit"
End Sub

Private Sub ThisDoesntExist()
    ' Download the file.
    URLDownloadToFile 0, _
        "http://www.vb-helper.com/vbhelper_425_64.gif", _
        "C:\vbhelper_425_64.gif", 0, 0

    ' Display the image.
    picBanner.Picture = _
        LoadPicture("C:\vbhelper_425_64.gif")
End Sub
