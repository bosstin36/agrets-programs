VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Notepad  - Untitled"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4335
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   4335
   StartUpPosition =   3  'Windows Default
   Begin Notepad.VBHotKey VBHotKey3 
      Left            =   2400
      Top             =   1200
      _ExtentX        =   1296
      _ExtentY        =   1296
      VKey            =   75
   End
   Begin Notepad.VBHotKey VBHotKey2 
      Left            =   1560
      Top             =   1200
      _ExtentX        =   1296
      _ExtentY        =   1296
      VKey            =   80
   End
   Begin Notepad.VBHotKey VBHotKey1 
      Left            =   720
      Top             =   1200
      _ExtentX        =   1296
      _ExtentY        =   1296
      VKey            =   79
   End
   Begin VB.CommandButton cmdHide 
      Caption         =   "Hide"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   1215
   End
   Begin VB.ListBox lstGames 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2370
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4335
   End
   Begin VB.Label lblCredits 
      Alignment       =   1  'Right Justify
      Caption         =   "* Created by Agret, 2008 Edition *"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3360
      Width           =   4095
   End
   Begin VB.Label lblAbout 
      Caption         =   "Press Windows key && a button:"
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   2880
      Width           =   2655
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
Dim AppNames(999) As String

Private Sub HideGames()
    Dim i As Integer
    Dim Minimize As Boolean
    For i = 0 To lstGames.ListCount - 1
        'lstGames.ListIndex = i
        WindowHandle = 0
        'WindowHandle = FindWindow(vbNullString, lstGames.Text)
        If Left(AppNames(i), 1) = "^" Then
            Minimize = True
            WindowHandle = FindWindow(vbNullString, Right(AppNames(i), Len(AppNames(i)) - 1))
        Else
            Minimize = False
            WindowHandle = FindWindow(vbNullString, AppNames(i))
        End If
        If WindowHandle <> 0 Then
            If Minimize = True Then ShowWindow WindowHandle, SW_MINIMIZE
            ShowWindow WindowHandle, SW_HIDE
            frmAltTab.Show
            frmAltTab.SetFocus
            frmAltTab.Hide
            Unload frmAltTab
        End If
    Next i
End Sub

Private Sub ShowGames()
    Dim i As Integer
    For i = 0 To lstGames.ListCount - 1
        'lstGames.ListIndex = i
        WindowHandle = 0
        'WindowHandle = FindWindow(vbNullString, lstGames.Text)
        If Left(AppNames(i), 1) = "^" Then
            WindowHandle = FindWindow(vbNullString, Right(AppNames(i), Len(AppNames(i)) - 1))
        Else
            WindowHandle = FindWindow(vbNullString, AppNames(i))
        End If
        If WindowHandle <> 0 Then
            ShowWindow WindowHandle, SW_SHOW
        End If
    Next i
End Sub

Private Sub cmdHide_Click()
    Me.Hide
End Sub

Private Function ReloadAppNames()
    For i = 0 To lstGames.ListCount - 1
        lstGames.ListIndex = i
        AppNames(i) = lstGames.Text
    Next i
End Function

Private Sub cmdAdd_Click()
    Dim TempTitle As String
    TempTitle = InputBox("Game Title?", "Enter Game Title")
    If Len(TempTitle) > 2 Then
        lstGames.AddItem TempTitle
    End If
    ReloadAppNames
End Sub

Private Sub cmdEdit_Click()
    Dim TempGame As String
    TempGame = lstGames.Text
    lstGames.RemoveItem lstGames.ListIndex
    lstGames.AddItem InputBox("Edit Game", "Edit Game", TempGame)
    ReloadAppNames
End Sub

Private Sub cmdRemove_Click()
    lstGames.RemoveItem lstGames.ListIndex
    ReloadAppNames
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveList App.Path & "\games.txt", lstGames
End Sub

Private Sub Form_Load()
    If App.PrevInstance = True Then
        MsgBox "Already Running." & vbNewLine & "Press Windows Key + K to show.", , "Already Running."
    End If
    lblAbout.Caption = "Press Windows key && a button:" & vbNewLine & "O - Hide | P - Show | K - Show Me"
    LoadList App.Path & "\games.txt", lstGames
End Sub

Private Sub LoadList(sLocation As String, lstListBox As ListBox)
    On Error GoTo dlgerror
    Dim sCurrent As String
    Dim i As Integer
    lstListBox.Clear
    Open sLocation For Input As #1
    i = 0
    Do Until EOF(1)
    Line Input #1, sCurrent
    AppNames(i) = sCurrent
    lstListBox.AddItem sCurrent, i
    i = i + 1
    Loop
    Close #1
    Exit Sub
dlgerror:
    MsgBox "An error has occured " & Err.Description
    Exit Sub
End Sub

Private Sub SaveList(sLocation As String, lstListBox As ListBox)
    On Error GoTo dlgerror
    Dim sCurrent As String
    Dim i As Integer
    Open sLocation For Output As #1
    i = 0
    Do Until i = lstListBox.ListCount
    sCurrent = lstListBox.List(i)
    Print #1, sCurrent
    i = i + 1
    Loop
    Close #1
    Exit Sub
dlgerror:
    MsgBox "An error has occured " & Err.Description
    Exit Sub
End Sub

Private Sub VBHotKey1_HotkeyPressed()
    HideGames
End Sub

Private Sub VBHotKey2_HotkeyPressed()
    ShowGames
End Sub

Private Sub VBHotKey3_HotkeyPressed()
    Me.Show
End Sub
