VERSION 5.00
Object = "{4881A3EC-DC21-11D4-8235-0010A4C42ABD}#32.33#0"; "ExtLVCTL.ocx"
Begin VB.Form frmMain 
   Caption         =   "360 Game ID Lookup - Powered by http://360.kingla.com"
   ClientHeight    =   7455
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   10935
   LinkTopic       =   "Form1"
   ScaleHeight     =   7455
   ScaleWidth      =   10935
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraGameIDs 
      Caption         =   "Game IDs"
      Height          =   5655
      Left            =   3240
      TabIndex        =   3
      Top             =   120
      Width           =   5535
      Begin VB.ComboBox cmbGameShow 
         Height          =   315
         ItemData        =   "frmMain.frx":0000
         Left            =   120
         List            =   "frmMain.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   5295
      End
      Begin ExtLVCTL.ExtLV lstGames 
         Height          =   4935
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   8705
         LineEvenColor   =   0
         LineOddColor    =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "Verdana"
         FontSize        =   8.25
         FullRowSelect   =   -1  'True
         View            =   3
         ListIndex       =   -1
         CalendarTrailingForeColor=   -2147483631
         CalendarTitleForeColor=   -2147483630
         CalendarTitleBackColor=   -2147483633
         CalendarForeColor=   -2147483630
         CalendarBackColor=   -2147483643
         TitleHeight     =   255
         PlaySounds      =   0   'False
         DropWidth       =   0
         DropLines       =   0
         DropDelay       =   0
         SortArrows      =   -1  'True
         TabCaptions     =   ""
         SortArrowSize   =   0
      End
      Begin ExtLVCTL.ExtLV lstUnknownGames 
         Height          =   4935
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Visible         =   0   'False
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   8705
         LineEvenColor   =   0
         LineOddColor    =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "Verdana"
         FontSize        =   8.25
         FullRowSelect   =   -1  'True
         View            =   3
         ListIndex       =   -1
         CalendarTrailingForeColor=   -2147483631
         CalendarTitleForeColor=   -2147483630
         CalendarTitleBackColor=   -2147483633
         CalendarForeColor=   -2147483630
         CalendarBackColor=   -2147483643
         TitleHeight     =   255
         PlaySounds      =   0   'False
         DropWidth       =   0
         DropLines       =   0
         DropDelay       =   0
         TabCaptions     =   ""
         SortArrowSize   =   0
      End
   End
   Begin VB.Frame fraDirBrowser 
      Caption         =   "Directory Browser"
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2775
      End
      Begin VB.DirListBox Dir1 
         Height          =   4815
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   2775
      End
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "&Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuExplorerFind 
         Caption         =   "&Find in Explorer"
      End
      Begin VB.Menu mnuExplorerOpen 
         Caption         =   "&Open in Explorer"
      End
   End
   Begin VB.Menu mnuUnknown 
      Caption         =   "mnuUnknown"
      Visible         =   0   'False
      Begin VB.Menu mnuUnknownEnterID 
         Caption         =   "&Enter ID"
      End
      Begin VB.Menu mnuUnknownFind 
         Caption         =   "&Find in Explorer"
      End
      Begin VB.Menu mnuUnknownOpen 
         Caption         =   "&Open in Explorer"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Function DirExists(ByVal DName As String) As Boolean
    Dim sDummy As String
    
    On Error Resume Next
    
    If Right(DName, 1) <> "\" Then DName = DName & "\"
    sDummy = Dir$(DName & "*.*", vbDirectory)
    DirExists = Not (sDummy = "")
End Function


Private Sub cmbGameShow_Click()
    If cmbGameShow.Text = cmbGameShow.List(0) Then
        lstGames.Visible = True
        lstUnknownGames.Visible = False
    ElseIf cmbGameShow.Text = cmbGameShow.List(1) Then
        lstGames.Visible = False
        lstUnknownGames.Visible = True
    End If
End Sub

Private Sub Dir1_Change()
    Dim GameID As String
    Dim i As Integer
    'Dim lItem As ListItem
    Dim lItem As Object
    
    lstGames.ListItems.Clear
    lstUnknownGames.ListItems.Clear
    
    For i = 0 To Dir1.ListCount - 1
        'Dir1.ListIndex = i
        GameID = Right(Dir1.List(i), Len(Dir1.List(i)) - Len(Dir1.Path) - 1)
        If Len(ReadINI(GameID, "N", App.Path & "\user.ini")) > 0 Then ' Trust the user.ini before the site games.ini
            Set lItem = lstGames.ListItems.Add(, , GameID)

            If ReadINI(GameID, "T", App.Path & "\user.ini") = "R" Then
                lItem.ListSubItems.Add , , "Retail"
            ElseIf ReadINI(GameID, "T", App.Path & "\user.ini") = "A" Then
                lItem.ListSubItems.Add , , "Arcade":
            ElseIf ReadINI(GameID, "T", App.Path & "\user.ini") = "C" Then
                lItem.ListSubItems.Add , , "Content"
            Else
                lItem.ListSubItems.Add , , "Unknown"
            End If

            lItem.ListSubItems.Add , , ReadINI(GameID, "N", App.Path & "\user.ini")
            lItem.ListSubItems.Add , , ReadINI(GameID, "U", App.Path & "\user.ini")
        ElseIf Len(ReadINI(GameID, "N", App.Path & "\games.ini")) > 0 Then
            Set lItem = lstGames.ListItems.Add(, , GameID)
            
            If ReadINI(GameID, "T", App.Path & "\games.ini") = "R" Then
                lItem.ListSubItems.Add , , "Retail"
            ElseIf ReadINI(GameID, "T", App.Path & "\games.ini") = "A" Then
                lItem.ListSubItems.Add , , "Arcade"
            ElseIf ReadINI(GameID, "T", App.Path & "\games.ini") = "C" Then
                lItem.ListSubItems.Add , , "Content"
            Else
                lItem.ListSubItems.Add , , "Unknown"
            End If
            
            lItem.ListSubItems.Add , , ReadINI(GameID, "N", App.Path & "\games.ini")
            lItem.ListSubItems.Add , , ReadINI(GameID, "U", App.Path & "\games.ini")
        ElseIf Len(GameID) = 8 Then
            lstUnknownGames.ListItems.Add , , GameID
        End If
    Next i
    
    SaveSetting "360GameID", "UserData", "LastDir", Dir1.Path
End Sub

Private Sub Dir1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Shell "explorer " & """" & Dir1.Path & "\" & """", vbNormalFocus
    End If
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
'   MsgBox ReadINI("415607E1", "N", App.Path & "\games.ini")
'   MsgBox ReadINI("41568813", "N", App.Path & "\games.ini")
   
    lstGames.ColumnHeaders.Add , "ID", "ID"
    lstGames.ColumnHeaders.Add , "Type", "Type"
    lstGames.ColumnHeaders.Add , "Name", "Name"
    lstGames.ColumnHeaders.Add , "Submitted By", "Submitted By"
        
    lstUnknownGames.ColumnHeaders.Add , "ID", "Game ID"
    lstUnknownGames.ColumnHeaders.Add , "Type", "Type"
    lstUnknownGames.ColumnHeaders.Add , "Name", "Game Name"
    
    If DirExists(GetSetting("360GameID", "UserData", "LastDir")) Then
        Dir1.Path = GetSetting("360GameID", "UserData", "LastDir")
    End If
    Drive1.Drive = Left(Dir1.Path, 2)
    
    Me.Height = GetSetting("360GameID", "UserData", "ProgramHeight", Me.Height)
    Me.Width = GetSetting("360GameID", "UserData", "ProgramWidth", Me.Width)
    
    cmbGameShow.ListIndex = 0
    
    'MsgBox fraDirBrowser.Height - Dir1.Height      '730
    fraDirBrowser.Height = frmMain.Height - 730
    
    'MsgBox frmMain.Height - fraDirBrowser.Height   '840
    Dir1.Height = fraDirBrowser.Height - 840
    
    ' Frame Game ID Height
    'MsgBox frmMain.Height - fraGameIDs.Height       '795
    fraGameIDs.Height = frmMain.Height - 795
    
    ' Frame Game ID Width
    'MsgBox frmMain.Width - fraGameIDs.Width       '3510
    fraGameIDs.Width = frmMain.Width - 3510
        
    ' Game List Height
    'MsgBox fraGameIDs.Height - lstGames.Height      '720
    lstGames.Height = fraGameIDs.Height - 720
    
    ' Game List Width
    'MsgBox fraGameIDs.Width - lstGames.Width      '240
    lstGames.Width = fraGameIDs.Width - 240
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveSetting "360GameID", "UserData", "ProgramHeight", Me.Height
    SaveSetting "360GameID", "UserData", "ProgramWidth", Me.Width
End Sub

Private Sub Form_Resize()
    ' Directry Browser Frame
    'MsgBox frmMain.Height - fraDirBrowser.Height ' 730
    fraDirBrowser.Height = frmMain.Height - 730
    
    ' Directroy Browser Height
    'MsgBox fraDirBrowser.Height - Dir1.Height ' 905
    Dir1.Height = fraDirBrowser.Height - 905
    
    ' Frame Game ID Height
    'MsgBox frmMain.Height - fraGameIDs.Height '795
    fraGameIDs.Height = frmMain.Height - 795
    
    ' Frame Game ID Width
    'MsgBox frmMain.Width - fraGameIDs.Width ' 3510
    fraGameIDs.Width = frmMain.Width - 3510
        
    ' Game List Height
    'MsgBox fraGameIDs.Height - lstGames.Height      '720
    lstGames.Height = fraGameIDs.Height - 720
    lstUnknownGames.Height = fraGameIDs.Height - 720
    
    ' Game List Width
    'MsgBox fraGameIDs.Width - lstGames.Width      '240
    lstGames.Width = fraGameIDs.Width - 240
    lstUnknownGames.Width = fraGameIDs.Width - 240
    
    lstGames.SizeColumns
End Sub

Private Sub lstGames_DblClick()
    mnuExplorerOpen_Click
End Sub

Private Sub lstGames_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And lstGames.ListItems.Count > 0 Then
        Me.PopupMenu mnuMenu, , X + 3400, Y + 400
    End If
End Sub

Private Sub lstUnknownGames_DblClick()
    mnuExplorerOpen_Click
End Sub

Private Sub lstUnknownGames_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And lstUnknownGames.ListItems.Count > 0 Then
        Me.PopupMenu mnuUnknown, , X + 3400, Y + 400
    End If
End Sub

Private Sub mnuExplorerFind_Click()
    If lstUnknownGames.Visible = True Then
        Shell "explorer /select," & """" & Dir1.Path & "\" & lstUnknownGames.SelectedItem.Text & """", vbNormalFocus
    Else
        Shell "explorer /select," & """" & Dir1.Path & "\" & lstGames.SelectedItem.Text & """", vbNormalFocus
    End If
End Sub

Private Sub mnuExplorerOpen_Click()
    If lstUnknownGames.Visible = True Then
        Shell "explorer " & """" & Dir1.Path & "\" & lstUnknownGames.SelectedItem.Text & """", vbNormalFocus
    Else
        Shell "explorer " & """" & Dir1.Path & "\" & lstGames.SelectedItem.Text & """", vbNormalFocus
    End If
End Sub

Private Sub mnuUnknownEnterID_Click()
    Dim tmpName As String
    Dim tmpUser As String
    Dim lItem As Object
    
    tmpName = InputBox("Enter the title for GameID " & lstUnknownGames.SelectedItem.Text, "Enter Title", "")
        
    tmpUser = InputBox("Enter your name", "Enter your name", GetSetting("360GameID", "UserData", "Username"))
    SaveSetting "360GameID", "UserData", "Username", tmpUser
    
    If Len(tmpName) > 0 Then
        WriteINI lstUnknownGames.SelectedItem.Text, "N", tmpName, App.Path & "\user.ini"
        WriteINI lstUnknownGames.SelectedItem.Text, "U", tmpUser, App.Path & "\user.ini"
        
        Set lItem = lstGames.ListItems.Add(, , lstUnknownGames.SelectedItem.Text)
            lItem.ListSubItems.Add , , tmpName
            lItem.ListSubItems.Add , , tmpUser
            
        lstUnknownGames.RemoveItem lstUnknownGames.SelectedItem.Index
    End If
End Sub

Private Sub mnuUnknownFind_Click()
    mnuExplorerFind_Click
End Sub

Private Sub mnuUnknownOpen_Click()
    mnuExplorerOpen_Click
End Sub
