VERSION 5.00
Begin VB.Form frmWLoader 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WoW Server Loader $VER$ by Agret <alias.zero2097@gmail.com>"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4575
   Icon            =   "frmWLoader.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   4575
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optWoW 
      Caption         =   "WoW.exe"
      Height          =   255
      Left            =   2520
      TabIndex        =   18
      Top             =   4640
      Value           =   -1  'True
      Width           =   2295
   End
   Begin VB.OptionButton optLauncher 
      Caption         =   "Launcher.exe"
      Height          =   255
      Left            =   840
      TabIndex        =   17
      Top             =   4640
      Width           =   4335
   End
   Begin VB.CommandButton cmdLaunchOfficial 
      Caption         =   "Launch WoW"
      Height          =   735
      Left            =   2400
      TabIndex        =   10
      Top             =   4920
      Width           =   2055
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdLaunchCustom 
      Caption         =   "Launch WoW"
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   4920
      Width           =   2055
   End
   Begin VB.Frame fraServers 
      Caption         =   "Servers"
      Height          =   4455
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   4335
      Begin VB.TextBox txtWebsite 
         Height          =   285
         Left            =   840
         TabIndex        =   2
         Top             =   2280
         Width           =   3375
      End
      Begin VB.CommandButton cmdWebsite 
         Caption         =   "View"
         Height          =   255
         Left            =   3480
         TabIndex        =   3
         Top             =   2640
         Width           =   735
      End
      Begin VB.CommandButton cmdRegister 
         Caption         =   "View"
         Height          =   255
         Left            =   3480
         TabIndex        =   5
         Top             =   3360
         Width           =   735
      End
      Begin VB.TextBox txtRegisterURL 
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Top             =   3000
         Width           =   3015
      End
      Begin VB.ListBox lstServers 
         Height          =   1230
         Left            =   120
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   240
         Width           =   4095
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   840
         TabIndex        =   0
         Top             =   1560
         Width           =   3375
      End
      Begin VB.TextBox txtAddress 
         Height          =   285
         Left            =   840
         TabIndex        =   1
         Top             =   1920
         Width           =   3375
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   2880
         TabIndex        =   8
         Top             =   3840
         Width           =   1215
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   3840
         Width           =   1215
      End
      Begin VB.Label lblWebsite 
         BackStyle       =   0  'Transparent
         Caption         =   "Website:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label lblRegister 
         BackStyle       =   0  'Transparent
         Caption         =   "Register URL:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label lblAddress 
         BackStyle       =   0  'Transparent
         Caption         =   "Address:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1920
         Width           =   1215
      End
   End
   Begin VB.Label lblXFireSettings 
      BackStyle       =   0  'Transparent
      Caption         =   "Click here if you use XFire and would like to make all your WoW installs detect properly =]"
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   5760
      Width           =   4335
   End
End
Attribute VB_Name = "frmWLoader"
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
    
Private Sub LoadList(sLocation As String, lstListBox As ListBox)
    On Error GoTo dlgerror
    Dim sCurrent As String
    Dim i As Integer
    lstListBox.Clear
    Open sLocation For Input As #1
    i = 0
    Do Until EOF(1)
    Line Input #1, sCurrent
    lstListBox.AddItem Replace(sCurrent, "|", vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "|", , 1), i
    i = i + 1
    Loop
    Close #1
    Exit Sub
dlgerror:
    If Err.Number <> 53 Then
        MsgBox "An error has occured: " & Err.Description
    End If
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
    Print #1, Replace(sCurrent, vbTab, "")
    i = i + 1
    Loop
    Close #1
    Exit Sub
dlgerror:
    MsgBox "An error has occured: " & Err.Description
    Exit Sub
End Sub

Private Sub cmdAdd_Click()
    If txtName.Text = "" Then txtName.Text = txtAddress.Text
    
    If txtAddress.Text = "" Then
        MsgBox "You must enter a server address", vbCritical, "Error"
    Else
        lstServers.AddItem Replace(txtName.Text, "|", "-") & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "|" & txtAddress.Text & "|" & txtWebsite.Text & "|" & txtRegisterURL.Text
    End If
End Sub

Private Sub cmdDel_Click()
    If lstServers.ListIndex > -1 Then
        lstServers.RemoveItem lstServers.ListIndex
    End If
End Sub

Private Sub cmdEdit_Click()
    If lstServers.ListIndex > -1 Then
        cmdDel_Click
        cmdAdd_Click
    End If
End Sub

Private Sub cmdLaunchCustom_Click()
    Dim ErrorHandled As Boolean
    Dim WoWPath As String
    Dim FakeWoWPath As String
    On Error GoTo errorhandle
    ErrorHandled = False
    WoWPath = App.Path
    If lstServers.ListIndex > -1 Then
        Dim ServerInfo() As String
        ServerInfo() = Split(lstServers.Text, "|")

        Dim Question As Integer
        Question = MsgBox("Launch WoW Custom?" & vbNewLine & "-----------------" & vbNewLine & "Server Info:" & vbNewLine & "-----------------" & vbNewLine & "Name: " & Replace(ServerInfo(0), vbTab, "") & vbNewLine & "Address: " & ServerInfo(1), vbYesNo)
        If Question = 6 Then
StartOver:
            If FileExists(WoWPath & "\realmlist.wtf") Then
                Open WoWPath & "\realmlist.wtf" For Output As #4
                    Print #4, "set realmlist " & ServerInfo(1)
                Close #4
            End If

            If FileExists(WoWPath & "\Data\enUS\realmlist.wtf") Then
                Open WoWPath & "\Data\enUS\realmlist.wtf" For Output As #5
                    Print #5, "set realmlist " & ServerInfo(1)
                Close #5
            End If
            
            ' XFire Check
            If GetSetting("WoW Loader", "Pref", "UseXFire") = "True" Then
                FakeWoWPath = GetSetting("WoW Loader", "Pref", "XFireFakeFolder")
                
                If FileExists(WoWPath & "\WoW.exe") Then
                    Open App.Path & "\tmp.bat" For Output As #6
                        Print #6, """" & App.Path & "\junction.exe" & """" & " " & """" & FakeWoWPath & """" & " " & """" & WoWPath & """"
                        Print #6, Left(FakeWoWPath, 2)
                        Print #6, "cd " & """" & FakeWoWPath & """"
                        Print #6, "start /WAIT WoW.exe"
                        Print #6, "regedit /S " & """" & App.Path & "\tmp.reg" & """"
                        Print #6, "del " & """" & App.Path & "\tmp.reg" & """"
                        Print #6, """" & App.Path & "\junction.exe" & """" & " -d " & """" & FakeWoWPath & """"
                        Print #6, "del " & """" & App.Path & "\tmp.bat" & """"
                    Close #6
                    Open App.Path & "\tmp.reg" For Output As #7
                        Print #7, "Windows Registry Editor Version 5.00" & vbCrLf & "[HKEY_LOCAL_MACHINE\SOFTWARE\Blizzard Entertainment\World of Warcraft]" & vbCrLf & Replace("""" & "InstallPath" & """" & "=" & """" & WoWPath & "\" & """", "\", "\\")
                    Close #7
                    Shell App.Path & "\tmp.bat", vbHide
                Else
                    Err.Raise 53
                End If
            Else ' No XFire Support, Use Classic Launch Method
                Shell WoWPath & "\WoW.exe", vbNormalFocus
            End If
            
        End If
    End If
    Exit Sub
errorhandle:
    If Err.Number = 53 Then
        If ErrorHandled = False Then
                ErrorHandled = True
                WoWPath = Left(GetString(HKEY_LOCAL_MACHINE, "SOFTWARE\Blizzard Entertainment\World of Warcraft", "InstallPath"), Len(GetString(HKEY_LOCAL_MACHINE, "SOFTWARE\Blizzard Entertainment\World of Warcraft", "InstallPath")) - 1)
                GoTo StartOver
            Else
                MsgBox "There was an error launching WoW." & vbNewLine & "Please put launcher into your WoW folder.", vbCritical, "Error"
        End If
    End If
End Sub


Private Sub cmdLaunchOfficial_Click()
    Dim ErrorHandled As Boolean
    Dim WoWPath As String
    On Error GoTo errorhandle
    ErrorHandled = False
    WoWPath = App.Path
    
    Dim Question As Integer
    Question = MsgBox("Launch WoW Official?", vbYesNo)
    If Question = 6 Then
StartOver:
            If FileExists(WoWPath & "\realmlist.wtf") Then
                Open WoWPath & "\realmlist.wtf" For Output As #4
                    Print #4, "set realmlist us.logon.worldofwarcraft.com"
                    Print #4, "set patchlist us.version.worldofwarcraft.com"
                Close #4
            End If
            
            If FileExists(WoWPath & "\Data\enUS\realmlist.wtf") Then
                Open WoWPath & "\Data\enUS\realmlist.wtf" For Output As #5
                    Print #5, "set realmlist us.logon.worldofwarcraft.com"
                    Print #5, "set patchlist us.version.worldofwarcraft.com"
                    Print #5, "set realmlistbn " & """" & """"""
                    Print #5, "set portal " & """" & "us" & """"
                Close #5
            End If
            
            
            ' XFire Check
            If GetSetting("WoW Loader", "Pref", "UseXFire") = "True" Then
                FakeWoWPath = GetSetting("WoW Loader", "Pref", "XFireFakeFolder")
                
                If optLauncher.Value = True Then
                    MsgBox "You may not use XFire detection mode and launcher at the same time." & vbNewLine & "Launching without XFire support", vbInformation, "Oh No!!"
                    Shell WoWPath & "\Launcher.exe", vbNormalFocus
                Else
                    If FileExists(WoWPath & "\WoW.exe") Then
                        Open App.Path & "\tmp.bat" For Output As #6
                            Print #6, """" & App.Path & "\junction.exe" & """" & " " & """" & FakeWoWPath & """" & " " & """" & WoWPath & """"
                            Print #6, Left(FakeWoWPath, 2)
                            Print #6, "cd " & """" & FakeWoWPath & """"
                            Print #6, "start /WAIT WoW.exe"
                            Print #6, "regedit /S " & """" & App.Path & "\tmp.reg" & """"
                            Print #6, "del " & """" & App.Path & "\tmp.reg" & """"
                            Print #6, """" & App.Path & "\junction.exe" & """" & " -d " & """" & FakeWoWPath & """"
                            Print #6, "del " & """" & App.Path & "\tmp.bat" & """"
                        Close #6
                        Open App.Path & "\tmp.reg" For Output As #7
                            Print #7, "Windows Registry Editor Version 5.00" & vbCrLf & "[HKEY_LOCAL_MACHINE\SOFTWARE\Blizzard Entertainment\World of Warcraft]" & vbCrLf & Replace("""" & "InstallPath" & """" & "=" & """" & WoWPath & "\" & """", "\", "\\")
                        Close #7
                        Shell App.Path & "\tmp.bat", vbHide
                    Else
                        Err.Raise 53
                    End If
                End If
            Else ' No XFire Support, Use Classic Launch Method
                If optLauncher.Value = True Then
                    Shell WoWPath & "\Launcher.exe", vbNormalFocus
                Else
                    Shell WoWPath & "\WoW.exe", vbNormalFocus
                End If
            End If
    End If
    
    Exit Sub
errorhandle:
    If Err.Number = 53 Then
        If ErrorHandled = False Then
                ErrorHandled = True
                WoWPath = Left(GetString(HKEY_LOCAL_MACHINE, "SOFTWARE\Blizzard Entertainment\World of Warcraft", "InstallPath"), Len(GetString(HKEY_LOCAL_MACHINE, "SOFTWARE\Blizzard Entertainment\World of Warcraft", "InstallPath")) - 1)
                GoTo StartOver
            Else
                MsgBox "There was an error launching WoW." & vbNewLine & "Please put launcher into your WoW folder.", vbCritical, "Error"
        End If
    End If
End Sub

Private Sub cmdRegister_Click()
    ShellExecute hwnd, "open", txtRegisterURL.Text, vbNullString, vbNullString, 1
End Sub

Private Sub cmdWebsite_Click()
    ShellExecute hwnd, "open", txtWebsite.Text, vbNullString, vbNullString, 1
End Sub

Private Sub Form_Load()
    Me.Caption = Replace(Me.Caption, "$VER$", "[" & App.Major & "." & App.Minor & "." & App.Revision & "]")
    LoadList App.Path & "\servers.txt", lstServers
    cmdLaunchCustom.Caption = "Launch WoW" & vbNewLine & "(Custom Server)"
    cmdLaunchOfficial.Caption = "Launch WoW" & vbNewLine & "(Official Server)"
    If GetSetting("WoW Loader", "Pref", "LaunchEXE") <> "" Then
        If GetSetting("WoW Loader", "Pref", "LaunchEXE") = "Launcher.exe" Then
            optLauncher.Value = True
            optWoW.Value = False
        Else
            optLauncher.Value = False
            optWoW.Value = True
        End If
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If optLauncher.Value = True Then
        SaveSetting "WoW Loader", "Pref", "LaunchEXE", "Launcher.exe"
    Else
        SaveSetting "WoW Loader", "Pref", "LaunchEXE", "WoW.exe"
    End If
    SaveList App.Path & "\servers.txt", lstServers
End Sub

Private Sub lblXFireSettings_Click()
    Load frmXFire
    frmXFire.Visible = True
End Sub

Private Sub lstServers_Click()
    If lstServers.ListIndex > -1 Then
        Dim ServerInfo() As String
        ServerInfo() = Split(lstServers.Text, "|")
        txtName.Text = Replace(ServerInfo(0), vbTab, "")
        txtAddress.Text = ServerInfo(1)
        If UBound(ServerInfo) > 1 Then
            txtWebsite.Text = ServerInfo(2)
            txtRegisterURL.Text = ServerInfo(3)
        Else
            txtWebsite.Text = ""
            txtRegisterURL.Text = ""
        End If
    End If
End Sub

Private Sub lstServers_DblClick()
    cmdLaunchCustom_Click
End Sub
