VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "Msinet.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Imageshack.us User Backup"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4575
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   4575
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   4560
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdDownload 
      Caption         =   "Download All Images"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   5280
      Width           =   4335
   End
   Begin VB.TextBox txtImgURL 
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   4200
      Width           =   4335
   End
   Begin VB.ListBox lstImages 
      Height          =   2010
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   4335
   End
   Begin InetCtlsObjects.Inet InetFTP 
      Left            =   3840
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Frame fraLogin 
      Caption         =   "Login"
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.CommandButton cmdLogin 
         Caption         =   "Login"
         Height          =   375
         Left            =   3000
         TabIndex        =   5
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   1080
         Width           =   4095
      End
      Begin VB.TextBox txtUsername 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   4095
      End
      Begin VB.Label lblPassword 
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   4095
      End
      Begin VB.Label lblUsername 
         BackStyle       =   0  'Transparent
         Caption         =   "Username:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4095
      End
   End
   Begin VB.Label lblProgress 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   5000
      Width           =   4335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim StopDownloads As Boolean

Private Function GetBetween(ByRef TextToParse As String, StartDelimiter As String, EndDelimiter As String) As String
    Dim lngStart As Long, lngend As Long
  
    lngStart = InStr(1, TextToParse$, StartDelimiter$)
    If lngStart Then
        lngStart = lngStart + Len(StartDelimiter$)
        lngend = InStr(lngStart, TextToParse$, EndDelimiter$)
      
        If lngend Then
            GetBetween = Mid$(TextToParse$, lngStart, lngend - lngStart)
        End If
    End If
End Function

' Download a file. Return True if we are successful.
Private Function DownloadFile(ByVal source_file As String, _
    ByVal dest_file As String) As Boolean
Dim bytes() As Byte
Dim fnum As Integer

    ' Get the file's contents.
    On Error GoTo DownloadError
    bytes() = InetFTP.OpenURL(source_file, icByteArray)

    ' Remove the file if it exists.
    On Error Resume Next
    Kill dest_file
    On Error GoTo DownloadError

    ' Write the contents into the destination file.
    fnum = FreeFile
    Open dest_file For Binary Access Write As #fnum
    Put #fnum, , bytes()
    Close #fnum

    DownloadFile = True
    Exit Function

DownloadError:
    MsgBox "Error " & Err.Number & _
        " downloading file '" & _
        source_file & "' to '" & _
        dest_file & "'." & vbCrLf & Err.Description, _
        vbExclamation Or vbOKOnly, _
        "Download Error"
    DownloadFile = False
    Exit Function
End Function

Private Sub cmdDownload_Click()
    If cmdDownload.Caption = "Stop Downloads" Then
        StopDownloads = True
        cmdDownload.Enabled = False
        cmdDownload.Caption = "Download All Images"
    Else
        cmdLogin.Enabled = False
        cmdDownload.Enabled = False
        txtUsername.Locked = True
        txtPassword.Locked = True
        cmdDownload.Caption = "Stop Downloads"
    End If

    Dim i As Integer
    Dim FileName() As String
    
    ProgressBar1.Max = lstImages.ListCount
    ProgressBar1.Min = 0
    
    If DirExists(App.Path & "\images\") <> True Then MkDir App.Path & "\images\"
    If DirExists(App.Path & "\images\" & txtUsername.Text & "\") <> True Then MkDir App.Path & "\images\" & txtUsername.Text & "\"
    
    For i = 0 To lstImages.ListCount - 1
        If StopDownloads = True Then
            StopDownloads = False
            cmdLogin.Enabled = True
            cmdDownload.Enabled = True
            txtUsername.Locked = False
            txtPassword.Locked = False
            MsgBox "Downloads Aborted", vbInformation, "Mission Success"
            Exit Sub
        End If
        lblProgress.Caption = "File " & i & " of " & lstImages.ListCount - 1 & " - " & Round(i / lstImages.ListCount * 100, 2) & "% Complete"
        ProgressBar1.Value = i
        FileName() = Split(lstImages.List(i), "/")
        If FileExists(App.Path & "\images\" & txtUsername.Text & "\" & FileName(UBound(FileName$))) = False Then
            DownloadFile lstImages.List(i), App.Path & "\images\" & txtUsername.Text & "\" & FileName(UBound(FileName$))
        End If
    Next i
    
    cmdLogin.Enabled = True
    cmdDownload.Enabled = True
    txtUsername.Locked = False
    txtPassword.Locked = False
    lblProgress.Caption = "Download Complete."
    MsgBox "Download Complete", vbInformation, "Complete"
    
End Sub

Private Sub cmdLogin_Click()
    cmdLogin.Enabled = False
    cmdDownload.Enabled = False
    txtUsername.Locked = True
    txtPassword.Locked = True
    
    SaveSetting "Imageshack Backup", "Settings", "Username", txtUsername.Text
    SaveSetting "Imageshack Backup", "Settings", "Password", txtPassword.Text
    
    Dim strHTML As String
    Dim ImgServer As String
    Dim ImgURL As String
    Dim ImgName As String
    Dim ParseObjects() As String
    Dim i As Integer
    Dim j As Integer
    
    Debug.Print "Logging in..."
    strHTML = InetFTP.OpenURL("http://my.imageshack.us/auth.php?username=" & txtUsername.Text & "&password=" & txtPassword.Text)
    If strHTML <> "OK" Then
        MsgBox "Login Error" & vbNewLine & "Please check user/pass", vbCritical, "Doh!"
        cmdLogin.Enabled = True
        cmdDownload.Enabled = True
        txtUsername.Locked = False
        txtPassword.Locked = False
        Exit Sub
    Else
        Debug.Print "Logged in."
    End If
    
    Open App.Path & "\urls.htm" For Output As #1
    Open App.Path & "\thumbs.htm" For Output As #2
        For j = 1 To 29
            Debug.Print "Requesting Page " & j & ".."
            strHTML = InetFTP.OpenURL("http://my.imageshack.us/images.php?ipage=" & j)
            Debug.Print "Parsing Page.."
            'txtDebug.Text = strHTML
            ParseObjects() = Split(strHTML, "<div class=""ii"">")
            Debug.Print UBound(ParseObjects$) & " objects"
            For i = LBound(ParseObjects$) + 1 To UBound(ParseObjects$)
                ImgServer$ = GetBetween(ParseObjects(i), "<div class=""is"">", "</div>")
                ImgName$ = GetBetween(ParseObjects(i), "<div class=""if"">", "</div>")
                ImgURL$ = "http://img" & ImgServer$ & ".imageshack.us/img" & ImgServer$ & "/" & GetBetween(ParseObjects(i), "<div class=""ib"">", "</div>") & "/" & ImgName$
                
                lstImages.AddItem "http://img" & ImgServer$ & ".imageshack.us/img" & ImgServer$ & "/" & GetBetween(ParseObjects(i), "<div class=""ib"">", "</div>") & "/" & GetBetween(ParseObjects(i), "<div class=""if"">", "</div>")
                Print #1, "<a href=""" & ImgURL$ & """>" & ImgName$ & "</a>"
                Print #2, "<a href=""" & ImgURL$ & """><img src=""" & ImgThumb(ImgURL$) & """ /></a>&nbsp;"
            Next i
        Next j
    Close #1
    Close #2
    
    cmdLogin.Enabled = True
    cmdDownload.Enabled = True
    txtUsername.Locked = False
    txtPassword.Locked = False
End Sub

Private Function ImgThumb(TheURL As String)
    If InStr(TheURL$, ".png") <> 0 Then
        ImgThumb = Replace(TheURL$, ".png", ".th.png")
    ElseIf InStr(TheURL$, ".jpg") <> 0 Then
        ImgThumb = Replace(TheURL$, ".jpg", ".th.jpg")
    ElseIf InStr(TheURL$, ".jpeg") <> 0 Then
        ImgThumb = Replace(TheURL$, ".jpeg", ".th.jpeg")
    ElseIf InStr(TheURL$, ".gif") <> 0 Then
        ImgThumb = Replace(TheURL$, ".gif", ".th.gif")
    ElseIf InStr(TheURL$, ".bmp") <> 0 Then
        ImgThumb = Replace(TheURL$, ".bmp", ".th.bmp")
    End If
End Function

Private Sub Form_Load()
    txtUsername.Text = GetSetting$("Imageshack Backup", "Settings", "Username", vbNullString)
    txtPassword.Text = GetSetting$("Imageshack Backup", "Settings", "Password", vbNullString)
End Sub

Function DirExists(ByVal DName As String) As Boolean
    Dim sDummy As String
    
    On Error Resume Next
    If Right(DName, 1) <> "\" Then DName = DName & "\"
    sDummy = Dir$(DName & "*.*", vbDirectory)
    DirExists = Not (sDummy = "")
End Function


Public Function FileExists(Fname As String) As Boolean
    If Fname = "" Or Right(Fname, 1) = "\" Then
        FileExists = False: Exit Function
    End If

    FileExists = (Dir(Fname) <> "")
End Function


Private Sub lstImages_Click()
    txtImgURL.Text = lstImages.Text$
End Sub
