Attribute VB_Name = "Module1"
Option Explicit
Public GameDir As String
Public cache As String
Public rundefclient As Boolean
Public customclient As String
Public othpath As Boolean
Public useIP As String
Public useName As String
Public usePort As String
Public IPlist As String
Public isEdit As Boolean
Public conState As Integer
' connection state:
' 1:Conencted, but not authenticated
' 2:Connected, authenticating
' 3:Connected, Authenticated

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL As Long = 1
Private Const SW_SHOWMINIMIZED As Long = 2
Private Const SW_SHOWMAXIMIZED As Long = 3
Public Const HKEY_CURRENT_USER = &H80000002
Public Const REG_SZ = 1 ' Unicode nul terminated string
Public Const REG_BINARY = 3 ' Free form binary
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Public Function GetString(hKey As Long, strPath As String, strValue As String)
    Dim Ret
    'Open the key
    RegOpenKey hKey, strPath, Ret
    'Get the key's content
    GetString = RegQueryStringValue(Ret, strValue)
    'Close the key
    RegCloseKey Ret
End Function

Public Function RegQueryStringValue(ByVal hKey As Long, ByVal strValueName As String) As String
    Dim lResult As Long, lValueType As Long, strBuf As String, lDataBufSize As Long
    'retrieve nformation about the key
    lResult = RegQueryValueEx(hKey, strValueName, 0, lValueType, ByVal 0, lDataBufSize)
    If lResult = 0 Then
        If lValueType = REG_SZ Then
            'Create a buffer
            strBuf = String(lDataBufSize, Chr$(0))
            'retrieve the key's content
            lResult = RegQueryValueEx(hKey, strValueName, 0, 0, ByVal strBuf, lDataBufSize)
            If lResult = 0 Then
                'Remove the unnecessary chr$(0)'s
                RegQueryStringValue = Left$(strBuf, InStr(1, strBuf, Chr$(0)) - 1)
            End If
        ElseIf lValueType = REG_BINARY Then
            Dim strData As Integer
            'retrieve the key's value
            lResult = RegQueryValueEx(hKey, strValueName, 0, 0, strData, lDataBufSize)
            If lResult = 0 Then
                RegQueryStringValue = strData
            End If
        End If
    End If
End Function

Public Sub SaveString(hKey As Long, strPath As String, strValue As String, strData As String)
    Dim Ret
    'Create a new key
    RegCreateKey hKey, strPath, Ret
    'Save a string to the key
    RegSetValueEx Ret, strValue, 0, REG_SZ, ByVal strData, Len(strData)
    'close the key
    RegCloseKey Ret
End Sub
Public Sub SaveStringLong(hKey As Long, strPath As String, strValue As String, strData As String)
    Dim Ret
    'Create a new key
    RegCreateKey hKey, strPath, Ret
    'Set the key's value
    RegSetValueEx Ret, strValue, 0, REG_BINARY, CByte(strData), 4
    'close the key
    RegCloseKey Ret
End Sub
Public Sub DelSetting(hKey As Long, strPath As String, strValue As String)
    Dim Ret
    'Create a new key
    RegCreateKey hKey, strPath, Ret
    'Delete the key's value
    RegDeleteValue Ret, strValue
    'close the key
    RegCloseKey Ret
End Sub

Public Sub RunGame(Parametres As String)
Parametres = Parametres & " "
    'Test.txt resides in "D:\MyStart\In\Path\Directory"
    'Since we specify a start in directory we can use a relative path for the file name
   If rundefclient = True Then ShellExecute frmMain.hwnd, "OPEN", GameDir & "\OTWO.exe", Parametres, GameDir, SW_SHOWNORMAL
    If rundefclient = False Then ShellExecute frmMain.hwnd, "OPEN", GameDir & "\" & customclient, Parametres, GameDir, SW_SHOWNORMAL
End
End Sub


Public Function checkFile() As Boolean
checkFile = True
Dim file1 As String
On Error GoTo 1
Dim cpath As Variant
cpath = App.Path & "\"
Open cpath & "settings.txt" For Input As #1
Input #1, file1
Close #1
Exit Function

1: makefile (cpath & "\settings.txt")
checkFile = False
End Function

Public Sub CreateFolders(strPath)
Dim i As Integer
Dim objFso
Dim arrFolders
Dim strDrive
Dim strFolder
Set objFso = CreateObject("Scripting.FileSystemObject")
' Format path (remove leading and trailing spaces, and final backslash)
strPath = Trim(strPath)
If Right(strPath, 1) = "\" Then
strPath = Left(strPath, Len(strPath) - 1)
End If
' Check that a drive is specified and that it exists (if not, exit sub)
If Mid(strPath, 2, 2) <> ":\" Then
MsgBox "The path specified is invalid."
Exit Sub
Else
strDrive = Left(strPath, 1)
If Not objFso.DriveExists(strDrive) Then
MsgBox "The drive specified does not exist."
Exit Sub
End If
End If
' Split the path into an array, first element is the drive, subsequent elements are the folders
arrFolders = Split(strPath, "\")
' Build the path, folder by folder. If folder doesn't exist, create it
strFolder = arrFolders(0)
For i = LBound(arrFolders) To UBound(arrFolders) - 1
strFolder = strFolder & "\" & arrFolders(i + 1)
If Not objFso.FolderExists(strFolder) Then
objFso.CreateFolder (strFolder)
End If
Next
End Sub

Public Function songlist(ByVal line As Integer) As Variant
'On Error GoTo 1
Dim cpath As Variant, i As Integer, songdetails As Variant
cpath = App.Path & "\save\p2plist.txt"
Open cpath For Input As #1
For i = 0 To line
If Not EOF(1) = True Then
Line Input #1, songdetails
ElseIf EOF(1) = True Then
songdetails = "[end]"
End If
Next
Close
songlist = songdetails
End Function

Public Sub makefile(ByVal file As Variant)
On Error GoTo 2
1:
Close
Open file For Append As #1
Print #1, "'File created on " & Date
Close #1
Exit Sub
2:
Dim b As Variant
Dim i As Integer
For i = 0 To UBound(Split(file, "\")) - 1
b = b & "\" & Split(file, "\")(i)
Next i
b = Right(b, Len(b) - 1)
CreateFolders (b)
GoTo 1
End Sub

Public Sub loadparam()
On Error Resume Next
Dim currentline As String
Dim i As Integer
Open App.Path & "\settings.txt" For Input As #1
Do Until EOF(1) = True
Line Input #1, currentline
Select Case LTrim(LCase(Split(currentline, "=")(0)))
    Case "useip":
    useIP = Split(currentline, "=")(1)
    Case "useport"
    usePort = Split(currentline, "=")(1)
    Case "usename"
    useName = Split(currentline, "=")(1)
    frmMain.Label2 = "Using <" & useName & "> server"
    Case "rundefclient":
    rundefclient = True
    Case "customexe":
    customclient = RTrim(Split(currentline, "=")(1))
    Case "listserver":
    IPlist = Split(currentline, "=")(1)
    Case "username":
    frmMain.txtUser = Split(currentline, "=")(1)
    Case "extpath":
    GameDir = Split(currentline, "=")(1)
    End Select
Loop
Reset
End Sub

Public Sub savesettings()
Open App.Path & "\settings.txt" For Output As #1
'Print #1, "IP=" & frmMain.txtIP.Text & "." & frmMain.txtIP2.Text & "." & frmMain.txtIP3.Text & "." & frmMain.txtIP4.Text
Print #1, "Username=" & frmMain.txtUser
Print #1, "ListServer=" & IPlist
Print #1, "UseIP=" & useIP
Print #1, "UsePort=" & usePort
Print #1, "UseName=" & useName
If rundefclient = True Then
    Print #1, "rundefclient"
End If
If rundefclient = False Then
    Print #1, "customexe=" & customclient
End If
If othpath = True Then
    Print #1, "extpath=" & GameDir
End If
Reset
End Sub

Sub Main()
Dim c As Integer
Dim a As Integer

loadparam

GameDir = GetString(HKEY_CURRENT_USER, "Software\e-games\o2jam", "location")

If Len(IPlist) = 0 Or Len(useName) = 0 Or Len(usePort) = 0 Or Len(useIP) = 0 Then
    MsgBox "Server list not found, please add server manually", vbInformation + vbOKOnly, "Servers"
    frmServers.Show
    Exit Sub
End If

'txtIP.Tag = "15011"
'txtIP.Tag = frmSet.txtPort
If GameDir = "" Then
c = MsgBox("O2jam client not defined, define a client path?", vbYesNo, "Error : Path not found")
    If c = 6 Then
    a = InputBox("O2jam client path undefined, please define the o2jam folder" & vbCrLf _
    & "e.g <C:\Program Files\e-games\O2jam>")
    othpath = True
    GameDir = a
    End If
End If
frmMain.Show
End Sub

Public Sub execRun()
Dim stoploop As Boolean
Dim looptimes As Integer
Dim x As Integer
frmMain.winsckConnect.Close
useIP = Split(useIP, ":")(0)
frmMain.winsckConnect.Connect useIP, usePort
loadState True
frmMain.tmr_connect.Enabled = True
frmMain.tmr_connect.Interval = 500
'With frmMain.winsckConnect
'Do Until stoploop '
'    If .State = sckConnected Then
'    stoploop = True
    
    
'    ElseIf .State = sckConnecting Then
    'frmLoad.bar1.Value = 40
'    End If
 '       looptimes = looptimes + 1
'    If looptimes > 30000 Then
'    x = MsgBox("Error connecting to server, Retry?", vbRetryCancel, "Connection error")
'        If x = 2 Then
'            MsgBox "Connection error", vbOKOnly + vbCritical, "Error"
'            loadState False
'            stoploop = True
'            Exit Do
'        ElseIf x = 4 Then
'            looptimes = 0
'        End If
'    End If
 '   looptimes = looptimes + 1
'frmLoad.bar1.Value = (looptimes / 1000)
'Loop
'End With
   

End Sub
Public Sub loadState(lstate As Boolean)
If lstate = True Then
Unload frmLoad
Load frmLoad
    frmMain.Hide
    frmLoad.Show
ElseIf lstate = False Then
    frmMain.Show
    frmLoad.Hide
End If
End Sub
