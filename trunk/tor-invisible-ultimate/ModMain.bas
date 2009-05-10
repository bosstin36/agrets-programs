Attribute VB_Name = "ModMain"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Well guys here it is, the final and most ultimate way to run Tor hidden '
' ---                                                                     '
' Created on 9th August 2005                                              '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
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

' -----------------
' ADVAPI32
' -----------------
' function prototypes, constants, and type definitions
' for Windows 32-bit Registry API

Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003
Private Const HKEY_PERFORMANCE_DATA = &H80000004
Private Const ERROR_SUCCESS = 0&

' Registry API prototypes

Declare Function RegCloseKey Lib "advapi32.dll" (ByVal Hkey As Long) As Long
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal Hkey As Long, ByVal lpValueName As String) As Long
Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Const REG_SZ = 1                         ' Unicode nul terminated string
Private Const REG_DWORD = 4                      ' 32-bit number
Private Const REG_BINARY = 3                    ' Binary


Sub Main()
    On Error GoTo ErrorHandle ' Shhhh this is quiet!
    
    ' First we see if Tor is running already
    WindowHandle = 0
    WindowHandle = FindWindow(vbNullString, App.Path & "\tor.exe")
    If WindowHandle <> 0 Then
        ShowWindow WindowHandle, SW_HIDE 'Tor is running. Just make sure it's hidden.
    Else
        Shell App.Path & "\tor.exe", vbHide 'Start Tor (hidden, we hope)
    End If
    
    ' Just in case it starts visible this *should* hide it
    WindowHandle = 0
    WindowHandle = FindWindow(vbNullString, App.Path & "\tor.exe")
    If WindowHandle <> 0 Then
        ShowWindow WindowHandle, SW_HIDE
    End If
    
    ' Ok now we have to set our proxy settings.
    ' Lets hope we have access to the registry keys we need.
    DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Internet Settings", "AutoConfigURL"
    SaveString HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Internet Settings", "ProxyServer", "socks=127.0.0.1:9050"
    SaveString HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Internet Settings", "ProxyOverride", "staffweb.dvallcoll.vic.edu.au;studentweb.dvallcoll.vic.edu.au;admin.dvallcoll.vic.edu.au;intranet.dvallcoll.vic.edu.au;mail.dvallcoll.vic.edu.au;students.dvallcoll.vic.edu.au;staff.dvallcoll.vic.edu.au;plato.vtac.edu.au;www.vtac.edu.au;www.sofweb.vic.edu.au;sofweb.vic.edu.au;microsoft.com;www.microsoft.com;<local>"
    DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Internet Settings", "ProxyEnable"
    SaveDword HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Internet Settings", "ProxyEnable", 1
    
    ' And we're outta here!
    ' Enjoy the net with no restrictions :)
    End
    
ErrorHandle:
     
    If Err.Number = 53 Then
        'End 'Tor isn't found, bail!
        Resume Next
    Else
        ' Probarbly something with the registry
        'Resume Next
        MsgBox Err.Number & Err.Description
    End If
    
End Sub

Private Sub SaveKey(Hkey As Long, strPath As String)
Dim keyhand&
r = RegCreateKey(Hkey, strPath, keyhand&)
r = RegCloseKey(keyhand&)
End Sub

Private Function GetString(Hkey As Long, strPath As String, strValue As String)

Dim keyhand As Long
Dim datatype As Long
Dim lResult As Long
Dim strBuf As String
Dim lDataBufSize As Long
Dim intZeroPos As Integer
r = RegOpenKey(Hkey, strPath, keyhand)
lResult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)
If lValueType = REG_SZ Then
    strBuf = String(lDataBufSize, " ")
    lResult = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)
    If lResult = ERROR_SUCCESS Then
        intZeroPos = InStr(strBuf, Chr$(0))
        If intZeroPos > 0 Then
            GetString = Left$(strBuf, intZeroPos - 1)
        Else
            GetString = strBuf
        End If
    End If
End If
End Function

Private Sub SaveString(Hkey As Long, strPath As String, strValue As String, strdata As String)
    Dim keyhand As Long
    Dim r As Long
    r = RegCreateKey(Hkey, strPath, keyhand)
    r = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strdata, Len(strdata))
    r = RegCloseKey(keyhand)
End Sub

Private Function WriteBinaryToRegistry(Hkey As Long, strPath As String, strValue As String, binData As Variant) As Boolean
 
'WRITES A BINARY VALUE TO REGISTRY:
'PARAMETERS:

'Hkey: Top Level Key as defined by
'REG_TOPLEVEL_KEYS Enum (See Declarations)

'strPath - 'Full Path of Subkey
'if path does not exist it will be created

'strValue ValueName

'binData - Value Data

'Returns: True if successful, false otherwise

'EXAMPLE
'Dim v As Variant
'Open "C:\myword.doc" For Binary As #1
'v = Input(LOF(1), #1)
'Close #1
'WriteBinaryToRegistry(HKEY_LOCAL_MACHINE, _
"Software\MySofware", "My Binary Data", v)

Dim bAns As Boolean

On Error GoTo ErrorHandler
   Dim keyhand As Long
   Dim r As Long
   r = RegCreateKey(Hkey, strPath, keyhand)
   If r = 0 Then
        r = RegSetValueEx(keyhand, strValue, 0, REG_BINARY, binData, Len(binData))
        r = RegCloseKey(keyhand)
    End If
    
   WriteBinaryToRegistry = (r = 0)

Exit Function

ErrorHandler:
    WriteBinaryToRegistry = False
    Exit Function
    
End Function

Public Function DeleteKey(ByVal Hkey As Long, ByVal strKey As String)
    Dim r As Long
    r = RegDeleteKey(Hkey, strKey)
End Function

Public Function DeleteValue(ByVal Hkey As Long, ByVal strPath As String, ByVal strValue As String)
    Dim keyhand As Long
    r = RegOpenKey(Hkey, strPath, keyhand)
    r = RegDeleteValue(keyhand, strValue)
    r = RegCloseKey(keyhand)
End Function

Function SaveDword(ByVal Hkey As Long, ByVal strPath As String, ByVal strValueName As String, ByVal lData As Long)
    Dim lResult As Long
    Dim keyhand As Long
    Dim r As Long
    r = RegCreateKey(Hkey, strPath, keyhand)
    lResult = RegSetValueEx(keyhand, strValueName, 0&, REG_DWORD, lData, 4)
    'If lResult <> error_success Then Call errlog("SetDWORD", False)
    r = RegCloseKey(keyhand)
End Function
