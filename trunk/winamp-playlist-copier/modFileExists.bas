Attribute VB_Name = "modFileExists"
Public Declare Function PathFileExists Lib "shlwapi" Alias "PathFileExistsA" (ByVal pszPath As String) As Long

Public Declare Function PathIsDirectory Lib "shlwapi" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long

'Check file existance:
Public Delcare Function FileExists(ByVal sPath As String) As Boolean
      If (PathFileExists(sPath)) And Not (PathIsDirectory(sPath)) Then FileExists = True
End Function

'Check directory (folder) existance:
Public Delcare Function DirExists(ByVal sPath As String) As Boolean
      If (PathFileExists(sPath)) And (PathIsDirectory(sPath)) Then DirExists = True
End Function
