Attribute VB_Name = "HotkeySubclass"
Option Explicit

Public hOldWndProc As Long

Public Const WM_HOTKEY = &H312

'\\ API Calls used in subclassing windows....
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'\\ Window specific information
Private Declare Function GetWindowLongApi Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLongApi Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_WNDPROC = (-4)
'\\ API Error decoding
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

Public colSubclassedwindows As New Collection
Public colControls As New Collection

'\\ -- [ LastSystemError ]------------------------------------------------------------------
'\\ Returns the message from the system which describes the last dll error to occur, as
'\\ held in Err.LastDllError.  This function should be called as soon after the API call
'\\ which might have errored, as this member can be reset to zero by subsequent API calls.
'\\ ----------------------------------------------------------------------------------------
'\\ You have a royalty free right to use, reproduce, modify, publish and mess with this code
'\\ I'd like you to visit http://www.merrioncomputing.com for updates, but won't force you
'\\ ----------------------------------------------------------------------------------------
Public Function LastSystemError() As String

Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Dim sError As String * 500 '\\ Preinitilise a string buffer to put any error message into
Dim lErrNum As Long
Dim lErrMsg As Long

lErrNum = Err.LastDllError

lErrMsg = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, lErrNum, 0, sError, Len(sError), 0)

LastSystemError = Trim(sError)

End Function
Public Sub FreeSubclassedWindow(ByVal hwnd As Long)

On Error Resume Next
Dim ctlThis As VBHotKey

With colSubclassedwindows.Item("hwnd:" & hwnd)
    Call SetWindowLongApi(hwnd, GWL_WNDPROC, .OldWndProc)
End With
Call colSubclassedwindows.Remove("hwnd:" & hwnd)

For Each ctlThis In colControls
    '\\ Remove any hotkeys on this form...
    If InStr(ctlThis.UniqueKey, "hwnd:" & hwnd) Then
        Call colControls.Remove(ctlThis.UniqueKey)
    End If
Next ctlThis

If colSubclassedwindows.Count = 0 Then
    Set colSubclassedwindows = Nothing
End If
If colControls.Count = 0 Then
    Set colControls = Nothing
End If

End Sub



Private Function LongFromLong(ByVal lIn As Long) As Long

LongFromLong = lIn

End Function

Public Function SubclassWindow(ByVal hwnd As Long)

Dim lRet As Long
Dim wndThis As ApiSubclassedWindow

Set wndThis = New ApiSubclassedWindow
'\\ DEJ 18 June 2001 - Only subclass each window once
If LongFromLong(AddressOf VB_WindowProc) <> GetWindowLongApi(hwnd, GWL_WNDPROC) Then
    With wndThis
        .hwnd = hwnd
        .OldWndProc = GetWindowLongApi(hwnd, GWL_WNDPROC)
        If SetWindowLongApi(hwnd, GWL_WNDPROC, AddressOf VB_WindowProc) > 0 Then
            .NewWndProc = GetWindowLongApi(hwnd, GWL_WNDPROC)
        End If
        If .OldWndProc <> .NewWndProc Then
            Debug.Print "Window : " & .hwnd & " subclassed OK"
        Else
            Debug.Print "Window : " & .hwnd & " subclassin failed"
        End If
    End With
    colSubclassedwindows.Add wndThis, "hwnd:" & wndThis.hwnd
End If

Set wndThis = Nothing

End Function

'\\ --[VB_WindowProc]-------------------------------------------------------------------
'\\ 'typedef LRESULT (CALLBACK* WNDPROC)(HWND, UINT, WPARAM, LPARAM);
'\\ Parameters:
'\\   hwnd - window handle receiving message
'\\   wMsg - The window message (WM_..etc.)
'\\   wParam - First message parameter
'\\   lParam - Second message parameter
'\\ Note:
'\\    When subclassing a window proc using this, set the global
'\\    hOldWndProc property to the window's previous window proc address.
'\\ ----------------------------------------------------------------------------------------
'\\ You have a royalty free right to use, reproduce, modify, publish and mess with this code
'\\ I'd like you to visit http://www.merrioncomputing.com for updates, but won't force you
'\\ ----------------------------------------------------------------------------------------
Public Function VB_WindowProc(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

On Local Error Resume Next
Dim VKey As Long
Dim Modifier As Long
Dim sWindow As ApiSubclassedWindow
Dim ctlThis As VBHotKey

Set sWindow = colSubclassedwindows.Item("hwnd:" & hwnd)


If wMsg = WM_HOTKEY Then
    Debug.Print "WM_HOTKEY"
    VKey = HiWord(lParam)
    Modifier = LoWord(lParam)
    Set ctlThis = colControls.Item("hwnd:" & hwnd & "VK:" & VKey & "Mod:" & Modifier)
    Call ctlThis.RaiseKeyPressEvent(VKey, Modifier)
End If

If sWindow.OldWndProc = 0 Then
    VB_WindowProc = DefWindowProc(hwnd, wMsg, wParam, lParam)
Else
    VB_WindowProc = CallWindowProc(sWindow.OldWndProc, hwnd, wMsg, wParam, lParam)
End If


End Function

'\\ --[HiWord]-----------------------------------------------------------------------------
'\\ Returns the high word component of a long value
'\\ Parameters:
'\\   dw - The long of which we need the HiWord
'\\
'\\ ----------------------------------------------------------------------------------------
'\\ You have a royalty free right to use, reproduce, modify, publish and mess with this code
'\\ I'd like you to visit http://www.merrioncomputing.com for updates, but won't force you
'\\ ----------------------------------------------------------------------------------------
Public Function HiWord(dw As Long) As Integer
 If dw And &H80000000 Then
      HiWord = (dw \ 65535) - 1
 Else
    HiWord = dw \ 65535
 End If
End Function

'\\ --[LoWord]-----------------------------------------------------------------------------
'\\ Returns the low word component of a long value
'\\ Parameters:
'\\   dw - The long of which we need the LoWord
'\\
'\\ ----------------------------------------------------------------------------------------
'\\ You have a royalty free right to use, reproduce, modify, publish and mess with this code
'\\ I'd like you to visit http://www.merrioncomputing.com for updates, but won't force you
'\\ ----------------------------------------------------------------------------------------
Public Function LoWord(dw As Long) As Integer
  If dw And &H8000& Then
      LoWord = &H8000 Or (dw And &H7FFF&)
   Else
      LoWord = dw And &HFFFF&
   End If
End Function
