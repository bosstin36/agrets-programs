VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApiHotkey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' --[ApiHotkey]--------------------------------------------------------
' This is for global hotkeys
' ---------------------------------------------------------------------

Private Declare Function RegisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Private Declare Function UnregisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long) As Long

Public Enum enHotkeyModifiers
    MOD_ALT = &H1
    MOD_CONTROL = &H2
    MOD_SHIFT = &H4
    MOD_WIN = &H8
End Enum

Private mModifier As Long
Private mvKey As Long
Private mHwnd As Long

Private mhAtom As Long

Private mhKey As Long

Private Declare Function GlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" (ByVal lpString As String) As Integer
Private Declare Function GlobalDeleteAtom Lib "kernel32" (ByVal nAtom As Integer) As Integer
Private Declare Function GlobalGetAtomName Lib "kernel32" Alias "GlobalGetAtomNameA" (ByVal nAtom As Integer, ByVal lpBuffer As String, ByVal nSize As Long) As Long


Public Property Let AltKey(ByVal newvalue As Boolean)

If newvalue Then
    mModifier = mModifier Or MOD_ALT
Else
    mModifier = mModifier And (Not MOD_ALT)
End If

End Property

Public Property Get AltKey() As Boolean

    AltKey = (mModifier And MOD_ALT)
    
End Property


Public Property Let ControlKey(ByVal newvalue As Boolean)

If newvalue Then
    mModifier = mModifier Or MOD_CONTROL
Else
    mModifier = mModifier And (Not MOD_CONTROL)
End If

End Property


Public Property Get ControlKey() As Boolean

    ControlKey = (mModifier And MOD_CONTROL)
    
End Property

Public Property Get hAtom() As Long

    hAtom = mhAtom
    
End Property


Public Property Get hKey() As Long

    hKey = mhKey
    
End Property

Public Property Get hwnd() As Long

    hwnd = mHwnd
    
End Property

Public Property Let hwnd(ByVal newHwnd As Long)

    If newHwnd <> mHwnd Then
        mHwnd = newHwnd
    End If
    
End Property

Public Function MatchedKey(ByVal TestKey As ApiHotkey) As Boolean

With TestKey
    If .AltKey = Me.AltKey And .ControlKey = Me.ControlKey And .ShiftKey = Me.ShiftKey And .VKey = Me.VKey And .WinKey = Me.WinKey Then
        MatchedKey = True
    End If
End With

End Function

Public Property Get Modifier() As Long

    Modifier = mModifier
    
End Property

Public Property Let Modifier(ByVal newval As Long)

    mModifier = newval
    
End Property

Public Property Get UniqueKey() As String

UniqueKey = "hwnd:" & mHwnd & "VK:" & Me.VKey & "Mod:" & Me.Modifier

End Property
Public Sub ReadFromAtom(ByVal hAtom As Long)

Dim sAtom As String
Dim lRet As Long

sAtom = String$(255, 0)

lRet = GlobalGetAtomName(hAtom, sAtom, Len(sAtom))

If lRet > 0 Then
    sAtom = Left$(sAtom, lRet)
End If

If InStr(sAtom, "WIN+") Then
    Me.WinKey = True
End If
If InStr(sAtom, "SHIFT+") Then
    Me.ShiftKey = True
End If
If InStr(sAtom, "CONTROL+") Then
    Me.ControlKey = True
End If
If InStr(sAtom, "ALT+") Then
    Me.AltKey = True
End If

If IsNumeric("0" & Right$(sAtom, 3)) Then
    Me.VKey = Val("0" & Right$(sAtom, 3))
End If

End Sub

Public Sub Register()

Dim lRet As Long

    If mhAtom <> 0 Then
        GlobalDeleteAtom mhAtom
    End If
    mhAtom = GlobalAddAtom(UniqueAtomName())
    
    If mhAtom <> 0 Then
        Debug.Print "Registered unique atom: " & mhAtom
    Else
        Debug.Print "Failed to register unique atom"
    End If
    
    On Error Resume Next
    mhKey = RegisterHotKey(mHwnd, mhAtom, mModifier, mvKey)
    Debug.Print mhKey
    If Err.LastDllError > 0 Then
        Call colControls.Item(Me.UniqueKey).RaiseError(LastSystemError(), "ApiHotkey:Register")
    End If
    
End Sub

Public Property Let ShiftKey(ByVal newvalue As Boolean)

If newvalue Then
    mModifier = mModifier Or MOD_SHIFT
Else
    mModifier = mModifier And (Not MOD_SHIFT)
End If

End Property

Public Property Get ShiftKey() As Boolean

    ShiftKey = (mModifier And MOD_SHIFT)
    
End Property


'\\ --[UniqueAtomName]-------------------------------------------------
'\\ Returns a string which uniquely identifies this hotkey combination
'\\ for use in the .id member when regitsering a new hotkey
'\\ -------------------------------------------------------------------
Public Function UniqueAtomName() As String

Dim sAtom As String

If (mModifier And MOD_WIN) Then
    sAtom = sAtom & "WIN+"
End If

If (mModifier And MOD_SHIFT) Then
    sAtom = sAtom & "SHIFT+"
End If

If (mModifier And MOD_CONTROL) Then
    sAtom = sAtom & "CONTROL+"
End If

If (mModifier And MOD_ALT) Then
    sAtom = sAtom & "ALT+"
End If

sAtom = sAtom & Format$(VKey, "000")
UniqueAtomName = sAtom

End Function

Public Sub Unregister()

If mhAtom <> 0 Then
    Call UnregisterHotKey(mHwnd, mhAtom)
    GlobalDeleteAtom mhAtom
End If

End Sub

Public Property Let VKey(ByVal newvalue As Long)

    If newvalue <> mvKey Then
        mvKey = newvalue
    End If
    
End Property

Public Property Get VKey() As Long

    VKey = mvKey
    
End Property

Public Property Let WinKey(ByVal newvalue As Boolean)

If newvalue Then
    mModifier = mModifier Or MOD_WIN
Else
    mModifier = mModifier And (Not MOD_WIN)
End If

End Property

Public Property Get WinKey() As Boolean

    WinKey = (mModifier And MOD_WIN)
    
End Property


