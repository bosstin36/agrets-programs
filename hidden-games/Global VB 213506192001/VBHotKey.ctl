VERSION 5.00
Begin VB.UserControl VBHotKey 
   ClientHeight    =   675
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   690
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   675
   ScaleWidth      =   690
   ToolboxBitmap   =   "VBHotKey.ctx":0000
End
Attribute VB_Name = "VBHotKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Default Property Values:
Const m_def_AltKey = False
Const m_def_ShiftKey = False
Const m_def_CtrlKey = False
Const m_def_WinKey = True
Const m_def_VKey = vbKeyA

'Property Variables:
Dim m_AltKey As Boolean
Dim m_ShiftKey As Boolean
Dim m_CtrlKey As Boolean
Dim m_VKey As Long
Dim m_WinKey As Boolean

Dim mHwnd As Long
Private mHotkey As ApiHotkey

'Event Declarations:
Event HotkeyPressed()
Event ErrorOccured(ByVal Message As String, ByVal source As String)



Public Property Get hAtom() As Long

    hAtom = mHotkey.hAtom
    
End Property

Public Property Get hKey() As Long

    hKey = mHotkey.hKey
    
End Property


Friend Sub RaiseError(ByVal sMessage As String, ByVal sSource As String)

RaiseEvent ErrorOccured(sMessage, sSource)

End Sub

Friend Sub RaiseKeyPressEvent(ByVal VKey As Long, ByVal Modifier As Long)

Dim keyThis As ApiHotkey

Set keyThis = New ApiHotkey

With keyThis
    .VKey = VKey
    .Modifier = Modifier
    If .MatchedKey(mHotkey) Then
        RaiseEvent HotkeyPressed
    End If
End With

Set keyThis = Nothing

End Sub

Public Sub StopHotkey()

If mHwnd <> 0 Then
    Call FreeSubclassedWindow(mHwnd)
    If Not (mHotkey Is Nothing) Then
        mHotkey.Unregister
        Set mHotkey = Nothing
    End If
End If
    
End Sub

Public Property Get UniqueKey() As String

UniqueKey = mHotkey.UniqueKey

End Property


Private Sub UserControl_Initialize()

Set mHotkey = New ApiHotkey

End Sub
Public Property Get AltKey() As Boolean
    AltKey = m_AltKey
End Property

Public Property Let AltKey(ByVal New_AltKey As Boolean)
    If Ambient.UserMode Then Err.Raise 393
    m_AltKey = New_AltKey
    PropertyChanged "AltKey"
End Property

Public Property Get ShiftKey() As Boolean
    ShiftKey = m_ShiftKey
End Property

Public Property Let ShiftKey(ByVal New_ShiftKey As Boolean)
    If Ambient.UserMode Then Err.Raise 393
    m_ShiftKey = New_ShiftKey
    PropertyChanged "ShiftKey"
End Property

Public Property Get CtrlKey() As Boolean
    CtrlKey = m_CtrlKey
End Property

Public Property Get WinKey() As Boolean
    WinKey = m_WinKey
End Property
Public Property Let CtrlKey(ByVal New_CtrlKey As Boolean)
    If Ambient.UserMode Then Err.Raise 393
    m_CtrlKey = New_CtrlKey
    PropertyChanged "CtrlKey"
End Property

Public Property Let WinKey(ByVal New_WinKey As Boolean)
    If Ambient.UserMode Then Err.Raise 393
    m_WinKey = New_WinKey
    PropertyChanged "WinKey"
End Property
Public Property Get VKey() As KeyCodeConstants
    VKey = m_VKey
End Property

Public Property Let VKey(ByVal New_VKey As KeyCodeConstants)
    If Ambient.UserMode Then Err.Raise 393
    m_VKey = New_VKey
    PropertyChanged "VKey"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()


    m_AltKey = m_def_AltKey
    m_ShiftKey = m_def_ShiftKey
    m_CtrlKey = m_def_CtrlKey
    m_VKey = m_def_VKey
    m_WinKey = m_def_WinKey
    

    
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_AltKey = PropBag.ReadProperty("AltKey", m_def_AltKey)
    m_ShiftKey = PropBag.ReadProperty("ShiftKey", m_def_ShiftKey)
    m_CtrlKey = PropBag.ReadProperty("CtrlKey", m_def_CtrlKey)
    m_VKey = PropBag.ReadProperty("VKey", m_def_VKey)
    m_WinKey = PropBag.ReadProperty("WinKey", m_def_WinKey)
    
    
    '\\ If we in RUN mode
    If UserControl.Ambient.UserMode Then
        
        With UserControl.Parent
            If .hwnd > 0 Then
                '\\ Register the hotkey
                mHwnd = .hwnd
                With mHotkey
                    .AltKey = m_AltKey
                    .ShiftKey = m_ShiftKey
                    .ControlKey = m_CtrlKey
                    .WinKey = m_WinKey
                    .VKey = m_VKey
                    .hwnd = mHwnd
                    .Register
                End With
                '\\ And subclass the parent to listen to it...
                Call SubclassWindow(mHwnd)
                '\\ DEJ 18-June-2001 Keys must be unique per window only...
                colControls.Add Me, mHotkey.UniqueKey
            End If
        End With
    End If
End Sub

Private Sub UserControl_Terminate()

Call Me.StopHotkey

End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("AltKey", m_AltKey, m_def_AltKey)
    Call PropBag.WriteProperty("ShiftKey", m_ShiftKey, m_def_ShiftKey)
    Call PropBag.WriteProperty("CtrlKey", m_CtrlKey, m_def_CtrlKey)
    Call PropBag.WriteProperty("VKey", m_VKey, m_def_VKey)
    Call PropBag.WriteProperty("WinKey", m_WinKey, m_def_WinKey)
    
End Sub

