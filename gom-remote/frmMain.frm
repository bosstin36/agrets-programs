VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GOMRemote"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long

Private Const SW_SHOWNORMAL = 1
Private Const SW_SHOWMINIMIZED = 2
Private Const SW_SHOWMAXIMIZED = 3

Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

' Constant declarations:
Const VK_NUMLOCK = &H90
Const VK_SCROLL = &H91
Const VK_CAPITAL = &H14
Const VK_CONTROL = &H11
Const VK_ENTER = 13
Const KEYEVENTF_EXTENDEDKEY = &H1
Const KEYEVENTF_KEYUP = &H2
 Public Sub Wait(ByVal Delay As Integer)
        Dim time As Long
        time = (Timer * 60) + Delay
        Do While time > (Timer * 60)
            DoEvents
        Loop
End Sub

Private Sub Form_Load()
    Dim StartupCommands() As String
    Dim GoFullScreen As Boolean
    Dim Maximized As Boolean
    Dim ExitAfterMovie As Boolean
    Dim ExitAfterPlaylist As Boolean
    StartupCommands = Split(Command, " ")
    
    Dim i As Integer
    For i = 0 To UBound(StartupCommands)
        If LCase(StartupCommands(i)) = "-fullscreen" Then
            GoFullScreen = True
        ElseIf LCase(StartupCommands(i)) = "-maximize" Then
            Maximized = True
        ElseIf LCase(StartupCommands(i)) = "-exitaftermovie" Then
            ExitAfterMovie = True
        ElseIf LCase(StartupCommands(i)) = "-exitafterplaylist" Then
            ExitAfterPlaylist = True
        Else
            MsgBox "Unknown Command Line Option:" & vbNewLine & StartupCommands(i), vbCritical, "Error"
        End If
    Next i
    
    If UBound(StartupCommands) = -1 Then
        MsgBox "GOMRemote v" & App.Major & "." & App.Minor & vbNewLine & vbNewLine & "Command Line Options:" & vbNewLine & "-fullscreen : Makes GOMPlayer fullscreen." & vbNewLine & "-maximize : Maximizes the GOMPlayer window." & vbNewLine & "-exitaftermovie : Exits GOMPlayer after the current movie ends." & vbNewLine & "-exitafterplaylist : Exits GOMPlayer after the current playlist ends." & vbNewLine & vbNewLine & "Created by Agret, alias.zero0297@gmail.com", vbInformation, "GOMRemote v" & App.Major & "." & App.Minor
        End
    End If
    
    If App.PrevInstance = True Then End ' Don't need this looping to hell :P Just don't use different command line and you'll be alright ;)
    
    Dim lHwnd As Long
    Wait 20 ' Since we start this from the batch we want a delay to allow GOM time to open.
    
    Do While lHwnd = 0
        lHwnd = FindWindow("GomPlayer1.x", vbNullString)
 
        If lHwnd <> 0 Then 'GOMPlayer Window Found :)
            Wait 30 'Give it a short time to initialize in case it's just starting
            BringWindowToTop lHwnd
            If Maximized = True Then
                ShowWindow lHwnd, SW_SHOWMAXIMIZED 'Set focus and maximize
            Else
                ShowWindow lHwnd, SW_SHOWNORMAL 'Set focus without maximizing
            End If
            Wait 200
            keybd_event VK_CONTROL, 0, KEYEVENTF_EXTENDEDKEY Or 0, 0
            If ExitAfterMovie = True Or ExitAfterPlaylist = True Then
                keybd_event Asc("X"), 0, 0, 0 ' Quit on movie end
                keybd_event Asc("X"), 0, KEYEVENTF_KEYUP, 0
            End If
            If ExitAfterPlaylist = True Then
                keybd_event Asc("X"), 0, 0, 0 ' Quit on playlist end (2x CTRL+X)
                keybd_event Asc("X"), 0, KEYEVENTF_KEYUP, 0
            End If
            If GoFullScreen = True Then
                Wait 10
                keybd_event VK_ENTER, 0, 0, 0 ' Fullscreen
                keybd_event VK_ENTER, 0, KEYEVENTF_KEYUP, 0
            End If
            keybd_event VK_CONTROL, 0, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0
        End If
               
        Wait 200
    Loop
    End
End Sub
