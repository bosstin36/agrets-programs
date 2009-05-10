VERSION 5.00
Object = "{95E3FB4C-5F47-11D5-B235-00C04F84CB14}#1.0#0"; "MCLHotkey.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MCLHotkey.VBHotKey VBHotKey1 
      Left            =   1440
      Top             =   1440
      _ExtentX        =   1296
      _ExtentY        =   873
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Terminate()

Me.VBHotKey1.StopHotkey

End Sub

Private Sub VBHotKey1_HotkeyPressed()

Debug.Print "Hotkey pressed"

End Sub
