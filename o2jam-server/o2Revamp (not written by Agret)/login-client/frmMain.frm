VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWinsck.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "O2JAM Client Login"
   ClientHeight    =   2475
   ClientLeft      =   6915
   ClientTop       =   1605
   ClientWidth     =   4845
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   2  'Cross
   ScaleHeight     =   2475
   ScaleWidth      =   4845
   Begin VB.Timer tmr_connect 
      Enabled         =   0   'False
      Left            =   3720
      Top             =   360
   End
   Begin VB.CommandButton cmdServers 
      Caption         =   "Servers"
      Height          =   255
      Left            =   1440
      TabIndex        =   8
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton helpset 
      Caption         =   "Help"
      Height          =   255
      Left            =   960
      TabIndex        =   6
      Top             =   2040
      Width           =   495
   End
   Begin MSWinsockLib.Winsock winsckConnect 
      Left            =   3360
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSet 
      Caption         =   "Configure"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000C0&
      Caption         =   "Lets Play!"
      Height          =   495
      Left            =   120
      MaskColor       =   &H8000000F&
      TabIndex        =   3
      Top             =   1440
      Width           =   4455
   End
   Begin VB.TextBox txtPas 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1080
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox txtUser 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Using server"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Copyright 2006 @ O2Emu Project"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2400
      TabIndex        =   7
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   1125
      Left            =   3240
      Picture         =   "frmMain.frx":0CCA
      Top             =   360
      Width           =   1500
   End
   Begin VB.Label lbl3 
      BackColor       =   &H00000000&
      Caption         =   "Password"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   855
   End
   Begin VB.Label lbl2 
      BackColor       =   &H00000000&
      Caption         =   "Username"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdServers_Click()
frmServers.Show
End Sub


Private Sub Form_Unload(Cancel As Integer)
savesettings
End
End Sub

Private Sub helpset_Click()
frmHelp.Show
End Sub

Private Sub cmdSet_Click()
frmSet.Show
frmMain.Hide
End Sub

Private Sub Command1_Click()
If frmMain.winsckConnect.State <> 7 Then
execRun
End If
End Sub

Function RegQueryStringValue(ByVal hKey As Long, ByVal strValueName As String) As String
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


Private Sub Image1_Click()
txtPas.SetFocus
End Sub

Private Sub tmr_connect_Timer()
frmLoad.bar1 = frmLoad.bar1 + 1
If winsckConnect.State = sckConnected Then
    If frmLoad.bar1 >= 99 Then
    Dim x As Integer
    frmLoad.bar1 = frmLoad.bar1 - 1
        x = MsgBox("System server does not response in time, retry connection?", vbYesNo + vbCritical, "Timeout Error")
        If x = 6 Then
            frmLoad.bar1.Value = 0
            conState = 0
        ElseIf x = 7 Then
            loadState False
            tmr_connect.Enabled = False
            Exit Sub
        End If
    End If
    Select Case conState
        Case Is = 0:
            winsckConnect.SendData "0001" & txtUser & Chr(1) & txtPas
            frmLoad.bar1.Value = 40
            frmLoad.Label1 = "Authenticating........"
            conState = 5
        Case Is = 3:
            
            tmr_connect.Interval = 0
            tmr_connect.Enabled = False
        End Select
    
ElseIf winsckConnect.State = sckError Then
        x = MsgBox("Error connecting to server, Retry?", vbRetryCancel, "Connection error")
        conState = 0
        If x = 2 Then
        loadState False
        tmr_connect.Enabled = False
            MsgBox "Connection error", vbOKOnly + vbCritical, "Error"
        ElseIf x = 4 Then
            frmLoad.bar1 = 0
        End If
End If
End Sub

Private Sub txtPas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Call Command1_Click
End If
End Sub

Private Sub winsckConnect_DataArrival(ByVal bytesTotal As Long)
Dim pcktCmd As String
Dim pcktDat As String
Dim pckt As String
winsckConnect.GetData pckt
pcktCmd = Left(pckt, 4)
pcktDat = Right(pckt, Len(pckt) - 4)
Select Case pcktCmd
    
    Case Is = "1101"
        If Mid(pckt, 5, 1) = "1" Then
            MsgBox "Authentication failure, Username or password error", vbOKOnly + vbCritical, "Authentication Failure"
            conState = 0
            loadState False
        ElseIf Mid(pckt, 5, 1) = "0" Then
            conState = 3
            Dim gpath As String
            gpath = Split(pckt, Chr(3))(1)
            RunGame gpath
        ElseIf Mid(pckt, 5, 1) = "2" Then
            MsgBox "Login server not available, please try again later", vbOKOnly + vbInformation, "Login server error"
            conState = 0
            loadState False
        End If
    Case Is = "1111"
        Dim serverMsg As String
        serverMsg = Split(pckt, Chr(3))(1)
        MsgBox serverMsg, vbInformation, "Server message"
        
    
End Select
End Sub

