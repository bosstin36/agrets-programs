VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Any App as a Service v"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   7710
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraServices 
      Caption         =   "Services"
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7455
      Begin MSComDlg.CommonDialog CommonDialog 
         Left            =   2040
         Top             =   3480
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   495
         Left            =   480
         TabIndex        =   3
         Top             =   5040
         Width           =   1815
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove"
         Height          =   495
         Left            =   5040
         TabIndex        =   2
         Top             =   5040
         Width           =   1815
      End
      Begin MSComctlLib.ListView lstServices 
         Height          =   4695
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   8281
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
    Dim ServiceName As String
    Dim ServicePath As String
    Dim ServiceDescription As String
    
    ServiceName = InputBox("Service Name:", "Add Service")
    If Len(ServiceName) > 0 Then
        CommonDialog.Filter = "*.exe|*.exe"
        CommonDialog.FileName = ""
        CommonDialog.ShowOpen
        ServicePath = CommonDialog.FileName
        If Len(ServicePath) > 0 Then
            Open App.Path & "\tmp.bat" For Output As #1
                Print #1, "@ECHO OFF"
                Print #1, "INSTSRV.EXE " & """" & ServiceName & """" & " " & """" & "REMOVE" & """"
                Print #1, "INSTSRV.EXE " & """" & ServiceName & """" & " " & """" & ServicePath & """"
                ServiceDescription = ""
                ServiceDescription = InputBox("Service Description:", "Add Service", "User service created by Service Manager")
                Print #1, "REG ADD " & """" & "HKLM\SYSTEM\CurrentControlSet\Services\" & ServiceName & """" & " /v " & """" & "Application" & """" & " /t REG_SZ /d " & """" & ServicePath & """"
                If Len(ServiceDescription) > 0 Then
                    Print #1, "REG ADD " & """" & "HKLM\SYSTEM\CurrentControlSet\Services\" & ServiceName & """" & " /v " & """" & "Description" & """" & " /t REG_SZ /d " & """" & ServiceDescription & """"
                End If
            Close #1
        End If
    End If
End Sub

Private Sub cmdRemove_Click()
    If lstServers.ListIndex > -1 Then
        lstAccounts.SelectedItem.Text & " " & lstAccounts.SelectedItem.SubItems(1)
        lstServers.RemoveItem lstServers.ListIndex
    End If
End Sub

Private Sub Form_Load()
    Me.Caption = "Any App as a Service v" & App.Major & "." & App.Minor & "." & App.Revision
    lstServices.ColumnHeaders.Add , "Service Name", "Service Name"
    lstServices.ColumnHeaders.Add , "Service Path", "Service Path"
    lstServices.ColumnHeaders(1).Width = lstServices.Width / 2 - 70
    lstServices.ColumnHeaders(2).Width = lstServices.Width / 2 - 70
    LoadServiceList App.Path & "\services.txt"
End Sub

Private Sub SaveServiceList(slocation As String)
    Open App.Path & "\services.txt" For Append As #1
        Print #1, ServiceName & "§" & ServicePath
    Close #1
End Sub

Private Sub LoadServiceList(slocation As String)
    lstServices.ListItems.Clear
    
    On Error GoTo dlgerror
    Dim sCurrent As String
    Dim ServicesList() As String
    Dim i, Filenum As Integer

    Filenum = FreeFile
    Open slocation For Input As #Filenum
    i = 0
    Do Until EOF(1)
    Line Input #Filenum, sCurrent
    ServicesList() = Split(sCurrent, "§")
    
    If UBound(ServicesList) = 1 Then
        Set lItem = lstServices.ListItems.Add(, , ServicesList(0))
        lItem.ListSubItems.Add , , ServicesList(1)
    End If
    
    i = i + 1
    Loop
    Close #Filenum
    Exit Sub
dlgerror:
    If Err.Number <> 53 Then
        MsgBox "An error has occured: " & Err.Description
    End If
    Exit Sub
End Sub
