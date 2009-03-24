Attribute VB_Name = "modAccounts"
Public accTablePath As String
Public accChain As String
Public Function accInfo(username As String) As String ' Level, Items,
    If Len(accChain) <> 0 Then
    Open accTablePath For Input As #1
        While Not EOF(1)
        Line Input #1, linenow
        If Split(linenow, Chr(2) & Chr(3))(0) = username Then
        accInfo = Split(linenow, Chr(2) & Chr(3))(3) & Chr(2) & Chr(3) & Split(linenow, Chr(2) & Chr(3))(4)
        End If
        Wend
    Close #1
    End If
End Function
Public Function chkAcc(username As String, password As String) As Boolean
chkAcc = False
Open accTablePath For Input As #1
    While Not EOF(1)
    Line Input #1, linenow
    If Split(linenow, Chr(2) & Chr(3))(0) = username And Split(linenow, Chr(2) & Chr(3))(1) = password Then
    chkAcc = True
    End If
    Wend
Close #1
End Function
Public Sub LoadAccs()
Open accTablePath For Input As #1
    While Not EOF(1)
        Line Input #1, linenow
    accChain = linenow & vbCrLf & accChain
    Wend
Close #1
End Sub

Public Function initAccs() As Boolean
On Error GoTo errRead
Open accTablePath For Input As #1
    While Not EOF(1)
        Line Input #1, linenow
    Wend
Close #1
initAccs = True
Exit Function
errRead:
    MsgBox "Account table file or path error, please recheck setting", vbCritical + vbOKOnly, "File I/O Error"
    initAccs = False
End Function
