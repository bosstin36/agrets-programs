Attribute VB_Name = "modCrypt"
Public Function decode(key As String) As Variant
Dim lkey As Integer
Dim SKey As String
Dim a, b, c, d As String
lkey = Len(key) / 8
For i0 = 1 To lkey ' split key into 8bit
a = i0 * 8
c = Mid(key, a - 7, 8)
    For i1 = 1 To 8
    Dim codeCHR As String
    codeCHR = Asc(Mid(c, i1, 1))
    On Error GoTo 1
    If codeCHR > 45 And codeCHR < 57 Then ' read the 8bit key and add to collection
    codeCHR = codeCHR + 4
        ElseIf 65 <= codeCHR And codeCHR < 91 Then
        codeCHR = codeCHR - 65
        ElseIf codeCHR > 96 And codeCHR <= 122 Then
        codeCHR = codeCHR - 71
        End If
    b = b & Get6Binary(codeCHR)
   Next i1
Next i0
    Dim key16 As String
    Debug.Print b
For i2 = 1 To Len(b) / 16 'decode the 16bit binary to character
    key16 = Mid(b, (i2 * 16) - 15, 16)
    d = d & Chr(bin2dec(key16))
Next i2

decode = d
Exit Function
1:
MsgBox "The length of key must be 8bit, length of key must be dividible by 8", vbOKOnly, "Decode error"
End Function

Function Get6Binary(ByVal inHex As String) As String
inHex = Hex(inHex)
    Dim mDec As Integer
    Dim s As String
    Dim i
    mDec = CInt("&h" & inHex)
    s = Trim(CStr(mDec Mod 2))
    i = mDec \ 2
    Do While i <> 0
        s = Trim(CStr(i Mod 2)) & s
        i = i \ 2
    Loop
    Do While Len(s) < 6
        s = "0" & s
    Loop
    Get6Binary = s
End Function

Function bin2dec(binary As String) As Integer
Dim factor As Integer
Dim bindec As Integer
Dim dec As Integer
dec = 0

If Left(binary, 1) = 0 Then
For i = 1 To Len(binary)
If Left(binary, 1) = 0 Then
binary = Right(binary, Len(binary) - 1)
ElseIf Left(binary, 1) = 1 Then
Exit For
End If
Next i
End If

binary = StrReverse(binary)
For i = 1 To Len(binary)
factor = 2 ^ (i - 1)
bindec = Mid(binary, i, 1) * factor
dec = bindec + dec
Next i
bin2dec = dec

End Function

Public Function encode(key As String) As String
Dim lkey As String
Dim a, d As Integer
Dim i, i1, i2 As Integer
lkey = Len(key)
For i = 1 To lkey
a = a & Get16Binary(Asc(Mid(key, i, 1))) ' turn 8bit binary into 16bit
Next i
Dim b As String
Dim c As String
For i1 = 1 To Len(a) / 6
b = bin2dec(Mid(a, i1 * 6 - 5, 6))
'MsgBox b
If b >= 0 And b <= 26 Then
b = b + 65
ElseIf b >= 21 And b <= 51 Then
b = b + 71
ElseIf b >= 50 And b <= 60 Then
b = b - 4
End If
c = c & Chr(b)
Next i1
encode = c
End Function
Function Get16Binary(ByVal inHex As String) As String
inHex = Hex(inHex)
    Dim mDec As Integer
    Dim s As String
    Dim i
    mDec = CInt("&h" & inHex)
    s = Trim(CStr(mDec Mod 2))
    i = mDec \ 2
    Do While i <> 0
        s = Trim(CStr(i Mod 2)) & s
        i = i \ 2
    Loop
    Do While Len(s) < 16
        s = "0" & s
    Loop
    Get16Binary = s
End Function

