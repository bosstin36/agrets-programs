Attribute VB_Name = "modCals"
Public Function HexToDec(ByVal HexStr As String) As Double
Dim mult As Double
Dim DecNum As Double
Dim ch As String
Dim i As Integer
mult = 1
DecNum = 0
For i = Len(HexStr) To 1 Step -1
    ch = Mid(HexStr, i, 1)
    If (ch >= "0") And (ch <= "9") Then
        DecNum = DecNum + (Val(ch) * mult)
    Else
        If (ch >= "A") And (ch <= "F") Then
            DecNum = DecNum + ((AscW(ch) - AscW("A") + 10) * mult)
        Else
            If (ch >= "a") And (ch <= "f") Then
                DecNum = DecNum + ((AscW(ch) - AscW("a") + 10) * mult)
            Else
                HexToDec = 0
                Exit Function
            End If
        End If
    End If
    mult = mult * 16
Next i
HexToDec = DecNum
End Function

Public Function Hex2Chr(ByVal HexStr As String)
Dim sArr() As String
Dim sArr2() As String
Dim i As Long
Dim X As Long
Dim cal1  As String
Dim cal2 As String
sArr() = Split(LTrim(HexStr), " ")
For i = 0 To UBound(sArr())
cal1 = cal1 & " " & HexToDec(sArr(i))
Next i

sArr2() = Split(LTrim(cal1), " ")
For X = 0 To UBound(sArr2())
cal2 = cal2 & ChrW(sArr2(X))
Next X
'If Len(cal2) = 1 Then cal2 = "0" & cal2
Hex2Chr = cal2
End Function
Public Function Chr2Hex(ByVal chr As String) As Variant
Dim i As Integer
Dim b As Integer
Dim c As Integer
Dim varR As String
Dim d As Variant
Do Until i = Len(chr)
c = i + 1
b = i - 2
If b <= 0 Then b = 1
    d = hex$(AscB(Mid(chr, c, 1)))
    If Len(d) = 1 Then d = "0" & d

varR = varR & " " & d
i = i + 1
Loop
Chr2Hex = LTrim(varR)
End Function
Public Function Hex2dec(ByVal hex As String)
Dim HexBhex As Variant
Dim sArr() As String
Dim i As Long
sArr() = Split(hex, " ")
For i = 0 To UBound(sArr())
HexBhex = HexBhex & " " & HexToDec(sArr(i))
Next i
Hex2dec = LTrim(HexBhex)
End Function

Public Function makechar(ByVal char As String, ByVal numbers As Integer) As String
For i = 1 To numbers
makechar1 = makechar1 & char
Next i
makechar = makechar1
End Function

