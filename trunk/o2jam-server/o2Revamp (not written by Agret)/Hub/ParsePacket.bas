Attribute VB_Name = "ParsePacket"

Public Sub MainParse(packet As String, Index As Integer)
Dim cmdPack As String
cmdPack = Chr2Hex(Mid(packet, 3, 2))
Select Case cmdPack
Case Is = "E8 03": 'login packet, reply E9 03
    If players.joinUser(Index, Mid(packet, 5, Len(packet) - 5)) Then
        parse_sendDate Index
    End If
Case Is = "EA 03":
    If players(Index).authED Then
    parse_SendChannel (Index)
    End If
End Select
End Sub

Public Sub parse_sendDate(Index As Integer)
Dim getmonth As String, getday As String
getmonth = Month(Date)
    If Len(getmonth) = 1 Then getmonth = "0" & getmonth
getday = Day(Date)
    If Len(getday) = 1 Then getday = "0" & getday
sendToClient Index, Hex2Chr("E9 03 00 00 00 00 44 42") & Year(Date) & getmonth & Hex2Chr("3C 00 00 00")
End Sub

Public Sub parse_SendChannel(Index As Integer)
Dim i As Integer
Dim s As String
Dim b As String

For i = 0 To Channels.Count
    s = chr(i) & chr(0) & chr(78) & makechar(chr(0), 3) & Channels(i + 1).Rooms.Count & makechar(chr(0), 3) & Channels(i + 1).enabled & makechar(chr(0), 2)
Next i

s = StrReverse(chr(i)) & makechar(chr(0), 2) & s
sendToClient Index, Hex2Chr("EB 03") & s
End Sub

Public Sub sendToClient(Index As Integer, packet As String)
Dim i1 As String
Dim i2 As String
Dim s As String

s = Chr2Hex(Len(packet))

Select Case Len(s)
Case Is = 1
    s = "000" & s
Case Is = 2
    s = "00" & s
Case Is = 3
    s = "0" & s
End Select

i1 = Left(s, 2)
i2 = Right(s, 2)

frmMain.sckControl(Index).SendData i1 & i2 & packet

End Sub

Public Function getRoomList(Index As Integer) As String
Dim a As String
Dim b As String

End Function
