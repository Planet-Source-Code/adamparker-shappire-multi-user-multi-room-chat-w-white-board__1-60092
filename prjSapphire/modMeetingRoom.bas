Attribute VB_Name = "modMeetingRoom"
'======================
'Project Sapphire
'======================
'
'All the meeting room functions and subs
'

Public Sub UpdateRoom(Index, strRoom, Add As Boolean)
Dim lngCount As Integer

For lngCount = 1 To frmMain.lstAccounts.ListItems.Count
    If frmMain.lstAccounts.ListItems(lngCount) = Index Then
            If Add = True Then
                frmMain.lstAccounts.ListItems(lngCount).ListSubItems(8) = frmMain.lstAccounts.ListItems(lngCount).ListSubItems(8) & " " & strRoom & ","
            Else
                frmMain.lstAccounts.ListItems(lngCount).ListSubItems(8) = Replace(frmMain.lstAccounts.ListItems(lngCount).ListSubItems(8), " " & strRoom & ",", "")
            End If
        Exit Sub
    End If
Next lngCount
End Sub

Public Sub RelayMessage(strRoom, strData)
Dim lngCount As Integer
For lngCount = 1 To frmMain.lstAccounts.ListItems.Count
    If Not InStr(1, frmMain.lstAccounts.ListItems(lngCount).ListSubItems(8), strRoom & ",", vbTextCompare) = 0 Then
        'User is in room, tell them a person entered
        Call SendData(frmMain.lstAccounts.ListItems(lngCount), "404-" & strRoom & "," & strData)
    End If
Next lngCount
End Sub

Public Sub SendDrawing2Room(strRoom, strData)
Dim lngCount As Integer
For lngCount = 1 To frmMain.lstAccounts.ListItems.Count
    If Not InStr(1, frmMain.lstAccounts.ListItems(lngCount).ListSubItems(8), strRoom & ",", vbTextCompare) = 0 Then
        'User is in room, send them the picture
        Call SendData(frmMain.lstAccounts.ListItems(lngCount), "410-" & strRoom & ":" & strData)
    End If
Next lngCount
End Sub

Public Function PasswordRoom(strRoom, strPassword, Index) As Boolean
On Error Resume Next

Dim lngCount As Integer
'Go through all the user list to see who's in that room
For lngCount = 1 To frmMain.lstRooms.ListItems.Count
    If frmMain.lstRooms.ListItems(lngCount) = strRoom Then
        If frmMain.lstRooms.ListItems(lngCount).ListSubItems(2) = strPassword Then
            PasswordRoom = True
            Call RoomCount(strRoom, True)
        Else
            If frmMain.lstRooms.ListItems(lngCount).ListSubItems(2) = "No" Then
                PasswordRoom = True
                Call RoomCount(strRoom, True)
            Else
                PasswordRoom = False
            End If
        End If
        Exit Function
    End If
Next lngCount

'Guess what, room wasn't in list, that means we get to add it :)
lngCount = frmMain.lstRooms.ListItems.Count + 1

Call frmMain.lstRooms.ListItems.Add(lngCount, , strRoom)
Call frmMain.lstRooms.ListItems(lngCount).ListSubItems.Add(1, , "User Created")
Call frmMain.lstRooms.ListItems(lngCount).ListSubItems.Add(2, , strPassword)
Call frmMain.lstRooms.ListItems(lngCount).ListSubItems.Add(3, , "No")
Call frmMain.lstRooms.ListItems(lngCount).ListSubItems.Add(4, , "0")
Call frmMain.lstRooms.ListItems(lngCount).ListSubItems.Add(5, , UserFromIndex(Index))

PasswordRoom = True
Call RoomCount(strRoom, True)
End Function

Public Function GetRoomList(strRoom) As String
Dim lngCount As Integer, strTmp As String
'Go through all the user list to see who's in that room
For lngCount = 1 To frmMain.lstAccounts.ListItems.Count
    If Not InStr(1, frmMain.lstAccounts.ListItems(lngCount).ListSubItems(8), strRoom & ",", vbTextCompare) = 0 Then
        'User is in room, add them to string
        strTmp = strTmp & frmMain.lstAccounts.ListItems(lngCount).ListSubItems(1) & ","
    End If
Next lngCount
    GetRoomList = strTmp
End Function

Public Sub UserEnters(strRoom, Index)
Dim lngCount As Integer
For lngCount = 1 To frmMain.lstAccounts.ListItems.Count
    If Not InStr(1, frmMain.lstAccounts.ListItems(lngCount).ListSubItems(8), strRoom & ",", vbTextCompare) = 0 Then
        'User is in room, tell them a person entered
        Call SendData(frmMain.lstAccounts.ListItems(lngCount), "402-" & strRoom & "," & UserFromIndex(Index))
    End If
Next lngCount
End Sub

Public Sub UserExits(strRoom, strUser)
Dim lngCount As Integer
For lngCount = 1 To frmMain.lstAccounts.ListItems.Count
    If Not InStr(1, frmMain.lstAccounts.ListItems(lngCount).ListSubItems(8), strRoom & ",", vbTextCompare) = 0 Then
        'User is in room, tell them a person entered
        Call SendData(frmMain.lstAccounts.ListItems(lngCount), "403-" & strRoom & "," & strUser)
        Call RoomCount(strRoom, False)
    End If
Next lngCount
End Sub

Public Sub Send2Room(strRoom, strPrefix, strData)
Dim lngCount As Integer
For lngCount = 1 To frmMain.lstAccounts.ListItems.Count
    If Not InStr(1, frmMain.lstAccounts.ListItems(lngCount).ListSubItems(8), strRoom & ",", vbTextCompare) = 0 Then
        'User is in room, tell them a person entered
        Call SendData(frmMain.lstAccounts.ListItems(lngCount), strPrefix & "-" & strData)
    End If
Next lngCount
End Sub


Public Sub RoomCount(strRoom, Add2Total As Boolean)
On Error Resume Next

Dim lngCount As Integer, lngTmp

For lngCount = 1 To frmMain.lstRooms.ListItems.Count
    If frmMain.lstRooms.ListItems(lngCount) = strRoom Then
        'Update the user count
        lngTmp = frmMain.lstRooms.ListItems(lngCount).ListSubItems(4)
        
        If Add2Total = True Then
            frmMain.lstRooms.ListItems(lngCount).ListSubItems(4) = lngTmp + 1
        Else
            frmMain.lstRooms.ListItems(lngCount).ListSubItems(4) = lngTmp - 1
        End If
        
        If frmMain.lstRooms.ListItems(lngCount).ListSubItems(4) < 1 And Not frmMain.lstRooms.ListItems(lngCount).ListSubItems(5) = "Admin" Then
            frmMain.lstRooms.ListItems.Remove (lngCount)
        End If
    End If
Next lngCount
End Sub

Public Sub NotifySignOff(strRoomList, strUser)
Dim Pos1 As Long, Pos2 As Long, strTmp As String
Do
    Pos1 = Pos2 + Len(",")
    Pos2 = InStr(Pos1, strData, ",")
    If Pos2 <> 0 Then
        strTmp = Mid(strRoomList, Pos1, Pos2 - Pos1)
        strTmp = Replace(strTmp, " ", "")
        Call UserExits(strTmp, strUser)
    End If
Loop Until Pos2 = 0
End Sub

Public Function UserFromIndex(Index) As String
Dim lngCount As Integer
For lngCount = 1 To frmMain.lstAccounts.ListItems.Count
    If frmMain.lstAccounts.ListItems(lngCount) = Index Then
        UserFromIndex = frmMain.lstAccounts.ListItems(lngCount).ListSubItems(1)
    End If
Next lngCount
End Function

Public Function InRoom(strRoom, Index) As Boolean
Dim lngCount As Integer
'Go through all the user list to see who's in that room
For lngCount = 1 To frmMain.lstAccounts.ListItems.Count
    If frmMain.lstAccounts.ListItems(lngCount) = Index Then
        If Not InStr(1, frmMain.lstAccounts.ListItems(lngCount).ListSubItems(8), strRoom & ",", vbTextCompare) = 0 Then
            InRoom = True
        End If
    End If
Next lngCount
End Function

Public Function GetDescription(strRoom) As String
Dim lngCount As Integer
'Go through all the user list to see who's in that room
For lngCount = 1 To frmMain.lstRooms.ListItems.Count
    If frmMain.lstRooms.ListItems(lngCount) = strRoom Then
        GetDescription = frmMain.lstRooms.ListItems(lngCount).ListSubItems(1)
        Exit Function
    End If
Next lngCount
End Function

Public Sub ChangePassword(Index, strRoom, strDescription)
Dim strPowerUser As String, lngCount As Integer
strPowerUser = UserFromIndex(Index)

For lngCount = 1 To frmMain.lstRooms.ListItems.Count
    If frmMain.lstRooms.ListItems(lngCount) = strRoom Then
        If frmMain.lstRooms.ListItems(lngCount).ListSubItems(5) = strPowerUser Then
            frmMain.lstRooms.ListItems(lngCount).ListSubItems(2) = strDescription
            Call Send2Room(strRoom, "404", strRoom & ",Online Host,Room Password Changed")
        Else
            Call SendData(Index, "500-Your request to change the password was denied.")
        End If
    End If
Next lngCount

End Sub

Public Sub ChangeDescription(Index, strRoom, strDescription)
Dim strPowerUser As String, lngCount As Integer
strPowerUser = UserFromIndex(Index)

For lngCount = 1 To frmMain.lstRooms.ListItems.Count
    If frmMain.lstRooms.ListItems(lngCount) = strRoom Then
        If frmMain.lstRooms.ListItems(lngCount).ListSubItems(5) = strPowerUser Then
            frmMain.lstRooms.ListItems(lngCount).ListSubItems(1) = strDescription
            Call Send2Room(strRoom, "408", strRoom & ",Topic Changed," & strDescription)
        Else
            Call SendData(Index, "500-Your request to change the topic was denied.")
        End If
    End If
Next lngCount

End Sub

Public Sub RejectUser(Index, strRoom, strUser)
Dim strPowerUser As String, lngCount As Integer
strPowerUser = UserFromIndex(Index)

For lngCount = 1 To frmMain.lstRooms.ListItems.Count
    If frmMain.lstRooms.ListItems(lngCount) = strRoom Then
        If frmMain.lstRooms.ListItems(lngCount).ListSubItems(5) = strPowerUser Then
            '
            'Kicking procedure
            '
        Else
            Call SendData(Index, "500-Your request to boot " & strUser & " was denied.")
        End If
    End If
Next lngCount

End Sub
