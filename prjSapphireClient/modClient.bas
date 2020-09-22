Attribute VB_Name = "modClient"
Option Explicit
'
'Protocol
'
'100- Reserved for normal communication between the server and client
'200- Reserved for internal mail/offline messages system
'300- Reserved for instant messaging
'400- Reserved for meeting rooms
'500- Errors
'
    
Public Function GetLeft(strString, strSeperater) As String
    On Error Resume Next
    GetLeft = Left$(strString, InStr(strString, strSeperater) - 1)
End Function

Public Function GetRight(strString, strSeperater) As String
    On Error Resume Next
    GetRight = Right(strString, Len(strString) - InStr(strString, strSeperater))
End Function

Public Sub ErrorDialog(strTitle, strMessage)
'create a new frmError and display message
Dim errorForm As New frmError
    errorForm.Caption = "Error: " & strTitle
    errorForm.txtError = strMessage
    errorForm.Show
End Sub

Public Sub SendData(strData)
On Error Resume Next
Call frmMain.sckClient.SendData(strData & "Œ")
End Sub

Public Sub SendIM(strSendTo, strUser, strMessage)
Call frmMain.sckClient.SendData("300-" & strSendTo & ":" & strUser & ": " & strMessage & "Œ")
End Sub

Public Sub ClientParse(Data As String)
    Dim strData As String, lngCount As Integer, booFound As Boolean, frm As Form, Pos1, Pos2
    Dim DrawBuff As String, DLine, WData, i, Tmp

    If Mid(Data, 1, 4) = "100-" Then
        '101- Handshake accepted
        strData = Mid(Data, 5)
        frmWelcome.Show
        frmMain.Hide
        RemoveBar frmMain.Name, frmMain.Caption 'Remove to task switch
    End If
    
    If Mid(Data, 1, 4) = "101-" Then
        '101- Handshake denied, display why
        strData = Mid(Data, 5)
        Call ErrorDialog("Login", strData)
        'Close the connection
        frmMain.sckClient.Close
    End If
    
    If Mid(Data, 1, 4) = "102-" Then
        '102- Message of the Day
        strData = Mid(Data, 5)
        frmWelcome.txtMOTD = strData
    End If
    
    If Mid(Data, 1, 4) = "103-" Then
        '103- News Headline
        strData = Mid(Data, 5)
        frmWelcome.lblHeadline = strData
    End If
    
    
    If Mid(Data, 1, 4) = "104-" Then
        '104- News Details
        strData = Mid(Data, 5)
        frmWelcome.txtMOTD = strData
    End If
    
    If Mid(Data, 1, 4) = "300-" Then
        '300- Incoming Instant Message
        '300- user:message
        strData = Mid(Data, 5)
        For Each frm In Forms
            If frm.Tag = GetLeft(strData, ":") Then
                frm.txtIncoming = frm.txtIncoming & vbNewLine & GetRight(strData, ":")
                booFound = True
                Exit For
            End If
        Next frm
        
        If Not booFound = True Then
            Dim frmNewIM As New frmIM
                frmNewIM.Tag = GetLeft(strData, ":")
                frmNewIM.txtIncoming = frmNewIM.txtIncoming & vbNewLine & frmNewIM.Tag & ": " & GetRight(strData, ":")
                frmNewIM.Caption = frmNewIM.Tag & " - Instant Message"
                frmNewIM.Show
        End If
        
    End If
    
    If Mid(Data, 1, 4) = "301-" Then
        '301- Incoming user info
        strData = Mid(Data, 5)
        Dim strTmp
        strTmp = Split(strData, ",")
        
        Dim frmNewInfo As New frmInfo
            frmNewInfo.txtOnline = strTmp(0)
            frmNewInfo.txtName = strTmp(1)
            frmNewInfo.txtTitle = strTmp(2)
            frmNewInfo.txtDepartment = strTmp(3)
            frmNewInfo.txtLocation = strTmp(4)
    End If

    If Mid(Data, 1, 4) = "400-" Then
        '101- List of rooms
        strData = Mid(Data, 5)
        'If waiting is displayed clear list
        'If frmRooms.lstRooms.Selected(1).Text = "Please wait while room list is populated..." Then frmRooms.lstRooms.Clear
        'Add room to the list
        lngCount = frmRooms.lstRooms.ListItems.Count + 1
        Call frmRooms.lstRooms.ListItems.Add(lngCount, , GetLeft(strData, ":"))
        Call frmRooms.lstRooms.ListItems(lngCount).ListSubItems.Add(1, , GetRight(strData, ":"))
    End If

    If Mid(Data, 1, 4) = "401-" Then
        '401- Incoming User List
        strData = Mid(Data, 5)
        'Loop through all the forms to find the room
        For Each frm In Forms
            If frm.Tag = GetLeft(strData, ",") Then
                'Remove the room's name from the list
                strData = Replace(strData, frm.Tag & ",", "")
                'Found the room now add the users to the list
                Do
                    Pos1 = Pos2 + Len(",")
                    Pos2 = InStr(Pos1, strData, ",")
                    If Pos2 <> 0 Then
                        Call frm.lstUsers.AddItem(Mid(strData, Pos1, Pos2 - Pos1))
                    End If
                Loop Until Pos2 = 0
            End If
        Next frm
    End If
    
    If Mid(Data, 1, 4) = "402-" Then
        '402- User enters
        strData = Mid(Data, 5)
        Call UserEntersRoom(GetLeft(strData, ","), GetRight(strData, ","))
    End If
    
    If Mid(Data, 1, 4) = "403-" Then
        '403- User exits
        strData = Mid(Data, 5)
        Call UserExitsRoom(GetLeft(strData, ","), GetRight(strData, ","))
    End If
    
    If Mid(Data, 1, 4) = "404-" Then
        '404- Incoming message
        strData = Mid(Data, 5)
        Call NewChatMessage(GetLeft(strData, ","), GetRight(strData, ","))
    End If
    
    If Mid(Data, 1, 4) = "405-" Then
        '405- Access denied to meeting room
        strData = Mid(Data, 5)
        Call ErrorDialog("Login", strData)
    End If
    
    If Mid(Data, 1, 4) = "406-" Then
        '406- Meeting room access granted
        strData = Mid(Data, 5)
        Dim frmNewChat As New frmRoom
            frmNewChat.Tag = strData
            frmNewChat.Caption = "Meeting Room [" & strData & "]"
            frmNewChat.txtIncoming = "Welcome to Meeting Room: " & strData
            frmNewChat.Show
    End If
    
    If Mid(Data, 1, 4) = "408-" Then
        '404- Incoming description
        strData = Mid(Data, 5)
        Call NewRoomCation(GetLeft(strData, ","), GetRight(strData, ","))
    End If
    
    If Mid(Data, 1, 4) = "410-" Then
        '410- Incoming drawing
        strData = Mid(Data, 5)
        For Each frm In Forms
            If frm.Tag = GetLeft(strData, ":") Then
                'Found the room. Draw what was sent
                Call frm.DrawOnMe(GetRight(strData, ":"))
            End If
        Next frm
    End If

    
    If Mid(Data, 1, 4) = "500-" Then
        '101- Handshake denied, display why
        strData = Mid(Data, 5)
        Call ErrorDialog("IM", strData)
    End If

End Sub

Public Sub AddBar(frmName, strCaption)
    On Error Resume Next
    Call mdiParent.tskBar.Buttons.Add(, frmName, strCaption, , 6)
End Sub

Public Sub RemoveBar(frmName, strCaption)
On Error Resume Next

Dim lngCount As Integer
    
    For lngCount = 1 To mdiParent.tskBar.Buttons.Count
        If mdiParent.tskBar.Buttons(lngCount).Key = frmName Then
            mdiParent.tskBar.Buttons.Remove (lngCount)
            Exit Sub
        End If
    Next
End Sub

Public Sub NewRoomCation(strRoom, strUserAndMessage)
Dim strUser As String, strMessage As String, frm As Form
'Find the room in the forms
For Each frm In Forms
    If frm.Tag = strRoom Then
        'There was a match, add text to room
        'and change title
        strUser = GetLeft(strUserAndMessage, ",")
        strMessage = GetRight(strUserAndMessage, ",")
        frm.lblTopic = strMessage
        Exit For
    End If
Next frm
End Sub

Public Sub NewChatMessage(strRoom, strUserAndMessage)
Dim strUser As String, strMessage As String, frm As Form
'Find the room in the forms
For Each frm In Forms
    If frm.Tag = strRoom Then
        'There was a match, add text to room
        strUser = GetLeft(strUserAndMessage, ",")
        strMessage = GetRight(strUserAndMessage, ",")
        frm.txtIncoming = frm.txtIncoming & vbNewLine & strUser & ": " & strMessage
        Exit For
    End If
Next frm
End Sub

Public Sub UserEntersRoom(strRoom, strWho)
Dim frm As Form
'Find the room in the forms
For Each frm In Forms
    If frm.Tag = strRoom Then
        'There was a match, advise user
        frm.txtIncoming = frm.txtIncoming & vbNewLine & strWho & " has entered the room"
        'Add to user list
        frm.lstUsers.AddItem (strWho)
        Exit For
    End If
Next frm
End Sub

Public Sub UserExitsRoom(strRoom, strWho)
On Error Resume Next
Dim lngCount As Integer, frm As Form
'Find the room in the forms
For Each frm In Forms
    If frm.Tag = strRoom Then
        'There was a match, advise user
        frm.txtIncoming = frm.txtIncoming & vbNewLine & strWho & " has exited the room"
        'Remove user from list
        For lngCount = 1 To frm.lstUsers.ListCount
            If frm.lstUsers.List(lngCount) = strWho Then
                frm.lstUsers.RemoveItem (lngCount)
                Exit For
            End If
        Next lngCount
    End If
Next frm
End Sub
