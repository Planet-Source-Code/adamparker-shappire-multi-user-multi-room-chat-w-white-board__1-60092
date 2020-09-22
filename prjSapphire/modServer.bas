Attribute VB_Name = "modServer"
'======================
'Project Sapphire
'======================
'
'lstAccounts - Index List
'
'0 - Index - If the user is online, then a Index number will be issued
'1 - Username - The username
'2 - Password - The password
'3 - Online - Are they online? Yes or No will be displayed
'4 - Full Name - Their full name
'5 - Title - Their company title
'6 - Department - The department they are in
'7 - Floor / Room - What floor they are in and the room number EX: Floor 2 Room 102
'8 - Meeting Room - What meeting room they are in
'
'lstRooms - Index List
'
'0 - Room Name - Name of the room
'1 - Topic - Description of the room
'2 - Password - Is a password required? If not then No will be displayed, otherwise the password will be kept here
'3 - Hidden - Is this room hidden off the general population list? Yes or No will be displayed
'4 - # of Users
'5 - Created By (also 'Seat of Power')
'
'Protocol
'
'100- Reserved for normal communication between the server and client
'   100- username:pass Handshake
'   101- Handshake denied, possibly due to wrong username/password
'   102- Message of the Day
'   103- Headline
'   104- Headline Details
'200- Reserved for internal mail/offline messages system
'   200- Tells user they have a new message (server to client)
'   201- Asks the server to send the messages (client to server)
'   201- Sender:Subject:Message (server delivers mail to client)
'   202- Sender:Subject:Message (client sends mail to server)
'300- Reserved for instant messaging
'   300- server/client
'   301-online,name,title,department,floor/room (user info) to client
'   301-user (user info) to server
'400- Reserved for meeting rooms
'   400-Please send me the list (list of rooms request)
'   401-room_name:description (room list)
'   401-lobby,jdoe,sbond,rthomas (user list)
'   402-lobby,jdoe (enters)
'   403-lobby,jdoe (exits)
'   404-lobby,jdoe,message (message sent to room)
'   405-access to room denied
'   406-access granted
'   407-room:password (request to change room password)
'            - sends 500- back if denied
'   408-room:description (request to change room description)
'            - sends 408- to everyone if approved
'            - sends 500- back if denied
'   409-room:user (request to remove user)
'            - sends 500-back if denied
'            - sends 409-room to user if it goes through
'   410-room:data (white board)
'500- Errors
'

Public Sub LoadRooms()
Call frmMain.lstRooms.ListItems.Add(1, , "GeneralLobby")
Call frmMain.lstRooms.ListItems(1).ListSubItems.Add(1, , "A place where co-workers can talk about current events.")
Call frmMain.lstRooms.ListItems(1).ListSubItems.Add(2, , "No")
Call frmMain.lstRooms.ListItems(1).ListSubItems.Add(3, , "No")
Call frmMain.lstRooms.ListItems(1).ListSubItems.Add(4, , "0")
Call frmMain.lstRooms.ListItems(1).ListSubItems.Add(5, , "Admin")
Call frmMain.lstRooms.ListItems.Add(2, , "ComputerHelp")
Call frmMain.lstRooms.ListItems(2).ListSubItems.Add(1, , "Get help from your IT department.")
Call frmMain.lstRooms.ListItems(2).ListSubItems.Add(2, , "No")
Call frmMain.lstRooms.ListItems(2).ListSubItems.Add(3, , "No")
Call frmMain.lstRooms.ListItems(2).ListSubItems.Add(4, , "0")
Call frmMain.lstRooms.ListItems(2).ListSubItems.Add(5, , "Admin")
Call frmMain.lstRooms.ListItems.Add(3, , "WaterCooler")
Call frmMain.lstRooms.ListItems(3).ListSubItems.Add(1, , "Use this room to talk to co-workers during your break.")
Call frmMain.lstRooms.ListItems(3).ListSubItems.Add(2, , "No")
Call frmMain.lstRooms.ListItems(3).ListSubItems.Add(3, , "No")
Call frmMain.lstRooms.ListItems(3).ListSubItems.Add(4, , "0")
Call frmMain.lstRooms.ListItems(3).ListSubItems.Add(5, , "Admin")
End Sub
    
Public Function GetLeft(strString, strSeperater) As String
    On Error Resume Next
    GetLeft = Left$(strString, InStr(strString, strSeperater) - 1)
End Function

Public Function GetRight(strString, strSeperater) As String
    GetRight = Right(strString, Len(strString) - InStr(strString, strSeperater))
End Function

Public Sub SendData(lngIndex, strData)
Call frmMain.sckServer(lngIndex).SendData(strData & "Œ")
End Sub

Public Sub SendByUser(strUser, strData)
For lngCount = 1 To frmMain.lstAccounts.ListItems.Count
    If frmMain.lstAccounts.ListItems(lngCount).ListSubItems(1) = strUser Then
        'Found username, now send to their index
        Call SendData(frmMain.lstAccounts.ListItems(lngCount), strData)
    End If
Next lngCount
End Sub

Public Function IsOnline(strUser) As Boolean
For lngCount = 1 To frmMain.lstAccounts.ListItems.Count
    If frmMain.lstAccounts.ListItems(lngCount).ListSubItems(1) = strUser Then
        'Found username, now see if they are online
        If frmMain.lstAccounts.ListItems(lngCount).ListSubItems(3) = "Yes" Then
            IsOnline = True
            Exit Function
        Else
            IsOnline = False
            Exit Function
        End If
    End If
Next lngCount

IsOnline = False
End Function

Public Sub ComesOnline(lngIndex, strUsername, strPassword)
Dim lngCount As Integer

For lngCount = 1 To frmMain.lstAccounts.ListItems.Count
    If frmMain.lstAccounts.ListItems(lngCount).ListSubItems(1) = strUsername Then
        'Found username, now see if password matches
            If frmMain.lstAccounts.ListItems(lngCount).ListSubItems(2) = strPassword Then
                'password matches, change online status from No to Yes
                'If the account is already online, don't allow the same account on
                If IsOnline(strUsername) = True Then Call SendData(lngIndex, "101-The requested username is already in use."): Exit Sub
                frmMain.lstAccounts.ListItems(lngCount).ListSubItems(3) = "Yes"
                frmMain.lstAccounts.ListItems(lngCount).Text = lngIndex
                Call SendData(lngIndex, "100-Welcome!")
                Call SendData(lngIndex, "102-" & frmMain.txtMOTD)
                Call SendData(lngIndex, "103-" & frmMain.txtHeadline)
                '
                'check for messages, return to user
                '
            Else
                'password mis-match
                Call SendData(lngIndex, "101-Username/Password incorrect, please verify and retry.")
            End If
        Exit Sub
    End If
Next lngCount

Call SendData(lngIndex, "101-Username/Password incorrect, please verify and retry.")
End Sub

Public Sub ComesOffline(lngIndex)
Dim lngCount As Integer

For lngCount = 1 To frmMain.lstAccounts.ListItems.Count
    If frmMain.lstAccounts.ListItems(lngCount) = lngIndex Then
        'Found Index, clear out Index
        frmMain.lstAccounts.ListItems(lngCount).Text = ""
        'change online status to No
        frmMain.lstAccounts.ListItems(lngCount).ListSubItems(3) = "No"
        'Notify rooms that user left
        Call NotifySignOff(frmMain.lstAccounts.ListItems(lngCount).ListSubItems(8), frmMain.lstAccounts.ListItems(lngCount).ListSubItems(1))
        'Erase rooms
        frmMain.lstAccounts.ListItems(lngCount).ListSubItems(8) = ""
        Exit Sub
    End If
Next lngCount

End Sub

Public Sub SendInfo(Index, strLookup)
Dim strTmp As String
For lngCount = 1 To frmMain.lstAccounts.ListItems.Count
    If frmMain.lstAccounts.ListItems(lngCount).ListSubItems(1) = strLookup Then
        strTmp = frmMain.lstAccounts.ListItems(lngCount).ListSubItems(3)
        strTmp = strTmp & "," & frmMain.lstAccounts.ListItems(lngCount).ListSubItems(4)
        strTmp = strTmp & "," & frmMain.lstAccounts.ListItems(lngCount).ListSubItems(5)
        strTmp = strTmp & "," & frmMain.lstAccounts.ListItems(lngCount).ListSubItems(6)
        strTmp = strTmp & "," & frmMain.lstAccounts.ListItems(lngCount).ListSubItems(7)
        Call SendData(Index, "301-" & strTmp)
        Exit Sub
    End If
Next lngCount

Call SendData(Index, "500-User (" & GetLeft(strData, ":") & ") does not have a profile.")
End Sub

Public Sub ServerParse(Data As String, Index As Integer)
Dim lngCount As Integer, strUsers As String, strRoom As String, strPassword As String


    If Mid(Data, 1, 4) = "100-" Then
        '100- Handshake protocol, client requesting access.
        strData = Mid(Data, 5)
        Call ComesOnline(Index, GetLeft(strData, ":"), GetRight(strData, ":"))
    End If

    If Mid(Data, 1, 4) = "104-" Then
        '104- Details of headline request
        Call SendData(Index, "104-" & frmMain.txtDetails)
    End If

    If Mid(Data, 1, 4) = "300-" Then
        '300- Sending instant message
        strData = Mid(Data, 5)
        If IsOnline(GetLeft(strData, ":")) = True Then
            Call SendByUser(GetLeft(strData, ":"), "300-" & strData)
        Else
            Call SendData(Index, "500-User (" & GetLeft(strData, ":") & ") is not online.")
        End If
    End If
    
    If Mid(Data, 1, 4) = "301-" Then
        '301- Requesting user info
        strData = Mid(Data, 5)
        Call SendInfo(Index, strData)
    End If
    
    If Mid(Data, 1, 4) = "400-" Then
        '400- List of rooms request
        strData = Mid(Data, 5)
        For lngCount = 1 To frmMain.lstRooms.ListItems.Count
             Call SendData(Index, "400-" & frmMain.lstRooms.ListItems(lngCount).Text & ":" & frmMain.lstRooms.ListItems(lngCount).ListSubItems(1))
        Next lngCount
    End If
    
    If Mid(Data, 1, 4) = "402-" Then
        '402-room:password
        strData = Mid(Data, 5)
        'Load variables
        strRoom = GetLeft(strData, ":")
        strPassword = GetRight(strData, ":")
        If InRoom(strRoom, Index) = True Then
           Call SendData(Index, "405-You are already in meeting rom """ & strRoom & """, if you think this is an error, please submit a trouble ticket")
           Exit Sub
        End If
        'See if the room has a password, and if so if the supplied password was accepted
        If PasswordRoom(strRoom, strPassword, Index) = True Then
            'Password was accepted
            Call UserEnters(strRoom, Index)
            Call UpdateRoom(Index, strRoom, True)
            Call SendData(Index, "406-" & strRoom)
            strUsers = GetRoomList(strRoom)
            Call SendData(Index, "401-" & strRoom & "," & strUsers)
            Call SendData(Index, "408-" & strRoom & ",Room Topic," & GetDescription(strRoom))
        Else
            'Password denied
            Call SendData(Index, "405-The password supplied to enter meeting room """ & strRoom & """ was invalid")
        End If
    End If
        
    If Mid(Data, 1, 4) = "403-" Then
        '403- User Exits room,user
        strData = Mid(Data, 5)
        Call UserExits(GetLeft(strData, ","), GetRight(strData, ","))
        Call UpdateRoom(Index, GetLeft(strData, ","), False)
    End If

    If Mid(Data, 1, 4) = "404-" Then
        '404- lobby,jdoe,message
        strData = Mid(Data, 5)
        Call RelayMessage(GetLeft(strData, ","), GetRight(strData, ","))
    End If


    If Mid(Data, 1, 4) = "407-" Then
        '407-room:password (request to change room password)
        strData = Mid(Data, 5)
        Call ChangePassword(Index, GetLeft(strData, ":"), GetRight(strData, ":"))
    End If
    
    If Mid(Data, 1, 4) = "408-" Then
        '408-room:topic (request to change room topic)
        strData = Mid(Data, 5)
        Call ChangeDescription(Index, GetLeft(strData, ":"), GetRight(strData, ":"))
    End If
    
    If Mid(Data, 1, 4) = "409-" Then
        '409-room:user (request to remove user)
        strData = Mid(Data, 5)
        Call RejectUser(Index, strRoom, strUser)
    End If
    
    If Mid(Data, 1, 4) = "410-" Then
        '409-room:user (request to remove user)
        strData = Mid(Data, 5)
        Call SendDrawing2Room(GetLeft(strData, ":"), GetRight(strData, ":"))
    End If

    
End Sub

