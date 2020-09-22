VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sapphire Server"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   3735
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   6588
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Console View"
      TabPicture(0)   =   "frmMain.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "txtConsole"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Accounts && Rooms"
      TabPicture(1)   =   "frmMain.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstRooms"
      Tab(1).Control(1)=   "lstAccounts"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Configuration"
      TabPicture(2)   =   "frmMain.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblHeadline"
      Tab(2).Control(1)=   "lblDetails"
      Tab(2).Control(2)=   "txtMOTD"
      Tab(2).Control(3)=   "cmdSet"
      Tab(2).Control(4)=   "txtHeadline"
      Tab(2).Control(5)=   "txtDetails"
      Tab(2).ControlCount=   6
      Begin VB.TextBox txtDetails 
         Height          =   1095
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Text            =   "frmMain.frx":0054
         Top             =   2040
         Width           =   4215
      End
      Begin VB.TextBox txtHeadline 
         Height          =   285
         Left            =   -73680
         TabIndex        =   8
         Text            =   "Headline"
         Top             =   1440
         Width           =   3015
      End
      Begin VB.CommandButton cmdSet 
         Caption         =   "&Set"
         Height          =   375
         Left            =   -71640
         TabIndex        =   6
         Top             =   3240
         Width           =   975
      End
      Begin VB.TextBox txtMOTD 
         Height          =   855
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Text            =   "frmMain.frx":0072
         Top             =   480
         Width           =   4215
      End
      Begin MSComctlLib.ListView lstAccounts 
         Height          =   1575
         Left            =   -74880
         TabIndex        =   3
         Top             =   360
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   2778
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Index"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Username"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Password"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Online"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Full Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Title"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Department"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Floor / Room"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Room"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.TextBox txtConsole 
         Height          =   3255
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   360
         Width           =   4215
      End
      Begin MSComctlLib.ListView lstRooms 
         Height          =   1575
         Left            =   -74880
         TabIndex        =   4
         Top             =   2040
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   2778
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Room Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Topic"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Password"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Hidden"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "# of Users"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Created By"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblDetails 
         Caption         =   "Headline Details:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   9
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label lblHeadline 
         Caption         =   "News Headline:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   7
         Top             =   1440
         Width           =   1335
      End
   End
   Begin MSWinsockLib.Winsock sckServer 
      Index           =   0
      Left            =   4080
      Top             =   4320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblTitle 
      Caption         =   "Sapphire Server"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4185
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'======================
'Project Sapphire
'======================
'
'lstAccounts - Index List
'
'1 - Index - If the user is online, then a Index number will be issued
'2 - Username - The username
'3 - Password - The password
'4 - Online - Are they online? Yes or No will be displayed
'4 - Full Name - Their full name
'6 - Title - Their company title
'7 - Department - The department they are in
'8 - Floor / Room - What floor they are in and the room number EX: Floor 2 Room 102
'
'lstRooms - Index List
'
'1 - Room Name - Name of the room
'2 - Description - Description of the room
'3 - Password - Is a password required? If not then No will be displayed, otherwise the password will be kept here
'4 - Hidden - Is this room hidden off the general population list? Yes or No will be displayed
'
'Protocol
'
'100- Reserved for normal communication between the server and client
'200- Reserved for internal mail/offline messages system
'300- Reserved for instant messaging
'400- Reserved for meeting rooms
'
 
Private Sub cmdSet_Click()
Call SaveSetting("Sapphire", "Settings", "MOTD", txtMOTD)
Call SaveSetting("Sapphire", "Settings", "Headline", txtHeadline)
Call SaveSetting("Sapphire", "Settings", "Details", txtDetails)
End Sub

'///////////////////////////
' START GENERAL FORM SETTINGS
'///////////////////////////
'
Private Sub Form_Load()
Call lstAccounts.ListItems.Add(1, , "0")
Call lstAccounts.ListItems(1).ListSubItems.Add(1, , "aparker")
Call lstAccounts.ListItems(1).ListSubItems.Add(2, , "aparker")
Call lstAccounts.ListItems(1).ListSubItems.Add(3, , "No")
Call lstAccounts.ListItems(1).ListSubItems.Add(4, , "Adam Parker")
Call lstAccounts.ListItems(1).ListSubItems.Add(5, , "MIS Specialist")
Call lstAccounts.ListItems(1).ListSubItems.Add(6, , "MIS/IT")
Call lstAccounts.ListItems(1).ListSubItems.Add(7, , "Floor 1 Room 103")
Call lstAccounts.ListItems(1).ListSubItems.Add(8, , "0,")

Call lstAccounts.ListItems.Add(1, , "0")
Call lstAccounts.ListItems(1).ListSubItems.Add(1, , "jtest")
Call lstAccounts.ListItems(1).ListSubItems.Add(2, , "jtest")
Call lstAccounts.ListItems(1).ListSubItems.Add(3, , "No")
Call lstAccounts.ListItems(1).ListSubItems.Add(4, , "Joe Tester")
Call lstAccounts.ListItems(1).ListSubItems.Add(5, , "Sapphire Tester")
Call lstAccounts.ListItems(1).ListSubItems.Add(6, , "Software Development")
Call lstAccounts.ListItems(1).ListSubItems.Add(7, , "Floor 1 Room 103")
Call lstAccounts.ListItems(1).ListSubItems.Add(8, , "0,")



Call LoadRooms

If Not GetSetting("Sapphire", "Settings", "MOTD") = "" Then
    'Load message of the day
    txtMOTD = GetSetting("Sapphire", "Settings", "MOTD")
End If

If Not GetSetting("Sapphire", "Settings", "Headline") = "" Then
    'Load message of the day
    txtHeadline = GetSetting("Sapphire", "Settings", "Headline")
End If

If Not GetSetting("Sapphire", "Settings", "Details") = "" Then
    'Load message of the day
    txtDetails = GetSetting("Sapphire", "Settings", "Details")
End If

sckServer(0).LocalPort = "6151"
sckServer(0).Listen
txtConsole = "Sapphire: Waiting for connections on port 6151"
End Sub
'
'///////////////////////////
' END GENERAL FORM SETTINGS
'///////////////////////////


'///////////////////////////
' START ALL THINGS WINSOCK
'///////////////////////////
'

Private Sub sckServer_Close(Index As Integer)
Call ComesOffline(Index)
End Sub

Private Sub sckServer_Connect(Index As Integer)
    txtConsole = txtConsole & vbNewLine & "Connected: " & sckServer(Index).RemoteHostIP & ":" & sckServer(Index).RemotePort
    txtConsole.SelLength = Len(txtConsole)
End Sub

Private Sub sckServer_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Dim lngIndex As Long, blnFlag As Boolean
   For lngIndex& = 1 To sckServer().UBound
      If sckServer(lngIndex&).State = sckClosed Then
          blnFlag = True
          Exit For
      End If
   Next lngIndex&
   If blnFlag = False Then
      lngIndex& = sckServer().UBound + 1
      Load sckServer(lngIndex&)
   End If
Call sckServer(lngIndex&).Accept(requestID&)
Call sckServer_Connect(Index)
End Sub

Private Sub sckServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim Pos1 As Long, Pos2 As Long, strData As String, strTmp As String
    'Get data
    Call sckServer(Index).GetData(strData, vbString)
    'Split all the incoming data
    Do
        Pos1 = Pos2 + Len("Œ")
        Pos2 = InStr(Pos1, strData, "Œ")
        If Pos2 <> 0 Then
            strTmp = Mid(strData, Pos1, Pos2 - Pos1)
            Call ServerParse(strTmp, Index)
            txtConsole = txtConsole & "Index(" & Index & ") " & strTmp & vbNewLine
        End If
    Loop Until Pos2 = 0
End Sub
'
'///////////////////////////
'   END ALL THINGS WINSOCK
'///////////////////////////


