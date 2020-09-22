VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sapphire Login"
   ClientHeight    =   3960
   ClientLeft      =   5430
   ClientTop       =   4725
   ClientWidth     =   3000
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
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   3000
   Begin VB.PictureBox Picture1 
      Height          =   1695
      Left            =   120
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   1635
      ScaleWidth      =   2715
      TabIndex        =   7
      Top             =   120
      Width           =   2775
      Begin MSWinsockLib.Winsock sckClient 
         Left            =   2040
         Top             =   1080
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "&Connect"
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox txtServer 
      Height          =   285
      Left            =   1080
      TabIndex        =   5
      Text            =   "localhost"
      Top             =   2640
      Width           =   1815
   End
   Begin VB.TextBox txtPass 
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Text            =   "aparker"
      Top             =   2280
      Width           =   1815
   End
   Begin VB.TextBox txtUser 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Text            =   "aparker"
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Caption         =   "Code name: Sapphire Beta"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3600
      Width           =   2775
   End
   Begin VB.Label lblServer 
      Caption         =   "&Server"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   1410
   End
   Begin VB.Label lblPass 
      Caption         =   "&Password"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   2280
      Width           =   1410
   End
   Begin VB.Label lblUser 
      Caption         =   "&Username:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   1380
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Resize()
If Me.WindowState = 1 Then Me.Hide ' Hide if minimized
End Sub

Private Sub cmdConnect_Click()
    'See if user or pass is blank
    If txtUser = "" Or txtPass = "" Then
        ErrorDialog "Login", "Please check your username or password"
        Exit Sub
    End If
    'Tell winsock control to connect
    Call sckClient.Connect(txtServer, 6151)
End Sub

Private Sub Form_Load()
AddBar Me.Name, Me.Caption 'Add to task switch
End Sub

Private Sub Form_Unload(Cancel As Integer)
RemoveBar Me.Name, Me.Caption 'Remove to task switch
End Sub

Private Sub sckClient_Connect()
    'Tell the server some information about yourself
   Call SendData("100-" & txtUser & ":" & txtPass)
End Sub

Private Sub sckClient_DataArrival(ByVal bytesTotal As Long)
    Dim Pos1 As Long, Pos2 As Long, strData As String, strTmp As String
    'Get data
    Call sckClient.GetData(strData, vbString)
    'Split all the incoming data
    Do
        Pos1 = Pos2 + Len("Œ")
        Pos2 = InStr(Pos1, strData, "Œ")
        If Pos2 <> 0 Then
            strTmp = Mid(strData, Pos1, Pos2 - Pos1)
            Call ClientParse(strTmp)
            frmConsole.txt = frmConsole.txt & strTmp & vbNewLine
        End If
    Loop Until Pos2 = 0
End Sub
