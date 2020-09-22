VERSION 5.00
Begin VB.Form frmIM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Instant Message"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
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
   ScaleHeight     =   4455
   ScaleWidth      =   3975
   Begin VB.CommandButton cmdInfo 
      Caption         =   "&Info"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   3840
      Width           =   1095
   End
   Begin VB.PictureBox pctPlaceHolder 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   3675
      TabIndex        =   3
      Top             =   1800
      Width           =   3735
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox txtIncoming 
      Height          =   1575
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   120
      Width           =   3735
   End
   Begin VB.TextBox txtMessage 
      Height          =   1575
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   2160
      Width           =   3735
   End
End
Attribute VB_Name = "frmIM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Resize()
If Me.WindowState = 1 Then Me.Hide ' Hide if minimized
End Sub

Private Sub cmdInfo_Click()
    Call SendData("301-" & Me.Tag)
End Sub

Private Sub cmdSend_Click()
    txtIncoming = txtIncoming & vbNewLine & txtMessage
    Call SendIM(Me.Tag, frmMain.txtUser, txtMessage)
    txtMessage = ""
    txtMessage.SetFocus
End Sub

Private Sub Form_Load()
AddBar Me.Name, Me.Caption 'Add to task switch
End Sub

Private Sub Form_Unload(Cancel As Integer)
RemoveBar Me.Name, Me.Caption 'Remove to task switch
End Sub

Private Sub txtIncoming_Change()
txtIncoming.SelLength = Len(txtIncoming)
End Sub

Private Sub txtMessage_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then cmdSend_Click
End Sub
