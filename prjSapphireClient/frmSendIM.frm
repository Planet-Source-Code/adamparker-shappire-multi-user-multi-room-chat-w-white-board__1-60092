VERSION 5.00
Begin VB.Form frmSendIM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Send IM"
   ClientHeight    =   2910
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
   ScaleHeight     =   2910
   ScaleWidth      =   3975
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox txtMessage 
      Height          =   1575
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   600
      Width           =   3735
   End
   Begin VB.TextBox txtUser 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   4
      X1              =   120
      X2              =   3840
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label lblUser 
      Caption         =   "&Username"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmSendIM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Resize()
If Me.WindowState = 1 Then Me.Hide ' Hide if minimized
End Sub

Private Sub cmdSend_Click()
Call SendIM(txtUser, frmMain.txtUser, txtMessage)
Unload Me
End Sub

Private Sub Form_Load()
AddBar Me.Name, Me.Caption 'Add to task switch
End Sub

Private Sub Form_Unload(Cancel As Integer)
RemoveBar Me.Name, Me.Caption 'Remove to task switch
End Sub
