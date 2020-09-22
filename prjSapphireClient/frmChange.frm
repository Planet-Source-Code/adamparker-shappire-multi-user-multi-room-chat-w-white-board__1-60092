VERSION 5.00
Begin VB.Form frmChange 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Room Details"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3735
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
   ScaleHeight     =   1890
   ScaleWidth      =   3735
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Text            =   "Click to set new password"
      Top             =   600
      Width           =   2535
   End
   Begin VB.TextBox txtDesc 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label lblInfo 
      Caption         =   "You can only change this information if you created the meeting room."
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   3495
   End
   Begin VB.Label lblPassword 
      Caption         =   "Password:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label lblDes 
      Caption         =   "Description:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdUpdate_Click()
If Not txtDesc = "" Then
    Call SendData("408-" & Me.Tag & ":" & txtDesc)
End If

If txtPassword = "" Or txtPassword = "Click to set new password" Then
Else
    Call SendData("407-" & Me.Tag & ":" & txtPassword)
End If
End Sub

Private Sub Form_Load()
AddBar Me, Me.Caption 'Add to task switch
End Sub

Private Sub Form_Resize()
If Me.WindowState = 1 Then Me.Hide ' Hide if minimized
End Sub

Private Sub Form_Unload(Cancel As Integer)
RemoveBar Me.Name, Me.Caption 'Remove to task switch
End Sub

Private Sub txtPassword_GotFocus()
If txtPassword = "Click to set new password" Then txtPassword = ""
End Sub

Private Sub txtPassword_LostFocus()
If txtPassword = "" Then txtPassword = "Click to set new password"
End Sub
