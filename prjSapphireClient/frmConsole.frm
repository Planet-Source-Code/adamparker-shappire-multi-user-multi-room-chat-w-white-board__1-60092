VERSION 5.00
Begin VB.Form frmConsole 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Incoming Console"
   ClientHeight    =   3195
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
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.TextBox txt 
      Height          =   2895
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmConsole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Resize()
If Me.WindowState = 1 Then Me.Hide ' Hide if minimized
End Sub

Private Sub Form_Load()
AddBar Me.Name, Me.Caption 'Add to task switch
End Sub

Private Sub Form_Unload(Cancel As Integer)
RemoveBar Me.Name, Me.Caption 'Remove to task switch
End Sub

Private Sub txt_Change()
txt.SelLength = Len(txt)
End Sub
