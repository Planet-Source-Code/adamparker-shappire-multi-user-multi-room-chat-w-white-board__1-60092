VERSION 5.00
Begin VB.Form frmError 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Error"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2895
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
   ScaleHeight     =   2415
   ScaleWidth      =   2895
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox txtError 
      Height          =   1575
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Resize()
If Me.WindowState = 1 Then Me.Hide ' Hide if minimized
End Sub

Private Sub cmdClose_Click()
    'Close form
    Unload Me
End Sub

Private Sub Form_Load()
AddBar Me.Name, Me.Caption 'Add to task switch
End Sub

Private Sub Form_Unload(Cancel As Integer)
RemoveBar Me.Name, Me.Caption 'Remove to task switch
End Sub
