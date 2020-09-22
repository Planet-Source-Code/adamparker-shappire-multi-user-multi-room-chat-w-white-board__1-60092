VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRooms 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Room List"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
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
   ScaleHeight     =   4185
   ScaleWidth      =   5175
   Begin VB.TextBox txtPassword 
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Top             =   3840
      Width           =   2295
   End
   Begin VB.TextBox txtRoom 
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Top             =   3480
      Width           =   2295
   End
   Begin VB.CommandButton cmdJoin 
      Caption         =   "&Join"
      Height          =   615
      Left            =   3840
      TabIndex        =   0
      Top             =   3480
      Width           =   1215
   End
   Begin MSComctlLib.ListView lstRooms 
      Height          =   3255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   5741
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Room Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   6068
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Room Password:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3840
      Width           =   2055
   End
   Begin VB.Label lblRoom 
      Caption         =   "Room Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3480
      Width           =   2055
   End
End
Attribute VB_Name = "frmRooms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Resize()
If Me.WindowState = 1 Then Me.Hide ' Hide if minimized
End Sub

Private Sub cmdJoin_Click()
Call SendData("402-" & txtRoom & ":" & txtPassword)
End Sub

Private Sub Form_Load()
'Request room list
Call SendData("400-Room list request")
AddBar Me.Name, Me.Caption 'Add to task switch
End Sub

Private Sub Form_Unload(Cancel As Integer)
RemoveBar Me.Name, Me.Caption 'Remove to task switch
End Sub

Private Sub lstRooms_Click()
On Error Resume Next
txtRoom = lstRooms.SelectedItem.Text
End Sub
