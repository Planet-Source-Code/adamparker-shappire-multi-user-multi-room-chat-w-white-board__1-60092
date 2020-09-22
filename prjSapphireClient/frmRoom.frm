VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRoom 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Meeting Room"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7455
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
   ScaleHeight     =   5175
   ScaleWidth      =   7455
   Begin VB.PictureBox pctWB 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   120
      ScaleHeight     =   1635
      ScaleWidth      =   7155
      TabIndex        =   8
      Top             =   3360
      Width           =   7215
      Begin VB.PictureBox pctHolder 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1815
         Left            =   -130
         ScaleHeight     =   1785
         ScaleWidth      =   870
         TabIndex        =   9
         Top             =   -130
         Width           =   905
         Begin MSComctlLib.Slider sldSize 
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   960
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            _Version        =   393216
            Min             =   1
            Max             =   20
            SelStart        =   1
            Value           =   1
         End
         Begin VB.CommandButton cmdClear 
            Caption         =   "&Clear"
            Height          =   375
            Left            =   240
            TabIndex        =   46
            Top             =   1320
            Width           =   540
         End
         Begin VB.PictureBox pctColor 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0FF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   135
            Index           =   35
            Left            =   720
            ScaleHeight     =   135
            ScaleWidth      =   135
            TabIndex        =   45
            Top             =   720
            Width           =   135
         End
         Begin VB.PictureBox pctColor 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF80FF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   135
            Index           =   34
            Left            =   600
            ScaleHeight     =   135
            ScaleWidth      =   135
            TabIndex        =   44
            Top             =   720
            Width           =   135
         End
         Begin VB.PictureBox pctColor 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF00FF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   135
            Index           =   33
            Left            =   480
            ScaleHeight     =   135
            ScaleWidth      =   135
            TabIndex        =   43
            Top             =   720
            Width           =   135
         End
         Begin VB.PictureBox pctColor 
            Appearance      =   0  'Flat
            BackColor       =   &H00C000C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   135
            Index           =   32
            Left            =   360
            ScaleHeight     =   135
            ScaleWidth      =   135
            TabIndex        =   42
            Top             =   720
            Width           =   135
         End
         Begin VB.PictureBox pctColor 
            Appearance      =   0  'Flat
            BackColor       =   &H00800080&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   135
            Index           =   31
            Left            =   240
            ScaleHeight     =   135
            ScaleWidth      =   135
            TabIndex        =   41
            Top             =   720
            Width           =   135
         End
         Begin VB.PictureBox pctColor 
            Appearance      =   0  'Flat
            BackColor       =   &H00400040&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   135
            Index           =   30
            Left            =   120
            ScaleHeight     =   135
            ScaleWidth      =   135
            TabIndex        =   40
            Top             =   720
            Width           =   135
         End
         Begin VB.PictureBox pctColor 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   135
            Index           =   29
            Left            =   720
            ScaleHeight     =   135
            ScaleWidth      =   135
            TabIndex        =   39
            Top             =   600
            Width           =   135
         End
         Begin VB.PictureBox pctColor 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   135
            Index           =   28
            Left            =   600
            ScaleHeight     =   135
            ScaleWidth      =   135
            TabIndex        =   38
            Top             =   600
            Width           =   135
         End
         Begin VB.PictureBox pctColor 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   135
            Index           =   27
            Left            =   480
            ScaleHeight     =   135
            ScaleWidth      =   135
            TabIndex        =   37
            Top             =   600
            Width           =   135
         End
         Begin VB.PictureBox pctColor 
            Appearance      =   0  'Flat
            BackColor       =   &H00C00000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   135
            Index           =   26
            Left            =   360
            ScaleHeight     =   135
            ScaleWidth      =   135
            TabIndex        =   36
            Top             =   600
            Width           =   135
         End
         Begin VB.PictureBox pctColor 
            Appearance      =   0  'Flat
            BackColor       =   &H00800000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   135
            Index           =   25
            Left            =   240
            ScaleHeight     =   135
            ScaleWidth      =   135
            TabIndex        =   35
            Top             =   600
            Width           =   135
         End
         Begin VB.PictureBox pctColor 
            Appearance      =   0  'Flat
            BackColor       =   &H00400000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   135
            Index           =   24
            Left            =   120
            ScaleHeight     =   135
            ScaleWidth      =   135
            TabIndex        =   34
            Top             =   600
            Width           =   135
         End
         Begin VB.PictureBox pctColor 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   135
            Index           =   23
            Left            =   720
            ScaleHeight     =   135
            ScaleWidth      =   135
            TabIndex        =   33
            Top             =   480
            Width           =   135
         End
         Begin VB.PictureBox pctColor 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   135
            Index           =   22
            Left            =   600
            ScaleHeight     =   135
            ScaleWidth      =   135
            TabIndex        =   32
            Top             =   480
            Width           =   135
         End
         Begin VB.PictureBox pctColor 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   135
            Index           =   21
            Left            =   480
            ScaleHeight     =   135
            ScaleWidth      =   135
            TabIndex        =   31
            Top             =   480
            Width           =   135
         End
         Begin VB.PictureBox pctColor 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   135
            Index           =   20
            Left            =   360
            ScaleHeight     =   135
            ScaleWidth      =   135
            TabIndex        =   30
            Top             =   480
            Width           =   135
         End
         Begin VB.PictureBox pctColor 
            Appearance      =   0  'Flat
            BackColor       =   &H00808000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   135
            Index           =   19
            Left            =   240
            ScaleHeight     =   135
            ScaleWidth      =   135
            TabIndex        =   29
            Top             =   480
            Width           =   135
         End
         Begin VB.PictureBox pctColor 
            Appearance      =   0  'Flat
            BackColor       =   &H00404000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   135
            Index           =   18
            Left            =   120
            ScaleHeight     =   135
            ScaleWidth      =   135
            TabIndex        =   28
            Top             =   480
            Width           =   135
         End
         Begin VB.PictureBox pctColor 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   135
            Index           =   17
            Left            =   720
            ScaleHeight     =   135
            ScaleWidth      =   135
            TabIndex        =   27
            Top             =   360
            Width           =   135
         End
         Begin VB.PictureBox pctColor 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   135
            Index           =   16
            Left            =   600
            ScaleHeight     =   135
            ScaleWidth      =   135
            TabIndex        =   26
            Top             =   360
            Width           =   135
         End
         Begin VB.PictureBox pctColor 
            Appearance      =   0  'Flat
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   135
            Index           =   15
            Left            =   480
            ScaleHeight     =   135
            ScaleWidth      =   135
            TabIndex        =   25
            Top             =   360
            Width           =   135
         End
         Begin VB.PictureBox pctColor 
            Appearance      =   0  'Flat
            BackColor       =   &H0000C000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   135
            Index           =   14
            Left            =   360
            ScaleHeight     =   135
            ScaleWidth      =   135
            TabIndex        =   24
            Top             =   360
            Width           =   135
         End
         Begin VB.PictureBox pctColor 
            Appearance      =   0  'Flat
            BackColor       =   &H00008000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   135
            Index           =   13
            Left            =   240
            ScaleHeight     =   135
            ScaleWidth      =   135
            TabIndex        =   23
            Top             =   360
            Width           =   135
         End
         Begin VB.PictureBox pctColor 
            Appearance      =   0  'Flat
            BackColor       =   &H00004000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   135
            Index           =   12
            Left            =   120
            ScaleHeight     =   135
            ScaleWidth      =   135
            TabIndex        =   22
            Top             =   360
            Width           =   135
         End
         Begin VB.PictureBox pctColor 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   135
            Index           =   11
            Left            =   720
            ScaleHeight     =   135
            ScaleWidth      =   135
            TabIndex        =   21
            Top             =   240
            Width           =   135
         End
         Begin VB.PictureBox pctColor 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   135
            Index           =   10
            Left            =   600
            ScaleHeight     =   135
            ScaleWidth      =   135
            TabIndex        =   20
            Top             =   240
            Width           =   135
         End
         Begin VB.PictureBox pctColor 
            Appearance      =   0  'Flat
            BackColor       =   &H0000FFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   135
            Index           =   9
            Left            =   480
            ScaleHeight     =   135
            ScaleWidth      =   135
            TabIndex        =   19
            Top             =   240
            Width           =   135
         End
         Begin VB.PictureBox pctColor 
            Appearance      =   0  'Flat
            BackColor       =   &H0000C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   135
            Index           =   8
            Left            =   360
            ScaleHeight     =   135
            ScaleWidth      =   135
            TabIndex        =   18
            Top             =   240
            Width           =   135
         End
         Begin VB.PictureBox pctColor 
            Appearance      =   0  'Flat
            BackColor       =   &H00008080&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   135
            Index           =   7
            Left            =   240
            ScaleHeight     =   135
            ScaleWidth      =   135
            TabIndex        =   17
            Top             =   240
            Width           =   135
         End
         Begin VB.PictureBox pctColor 
            Appearance      =   0  'Flat
            BackColor       =   &H00004040&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   135
            Index           =   6
            Left            =   120
            ScaleHeight     =   135
            ScaleWidth      =   135
            TabIndex        =   16
            Top             =   240
            Width           =   135
         End
         Begin VB.PictureBox pctColor 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   135
            Index           =   5
            Left            =   720
            ScaleHeight     =   135
            ScaleWidth      =   135
            TabIndex        =   15
            Top             =   120
            Width           =   135
         End
         Begin VB.PictureBox pctColor 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   135
            Index           =   4
            Left            =   600
            ScaleHeight     =   135
            ScaleWidth      =   135
            TabIndex        =   14
            Top             =   120
            Width           =   135
         End
         Begin VB.PictureBox pctColor 
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   135
            Index           =   3
            Left            =   480
            ScaleHeight     =   135
            ScaleWidth      =   135
            TabIndex        =   13
            Top             =   120
            Width           =   135
         End
         Begin VB.PictureBox pctColor 
            Appearance      =   0  'Flat
            BackColor       =   &H000000C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   135
            Index           =   2
            Left            =   360
            ScaleHeight     =   135
            ScaleWidth      =   135
            TabIndex        =   12
            Top             =   120
            Width           =   135
         End
         Begin VB.PictureBox pctColor 
            Appearance      =   0  'Flat
            BackColor       =   &H00000080&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   135
            Index           =   1
            Left            =   240
            ScaleHeight     =   135
            ScaleWidth      =   135
            TabIndex        =   11
            Top             =   120
            Width           =   135
         End
         Begin VB.PictureBox pctColor 
            Appearance      =   0  'Flat
            BackColor       =   &H00000040&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   135
            Index           =   0
            Left            =   120
            ScaleHeight     =   135
            ScaleWidth      =   135
            TabIndex        =   10
            Top             =   120
            Width           =   135
         End
         Begin VB.Line lnDiv 
            Index           =   0
            X1              =   840
            X2              =   120
            Y1              =   855
            Y2              =   855
         End
      End
      Begin VB.Timer wbTimer 
         Interval        =   500
         Left            =   840
         Top             =   0
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   720
         ScaleHeight     =   345
         ScaleWidth      =   6465
         TabIndex        =   48
         Top             =   1360
         Width           =   6495
         Begin VB.Label lblInfo 
            Caption         =   "&Notice: You may experence lag when drawing."
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   30
            Width           =   6255
         End
      End
   End
   Begin VB.CommandButton cmdIfno 
      Caption         =   "I&nfo"
      Height          =   495
      Left            =   6600
      TabIndex        =   5
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton cmdIM 
      Caption         =   "&IM"
      Height          =   495
      Left            =   5880
      TabIndex        =   4
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "Change"
      Height          =   375
      Left            =   5880
      TabIndex        =   6
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Height          =   375
      Left            =   5880
      TabIndex        =   3
      Top             =   2880
      Width           =   1455
   End
   Begin VB.TextBox txtSend 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Width           =   5655
   End
   Begin VB.ListBox lstUsers 
      Height          =   1500
      IntegralHeight  =   0   'False
      Left            =   5880
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox txtIncoming 
      Height          =   2415
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   360
      Width           =   5655
   End
   Begin VB.Label lblTopic 
      Caption         =   "Unknown Topic"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   20
      Width           =   7215
   End
End
Attribute VB_Name = "frmRoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LastX As Integer
Dim LastY As Integer

Dim LneColor

Dim TransBuff As String
'White board stuff above

Private Sub cmdClear_Click()
Call SendData("410-" & Me.Tag & ":clear")
End Sub

Private Sub Form_Resize()
If Me.WindowState = 1 Then Me.Hide ' Hide if minimized
End Sub


Private Sub cmdChange_Click()
Dim frmNewChange As New frmChange
    frmNewChange.Tag = Me.Tag
    frmNewChange.Caption = "Change " & Me.Tag & "'s Details"
    frmNewChange.txtDesc = Me.lblTopic
    frmNewChange.Show
End Sub

Private Sub cmdIfno_Click()
Call SendData("301-" & Me.lstUsers.Text)
End Sub

Private Sub cmdIM_Click()
Dim frm As New frmSendIM
    frm.txtUser = Me.lstUsers.Text
    frm.Show
End Sub

Private Sub cmdSend_Click()
    Call SendData("404-" & Me.Tag & "," & frmMain.txtUser & "," & Me.txtSend)
    txtSend = ""
    txtSend.SetFocus
End Sub

Private Sub Form_Load()
LneColor = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Tell the server your leaving this room
Call SendData("403-" & Me.Tag & "," & frmMain.txtUser)
RemoveBar Me.Name, Me.Caption 'Remove to task switch
End Sub

Private Sub pctColor_Click(Index As Integer)
LneColor = pctColor(Index).BackColor
End Sub

Private Sub sldSize_Click()
pctWB.DrawWidth = sldSize.Value
End Sub

Private Sub txtIncoming_Change()
txtIncoming.SelLength = Len(txtIncoming)
End Sub

Private Sub txtSend_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdSend_Click
End Sub

Private Sub pctWB_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    'pctWB.Line (LastX, LastY)-(X, Y), LneColor
    TransBuff = TransBuff & LastX & "," & LastY & "," & X & "," & Y & "," & LneColor & "," & sldSize.Value & "¿"
End If

LastX = X
LastY = Y
End Sub

Public Sub DrawOnMe(strData)
Dim Tmp

If strData = "clear" Then pctWB.Cls: Exit Sub

DLine = Split(strData, "¿")
For i = 0 To (UBound(DLine) - 1)
    Tmp = Split(DLine(i), ",")
    pctWB.DrawWidth = Tmp(5)
    pctWB.Line (Tmp(0), Tmp(1))-(Tmp(2), Tmp(3)), Tmp(4)
Next
End Sub

Private Sub wbTimer_Timer()
If Len(TransBuff) > 0 Then
    
    If Len(TransBuff) > 1024 Then TransBuff = "": Exit Sub
    
    Call SendData("410-" & Me.Tag & ":" & TransBuff)
    TransBuff = ""
    
End If
End Sub
