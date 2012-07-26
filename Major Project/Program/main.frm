VERSION 5.00
Begin VB.Form main 
   Caption         =   "Main"
   ClientHeight    =   6075
   ClientLeft      =   4560
   ClientTop       =   2775
   ClientWidth     =   7920
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   Picture         =   "main.frx":0000
   ScaleHeight     =   6075
   ScaleWidth      =   7920
   Begin VB.CommandButton cmd_user 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Users"
      Height          =   1335
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2640
      Width           =   3375
   End
   Begin VB.CommandButton cmdLogout 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Logout"
      Height          =   1335
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4200
      Width           =   3375
   End
   Begin VB.CommandButton cmd_log 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Action Log"
      Height          =   1335
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1080
      Width           =   3375
   End
   Begin VB.CommandButton cmd_mark 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Teacher's Mark Book"
      Height          =   1335
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2640
      Width           =   3375
   End
   Begin VB.CommandButton cmd_stuexp 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Student Explorer"
      Height          =   1335
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1080
      Width           =   3375
   End
   Begin VB.Image Image2 
      Height          =   1695
      Left            =   0
      Picture         =   "main.frx":E1042
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   1650
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Action Selection"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Current User:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      TabIndex        =   4
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_log_Click()
    Me.Hide
    log.Show
End Sub

Private Sub cmd_mark_Click()
    Marks.Visible = True
    Marks.Label1.Caption = Label1.Caption
    Me.Hide
End Sub

Private Sub cmd_stuexp_Click()
    stu_exp.Show
    Me.Hide
End Sub

Private Sub cmd_user_Click()
    users.Show
    Me.Hide
End Sub

Private Sub cmdLogout_Click()
    Unload Me
    login.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
    login.Show
End Sub

