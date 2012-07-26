VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Jay Jay's Area and Volume Calculator"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   660
   ClientWidth     =   6945
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   Picture         =   "main.frx":038A
   ScaleHeight     =   300
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   463
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.CommandButton cmd3d 
      BackColor       =   &H00F3E5CE&
      Caption         =   "3D Shapes"
      Height          =   855
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2880
      Width           =   1935
   End
   Begin VB.CommandButton cmd2d 
      BackColor       =   &H00F3E5CE&
      Caption         =   "2D Shapes"
      Height          =   855
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Please Select Your Shape Type"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   1800
      Width           =   4695
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Begin VB.Menu mnuexit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuabout 
      Caption         =   "About"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd2d_Click()
    Form1.Hide
    Form2.Show
End Sub

Private Sub cmd3d_Click()
    Form1.Hide
    Form3.Show
End Sub

Private Sub Form_QueryUnLoad(cancel As Integer, UnloadMode As Integer)
    If MsgBox("Are you sure you want to Exit?", vbQuestion + vbYesNo, "Exit") = vbNo Then
        cancel = True
    End If
End Sub
Private Sub form_unload(cancel As Integer)
    End
End Sub

Private Sub mnuabout_Click()
    Form1.Enabled = False
    frmAbout.Show
End Sub

Private Sub mnuexit_Click()
    Unload Me
End Sub

