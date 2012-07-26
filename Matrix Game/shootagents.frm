VERSION 5.00
Object = "{EE757A1F-B0AC-40BC-9E72-B8651740F53E}#1.0#0"; "ARProgBar.ocx"
Begin VB.Form shootagents 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shoot the Agents"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   9555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MousePointer    =   2  'Cross
   Picture         =   "shootagents.frx":0000
   ScaleHeight     =   7665
   ScaleWidth      =   9555
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   9000
      Top             =   720
   End
   Begin VB.TextBox sec 
      Height          =   495
      Left            =   8160
      TabIndex        =   17
      Text            =   "0"
      Top             =   2280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox min 
      Height          =   495
      Left            =   7680
      TabIndex        =   16
      Text            =   "0"
      Top             =   2280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdrestart 
      Caption         =   "Restart"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   5880
      Width           =   1335
   End
   Begin ARProgBarCtrl.ARProgressBar progress 
      Height          =   3735
      Left            =   600
      TabIndex        =   14
      Top             =   1560
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   6588
      Value           =   100
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseGradient     =   -1  'True
      IniColor        =   255
      EndColor        =   65535
      Orientation     =   1
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   9000
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8280
      Top             =   120
   End
   Begin VB.PictureBox ball 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1395
      Index           =   2
      Left            =   4320
      Picture         =   "shootagents.frx":20D81
      ScaleHeight     =   1395
      ScaleWidth      =   1215
      TabIndex        =   6
      Top             =   240
      Width           =   1215
   End
   Begin VB.PictureBox ball 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1515
      Index           =   1
      Left            =   6120
      Picture         =   "shootagents.frx":214D2
      ScaleHeight     =   1515
      ScaleWidth      =   1020
      TabIndex        =   5
      Top             =   240
      Width           =   1020
   End
   Begin VB.PictureBox ball 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1410
      Index           =   0
      Left            =   3120
      Picture         =   "shootagents.frx":21BD0
      ScaleHeight     =   1410
      ScaleWidth      =   990
      TabIndex        =   4
      Top             =   240
      Width           =   990
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Back"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   6240
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   5520
      Width           =   1335
   End
   Begin VB.Timer Fall 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   7560
      Top             =   120
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   2640
      TabIndex        =   13
      Top             =   480
      Width           =   225
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hacking Matrix..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   120
      TabIndex        =   12
      Top             =   480
      Width           =   2010
   End
   Begin VB.Label hack 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   2160
      TabIndex        =   11
      Top             =   480
      Width           =   465
   End
   Begin VB.Label crash 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   2280
      TabIndex        =   10
      Top             =   120
      Width           =   345
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "System Exploit in"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   2070
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Level :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   795
   End
   Begin VB.Label lbllevel 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   1080
      TabIndex        =   7
      Top             =   840
      Width           =   1665
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Score :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   870
   End
   Begin VB.Label SCORE 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   1080
      TabIndex        =   0
      Top             =   1200
      Width           =   825
   End
End
Attribute VB_Name = "shootagents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NoOnBorder As Integer
Dim Speed As Integer
Dim i As Integer
Dim number As Integer

Private Sub ball_Click(Index As Integer)
    Select Case Index
        Case 0
            number = 0
        Case 1
            number = 1
        Case 2
            number = 2
    End Select
    RandomTopFall
    SCORE.Caption = SCORE.Caption + 10 'Every time you click on the ball you win 20 points
    level
    Speed = Speed + 1 'Every time you click on the ball the fall speed of the ball increases to make the game harder
End Sub

Private Sub cmdrestart_Click()
    If MsgBox("Are you sure you would want to restart?", vbQuestion + vbYesNo, "Restart") = vbYes Then
        reset
    End If
    Fall.Enabled = True
    Timer1.Enabled = True
    Command1.Enabled = False
End Sub

Private Sub Command1_Click()
    Fall.Enabled = True
    Timer1.Enabled = True
    Timer4.Enabled = True
    Command1.Enabled = False
    For i% = 0 To 2
        ball(i%).Visible = True
    Next
    cmdrestart.Enabled = True
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Fall_Timer()
    For i% = 0 To 2
        ball(i%).Top = ball(i%).Top + Speed 'Ball fall code
        die
    Next
End Sub

Private Sub Form_Load()
    Speed = 10
    position
End Sub
Private Sub Form_QueryUnLoad(cancel As Integer, UnloadMode As Integer)
    If MsgBox("Are you sure you want to unplug?", vbQuestion + vbYesNo, "Exit") = vbNo Then
        cancel = True
    End If
End Sub
Private Sub form_unload(cancel As Integer)
    Unload Me
    matrixmain.Enabled = True
End Sub
Sub RandomTopFall()
    Randomize
    ball(number).Top = 0 - ball(number).Height 'Ball is on top of the form
    ball(number).Left = (Rnd * (Me.ScaleHeight - 1000) + 1700)
End Sub

Public Sub die()
    If ball(i%).Top > shootagents.Height Then
        progress.Value = progress.Value - 5
        ball(i%).Top = 0 - ball(i%).Height 'Ball is on top of the form
        gameover
    End If
End Sub

Public Sub gameover()
    If progress.Value = 0 Then
        MsgBox "You've lasted " & min.Text & " minutes and " & sec.Text & " seconds!"
        MsgBox "You are not the one..please try again later...refer to your oracle for further information..."
        reset
    End If
End Sub

Public Sub position()
    For i% = 0 To 2
        Randomize
        ball(i%).Top = 0 - ball(i%).Height 'Ball is on top of the form
        ball(i%).Left = (Rnd * (Me.ScaleHeight - 1000) + 1700)
    Next
End Sub

Public Sub level()
    lbllevel.Caption = Int(SCORE.Caption / 100 + 1)
End Sub


Private Sub Timer1_Timer()
    ' timer for the exploits and hacks
    crash.Caption = crash.Caption - 1
    hack.Caption = hack.Caption + 4
    If crash.Caption = 0 Then
        Speed = 200
        Timer2.Interval = 500
        timer
    End If
    If hack.Caption = 100 Then
        Speed = 1
        Timer2.Interval = 5000
        timer
    End If
End Sub

Private Sub Timer2_Timer()
    ' hacking matrix
    If crash.Caption = 0 Then
        crash.Caption = 10
    End If
    If hack.Caption = 100 Then
        hack.Caption = 0
    End If
    Speed = SCORE.Caption / 10 + 1
    Timer1.Enabled = True
    Timer2.Enabled = False
End Sub

Public Sub timer()
    Timer2.Enabled = True
    Timer1.Enabled = False
End Sub

Public Sub reset()
    Form_Load
    SCORE.Caption = 0
    BallFallSpeed = 10
    progress.Value = 100
    lbllevel.Caption = 1
    crash.Caption = 10
    hack.Caption = 0
    Timer4.Enabled = False
    Timer3.Enabled = False
    Timer1.Enabled = False
    Timer2.Enabled = False
    Fall.Enabled = False
    cmdrestart.Enabled = False
    Command1.Enabled = True
End Sub

Private Sub Timer4_Timer()
    sec.Text = sec.Text + 1
    If sec.Text = 60 Then
        sec.Text = 0
        min.Text = 1
    End If
End Sub
