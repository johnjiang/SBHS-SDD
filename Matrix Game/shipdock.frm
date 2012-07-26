VERSION 5.00
Begin VB.Form shipdock 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dock the Hovercrafts"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   11985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "shipdock.frx":0000
   ScaleHeight     =   6495
   ScaleWidth      =   11985
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdrestart 
      Caption         =   "Restart"
      Enabled         =   0   'False
      Height          =   375
      Left            =   9120
      TabIndex        =   23
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "Back"
      Height          =   375
      Left            =   10440
      TabIndex        =   24
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdstart 
      Caption         =   "Start"
      Height          =   375
      Left            =   7800
      TabIndex        =   22
      Top             =   120
      Width           =   1335
   End
   Begin MatrixGameRoom.cTransPictureBox ship 
      Height          =   735
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   2040
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1296
      PictureFile     =   "shipdock.frx":11555
      TransparentColor=   16777215
      BackColor       =   0
      Style           =   1
      SyncTransColor  =   -1  'True
      BorderStyle     =   0
      MouseIcon       =   "shipdock.frx":11CA5
      MousePointer    =   0
      Enabled         =   -1  'True
      Object.Width           =   975
      Object.Height          =   735
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8400
      Top             =   5880
   End
   Begin VB.Timer sentinel 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   9240
      Top             =   5880
   End
   Begin MatrixGameRoom.cTransPictureBox ship 
      Height          =   735
      Index           =   1
      Left            =   120
      TabIndex        =   14
      Top             =   360
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1296
      PictureFile     =   "shipdock.frx":11CC1
      TransparentColor=   16777215
      BackColor       =   0
      Style           =   1
      SyncTransColor  =   -1  'True
      BorderStyle     =   0
      MouseIcon       =   "shipdock.frx":12411
      MousePointer    =   0
      Enabled         =   -1  'True
      Object.Width           =   975
      Object.Height          =   735
   End
   Begin MatrixGameRoom.cTransPictureBox ship 
      Height          =   735
      Index           =   2
      Left            =   120
      TabIndex        =   15
      Top             =   600
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1296
      PictureFile     =   "shipdock.frx":1242D
      TransparentColor=   16777215
      BackColor       =   0
      Style           =   1
      SyncTransColor  =   -1  'True
      BorderStyle     =   0
      MouseIcon       =   "shipdock.frx":12B7D
      MousePointer    =   0
      Enabled         =   -1  'True
      Object.Width           =   975
      Object.Height          =   735
   End
   Begin MatrixGameRoom.cTransPictureBox ship 
      Height          =   735
      Index           =   3
      Left            =   120
      TabIndex        =   16
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1296
      PictureFile     =   "shipdock.frx":12B99
      TransparentColor=   16777215
      BackColor       =   0
      Style           =   1
      SyncTransColor  =   -1  'True
      BorderStyle     =   0
      MouseIcon       =   "shipdock.frx":132E9
      MousePointer    =   0
      Enabled         =   -1  'True
      Object.Width           =   975
      Object.Height          =   735
   End
   Begin MatrixGameRoom.cTransPictureBox ship 
      Height          =   735
      Index           =   4
      Left            =   120
      TabIndex        =   17
      Top             =   840
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1296
      PictureFile     =   "shipdock.frx":13305
      TransparentColor=   16777215
      BackColor       =   0
      Style           =   1
      SyncTransColor  =   -1  'True
      BorderStyle     =   0
      MouseIcon       =   "shipdock.frx":13A55
      MousePointer    =   0
      Enabled         =   -1  'True
      Object.Width           =   975
      Object.Height          =   735
   End
   Begin MatrixGameRoom.cTransPictureBox ship 
      Height          =   735
      Index           =   5
      Left            =   120
      TabIndex        =   18
      Top             =   1320
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1296
      PictureFile     =   "shipdock.frx":13A71
      TransparentColor=   16777215
      BackColor       =   0
      Style           =   1
      SyncTransColor  =   -1  'True
      BorderStyle     =   0
      MouseIcon       =   "shipdock.frx":141C1
      MousePointer    =   0
      Enabled         =   -1  'True
      Object.Width           =   975
      Object.Height          =   735
   End
   Begin MatrixGameRoom.cTransPictureBox ship 
      Height          =   735
      Index           =   6
      Left            =   120
      TabIndex        =   19
      Top             =   1560
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1296
      PictureFile     =   "shipdock.frx":141DD
      TransparentColor=   16777215
      BackColor       =   0
      Style           =   1
      SyncTransColor  =   -1  'True
      BorderStyle     =   0
      MouseIcon       =   "shipdock.frx":1492D
      MousePointer    =   0
      Enabled         =   -1  'True
      Object.Width           =   975
      Object.Height          =   735
   End
   Begin MatrixGameRoom.cTransPictureBox ship 
      Height          =   735
      Index           =   7
      Left            =   120
      TabIndex        =   20
      Top             =   1800
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1296
      PictureFile     =   "shipdock.frx":14949
      TransparentColor=   16777215
      BackColor       =   0
      Style           =   1
      SyncTransColor  =   -1  'True
      BorderStyle     =   0
      MouseIcon       =   "shipdock.frx":15099
      MousePointer    =   0
      Enabled         =   -1  'True
      Object.Width           =   975
      Object.Height          =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "minutes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4440
      TabIndex        =   21
      Top             =   120
      Width           =   735
   End
   Begin VB.Label time 
      BackStyle       =   0  'Transparent
      Caption         =   "60"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4080
      TabIndex        =   12
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lbllevel 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   11
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Level :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Time Until The Sentinels Arrive :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   9
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label dock 
      BackStyle       =   0  'Transparent
      Caption         =   "                                                                                 Station 8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Index           =   7
      Left            =   1080
      TabIndex        =   8
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label dock 
      BackStyle       =   0  'Transparent
      Caption         =   $"shipdock.frx":150B5
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   5
      Left            =   6240
      TabIndex        =   6
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label dock 
      BackStyle       =   0  'Transparent
      Caption         =   "                                                                                                                Station 5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Index           =   4
      Left            =   3600
      TabIndex        =   5
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label dock 
      BackStyle       =   0  'Transparent
      Caption         =   "                                                                                                                Station 4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Index           =   3
      Left            =   10080
      TabIndex        =   4
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Label dock 
      BackStyle       =   0  'Transparent
      Caption         =   $"shipdock.frx":15151
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   2
      Left            =   8280
      TabIndex        =   3
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Label dock 
      BackStyle       =   0  'Transparent
      Caption         =   "                                                  Station 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Index           =   1
      Left            =   2520
      TabIndex        =   2
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Label dock 
      BackStyle       =   0  'Transparent
      Caption         =   "                                                                                                                Station 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Index           =   0
      Left            =   4320
      TabIndex        =   1
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Label station 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Zion Dock"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4920
      TabIndex        =   0
      Top             =   5880
      Width           =   3495
   End
   Begin VB.Label dock 
      BackStyle       =   0  'Transparent
      Caption         =   "                                                                                 Station 7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Index           =   6
      Left            =   8520
      TabIndex        =   7
      Top             =   2880
      Width           =   1815
   End
End
Attribute VB_Name = "shipdock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim Speed(8) As Integer
Dim number As Integer
Dim docked As Integer
Dim rndspeed As Integer

Private Sub cmdback_Click()
    Unload Me
End Sub

Private Sub cmdrestart_Click()
    If MsgBox("Are you sure you would want to restart?", vbQuestion + vbYesNo, "Restart") = vbYes Then
        restart
    End If
End Sub

Private Sub cmdstart_Click()
    sentinel.Enabled = True
    Timer2.Enabled = True
    visible_t
    cmdstart.Enabled = False
    cmdrestart.Enabled = True
End Sub

Private Sub Form_Load()
    rndspeed = 50
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

Private Sub dock_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    If Source.Index = dock(Index).Index Then
        number = dock(Index).Index
        dropship
    End If
End Sub

Private Sub ship_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    station.Caption = "Dock on Station " & Index + 1
    ship(Index).Drag
End Sub

Private Sub sentinel_Timer()
    For i% = 0 To 7
        If ship(i%).Enabled = True Then
            ship(i%).Left = ship(i%).Left - Speed(i%)
            If ship(i%).Left < 0 - ship(i%).Width Then
                ship(i%).Left = Me.ScaleWidth
            End If
        End If
    Next
    level
End Sub

Public Sub dropship()
        ship(number).Left = dock(number).Left
        ship(number).Top = dock(number).Top - 100
        ship(number).Enabled = False
        docked = docked + 1
End Sub

Public Sub level()
    If docked = 8 Then
        docked = 0
        Timer2.Enabled = False
        MsgBox "You have ddocked the ships in time and annihilated the wave of sentinels! Now for the next wave!"
        lbllevel.Caption = lbllevel.Caption + 1
        rndspeed = rndspeed + 10
        time.Caption = 60
        position
        Timer2.Enabled = True
    End If
End Sub

Private Sub Timer2_Timer()
    time.Caption = time.Caption - 1
    If time.Caption = 0 Then
        sentinel.Enabled = False
        Timer2.Enabled = False
        visible_f
        Me.Picture = LoadPicture(App.Path & "\images\rev_11schwarm.jpg")
        MsgBox "You have failed to deploy the ships in time. The Sentinels have taken over the dock! Zion has been defeated, the only hope now is Neo..."
    End If
End Sub

Public Sub visible_t()
    For i% = 0 To 7
        ship(i%).Visible = True
        dock(i%).Visible = True
    Next
    station.Visible = True
End Sub

Public Sub visible_f()
    For i% = 0 To 7
        ship(i%).Visible = False
        dock(i%).Visible = False
    Next
    station.Visible = False
End Sub

Public Sub restart()
    Form_Load
    lbllevel.Caption = 1
    time.Caption = 60
    rndspeed = 50
    Me.Picture = LoadPicture(App.Path & "\images\nebdocking_reveal_ghull_fin.jpg")
    visible_t
    sentinel.Enabled = True
    Timer2.Enabled = True
End Sub

Public Sub position()
    For i% = 0 To 7
        Randomize
        ship(i%).Left = shipdock.ScaleWidth * Rnd
        ship(i%).Top = (shipdock.ScaleHeight - ship(i%).Top) * Rnd
        Speed(i%) = Rnd * rndspeed + 50
        ship(i%).Enabled = True
    Next
End Sub
