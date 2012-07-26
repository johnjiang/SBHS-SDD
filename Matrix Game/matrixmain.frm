VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form matrixmain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Path of the Real World - The Story Continues"
   ClientHeight    =   7500
   ClientLeft      =   4965
   ClientTop       =   2265
   ClientWidth     =   10485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "matrixmain.frx":0000
   ScaleHeight     =   7500
   ScaleWidth      =   10485
   StartUpPosition =   2  'CenterScreen
   Begin MCI.MMControl mmcMP3 
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   1085
      _Version        =   393216
      PlayEnabled     =   -1  'True
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PlayVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Image screen 
      Height          =   1215
      Index           =   6
      Left            =   1560
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Image screen 
      Height          =   855
      Index           =   5
      Left            =   9360
      Top             =   4560
      Width           =   615
   End
   Begin VB.Image screen 
      Height          =   1215
      Index           =   4
      Left            =   7800
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Image screen 
      Height          =   1335
      Index           =   3
      Left            =   3480
      Top             =   3840
      Width           =   3615
   End
   Begin VB.Image screen 
      Height          =   1455
      Index           =   2
      Left            =   4560
      Top             =   360
      Width           =   1455
   End
   Begin VB.Image screen 
      Height          =   855
      Index           =   1
      Left            =   6480
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label window 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   1575
      Left            =   4440
      TabIndex        =   0
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Image screen 
      Height          =   855
      Index           =   0
      Left            =   2760
      Top             =   2520
      Width           =   1215
   End
End
Attribute VB_Name = "matrixmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    window.Caption = "Select a screen to continue..."
    For i% = 0 To 6
        screen(i%).BorderStyle = 0
    Next
End Sub

Private Sub screen_Click(Index As Integer)
    Select Case Index
        Case 0
            sentdodge.Show
            matrixmain.Enabled = False
        Case 1
            shootagents.Show
            matrixmain.Enabled = False
        Case 2
            shipdock.Show
            matrixmain.Enabled = False
        Case 3
            refuel.Show
            matrixmain.Enabled = False
        Case 4
            matrixabout.Show
            matrixmain.Enabled = False
        Case 5
            matrixcouncil.Show
            matrixmain.Enabled = False
        Case 6
            mmcMP3.DeviceType = "MPEGVideo" '\\Change MCI device type to MPEG
            mmcMP3.FileName = "sounds\music.mx" '\\designate file to be played
            mmcMP3.Command = "Open" '\\Open file for playing
            mmcMP3.Command = "Play" '\\Play file
    End Select
End Sub

Private Sub Form_QueryUnLoad(cancel As Integer, UnloadMode As Integer)
    If MsgBox("Are you sure you want to return to the Matrix?", vbQuestion + vbYesNo, "Exit") = vbNo Then
        cancel = True
    End If
End Sub

Private Sub form_unload(cancel As Integer)
    mmcMP3.Command = "Stop" '\\Stop playing the file
    mmcMP3.Command = "Close" '\\Close the file
    Unload Me
End Sub

Private Sub screen_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
        Case 0
            window.Caption = "Dodge The Sentinels"
            screen(0).BorderStyle = 1
        Case 1
            window.Caption = "Shoot the Agents"
            screen(1).BorderStyle = 1
        Case 2
            window.Caption = "Dock The Hovercraft"
            screen(2).BorderStyle = 1
        Case 3
            window.Caption = "Fuel The Osiris"
            screen(3).BorderStyle = 1
        Case 4
            window.Caption = "About"
            screen(4).BorderStyle = 1
        Case 5
            window.Caption = "Seek Help From Council"
            screen(5).BorderStyle = 1
        Case 6
            window.Caption = "Play Music"
            screen(6).BorderStyle = 1
    End Select
End Sub
