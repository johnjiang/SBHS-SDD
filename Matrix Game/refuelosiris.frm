VERSION 5.00
Object = "{EE757A1F-B0AC-40BC-9E72-B8651740F53E}#1.0#0"; "ARProgBar.ocx"
Begin VB.Form refuel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Refuel Osiris"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   11265
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   Picture         =   "refuelosiris.frx":0000
   ScaleHeight     =   4815
   ScaleWidth      =   11265
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdrestart 
      Caption         =   "Restart"
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "Back"
      Height          =   375
      Left            =   1560
      TabIndex        =   12
      Top             =   4200
      Width           =   1335
   End
   Begin VB.PictureBox picdice 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   870
      Left            =   480
      ScaleHeight     =   840
      ScaleWidth      =   840
      TabIndex        =   11
      Top             =   2640
      Width           =   870
   End
   Begin VB.CommandButton emp 
      Caption         =   "Use EMP"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4200
      TabIndex        =   10
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton delay 
      Caption         =   "Delay Sentinel"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2880
      TabIndex        =   9
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton fuel 
      Caption         =   "Fuel Ship"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      TabIndex        =   8
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2280
      Top             =   3000
   End
   Begin MatrixGameRoom.cTransPictureBox cTransPictureBox1 
      Height          =   3300
      Left            =   8400
      TabIndex        =   6
      Top             =   1440
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   5821
      PictureFile     =   "refuelosiris.frx":9FAE
      TransparentColor=   3659210
      BackColor       =   0
      Style           =   1
      SyncTransColor  =   -1  'True
      BorderStyle     =   0
      MouseIcon       =   "refuelosiris.frx":E041
      MousePointer    =   0
      Enabled         =   -1  'True
      Object.Width           =   2760
      Object.Height          =   3300
   End
   Begin ARProgBarCtrl.ARProgressBar energy 
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   0
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   450
      Value           =   0
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
      IniColor        =   65535
      EndColor        =   255
   End
   Begin VB.CommandButton roll 
      Caption         =   "Roll Dice"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label dice 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   7
      Top             =   2880
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Energy :"
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
      Left            =   480
      TabIndex        =   5
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Time Until Sentinels Attack  :"
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
      Left            =   7080
      TabIndex        =   4
      Top             =   0
      Width           =   2535
   End
   Begin VB.Label time 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "30"
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
      Left            =   9600
      TabIndex        =   3
      Top             =   0
      Width           =   375
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
      Left            =   10080
      TabIndex        =   2
      Top             =   0
      Width           =   735
   End
End
Attribute VB_Name = "refuel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim die As Integer

Private Sub cmdback_Click()
    Unload Me
End Sub

Private Sub cmdrestart_Click()
    If MsgBox("Are you sure you would want to restart?", vbQuestion + vbYesNo, "Restart") = vbYes Then
        restart
    End If
End Sub

Private Sub delay_Click()
    time.Caption = time.Caption - 1
    Select Case dice.Caption
        Case 1
            time.Caption = time.Caption + 1
        Case 2
            time.Caption = time.Caption + 2
        Case 3
            time.Caption = time.Caption + 3
        Case 4
            time.Caption = time.Caption + 4
        Case 5
            time.Caption = time.Caption + 5
        Case 6
            time.Caption = time.Caption + 6
    End Select
    fuel_delay_off
    check
End Sub

Private Sub emp_Click()
    time.Caption = time.Caption - 1
    energy.Value = energy.Value - 50
    time.Caption = time.Caption + 30
    fuel_delay_off
End Sub

Private Sub fuel_Click()
    time.Caption = time.Caption - 1
    Select Case dice.Caption
        Case 1
            energy.Value = energy.Value + 1
        Case 2
            energy.Value = energy.Value + 2
        Case 3
            energy.Value = energy.Value + 3
        Case 4
            energy.Value = energy.Value + 4
        Case 5
            energy.Value = energy.Value + 5
        Case 6
            energy.Value = energy.Value + 6
    End Select
    fuel_delay_off
    check
End Sub

Private Sub roll_Click()
    Timer1.Enabled = True
    roll.Enabled = False
End Sub

Private Sub Form_Load()
    die = 0
    picdice.Picture = LoadPicture(App.Path & "\images\1.bmp")
End Sub
Private Sub Form_QueryUnLoad(cancel As Integer, UnloadMode As Integer)
    If MsgBox("Are you sure you want to unplug?", vbQuestion + vbYesNo, "Exit") = vbNo Then
        cancel = True
    End If
End Sub
Private Sub form_unload(cancel As Integer)
    matrixmain.Enabled = True
End Sub

Private Sub Timer1_Timer()
    Randomize
    dice.Caption = dice.Caption + 1
    die = die + 1
    If dice.Caption = 7 Then
        dice.Caption = 1
    End If
    If die = 25 Then
        dice.Caption = Int(Rnd * 6 + 1)
        Timer1.Enabled = False
        die = 0
        action
        If energy.Value >= 50 Then
            emp.Enabled = True
        End If
    End If
    dicepicture
End Sub

Public Sub action()
    Select Case dice.Caption
        Case 1
            fuel_delay_on
        Case 2
            fuel_delay_on
        Case 3
            fuel_delay_on
        Case 4
            fuel_delay_on
        Case 5
            fuel_delay_on
        Case 6
            fuel_delay_on
    End Select
End Sub

Public Sub fuel_delay_on()
    fuel.Enabled = True
    delay.Enabled = True
End Sub

Public Sub fuel_delay_off()
    fuel.Enabled = False
    delay.Enabled = False
    roll.Enabled = True
    emp.Enabled = False
End Sub

Public Sub check()
    If energy.Value = 100 Then
        MsgBox "You win"
        restart
    End If
    If time.Caption = 0 Then
        MsgBox "You have failed to fuel Osiris, Thadues's message have not been broadcasted."
        restart
    End If
End Sub

Public Sub dicepicture()
    Select Case dice.Caption
        Case 1
            picdice.Picture = LoadPicture(App.Path & "\images\1.bmp")
        Case 2
            picdice.Picture = LoadPicture(App.Path & "\images\2.bmp")
        Case 3
            picdice.Picture = LoadPicture(App.Path & "\images\3.bmp")
        Case 4
            picdice.Picture = LoadPicture(App.Path & "\images\4.bmp")
        Case 5
            picdice.Picture = LoadPicture(App.Path & "\images\5.bmp")
        Case 6
            picdice.Picture = LoadPicture(App.Path & "\images\6.bmp")
    End Select
End Sub

Public Sub restart()
    roll.Enabled = True
    time.Caption = 30
    energy.Value = 0
    Timer1.Enabled = False
    fuel.Enabled = False
    delay.Enabled = False
    emp.Enabled = False
    picdice.Picture = LoadPicture(App.Path & "\images\1.bmp")
End Sub
