VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form matrixsplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5535
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   11085
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "matrixsplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "matrixsplash.frx":000C
   ScaleHeight     =   5535
   ScaleWidth      =   11085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MCI.MMControl realworld 
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   661
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
   Begin VB.Timer Timer1 
      Interval        =   4000
      Left            =   240
      Top             =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "The Story Continues..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   510
      Left            =   6000
      TabIndex        =   4
      Top             =   1560
      Width           =   4635
   End
   Begin VB.Label lblCompanyProduct 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "J²exel Productions"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   6720
      TabIndex        =   3
      Top             =   240
      Width           =   3225
   End
   Begin VB.Label lblProductName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Product"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   510
      Left            =   5880
      TabIndex        =   2
      Top             =   840
      Width           =   1650
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   10080
      TabIndex        =   1
      Top             =   4800
      Width           =   885
   End
   Begin VB.Label lblCopyright 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright 2005"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   9840
      TabIndex        =   0
      Top             =   5160
      Width           =   1095
   End
End
Attribute VB_Name = "matrixsplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = App.Title
    realworld.DeviceType = "WaveAudio" '\\Change MCI device type to MPEG
    realworld.FileName = "sounds\theme.mx" '\\designate file to be played
    realworld.Command = "Open" '\\Open file for playing
    realworld.Command = "Play" '\\Play file
End Sub

Private Sub Timer1_Timer()
    matrixmain.Show
    Unload Me
End Sub
Private Sub form_unload(cancel As Integer)
    realworld.Command = "Stop" '\\Stop playing the file
    realworld.Command = "Close" '\\Close the file
End Sub
