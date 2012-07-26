VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3780
   ClientLeft      =   4320
   ClientTop       =   2925
   ClientWidth     =   7530
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":000C
   ScaleHeight     =   3780
   ScaleWidth      =   7530
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrShow 
      Interval        =   2000
      Left            =   0
      Top             =   600
   End
   Begin VB.Label lblLicenseTo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Licensed To John Jiang"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5640
      TabIndex        =   4
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label lblProductName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Leet Speak Translator"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   3330
   End
   Begin VB.Label lblPlatform 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Designed For Windows Xp SP2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2520
      TabIndex        =   2
      Top             =   720
      Width           =   4680
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
      Height          =   285
      Left            =   6360
      TabIndex        =   1
      Top             =   1200
      Width           =   885
   End
   Begin VB.Label lblCopyright 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright - Enth E Nd Productions 2005"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      TabIndex        =   0
      Top             =   1680
      Width           =   2895
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub tmrShow_Timer()
    tmrShow.Enabled = False
    Languages.Show
    Unload Me 'unload the form (hide it)
End Sub
