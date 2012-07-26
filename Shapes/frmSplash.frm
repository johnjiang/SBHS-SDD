VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4245
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   8895
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   Picture         =   "frmSplash.frx":000C
   ScaleHeight     =   283
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   593
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   3500
      Left            =   0
      Top             =   0
   End
   Begin VB.Label lblCompanyProduct 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "32 Minutes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   3480
      TabIndex        =   3
      Top             =   2880
      Width           =   1875
   End
   Begin VB.Label lblProductName 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Product"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   3840
      TabIndex        =   2
      Top             =   2400
      Width           =   1185
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3960
      TabIndex        =   1
      Top             =   3480
      Width           =   885
   End
   Begin VB.Label lblCopyright 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright 2005"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3720
      TabIndex        =   0
      Top             =   3840
      Width           =   1335
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()
    Load Form1
    Load Form2
    Load Form3
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = App.Title
    Form1.MouseIcon = frmSplash.MouseIcon
    Form2.MouseIcon = frmSplash.MouseIcon
    Form3.MouseIcon = frmSplash.MouseIcon
    frmAbout.MouseIcon = frmSplash.MouseIcon
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    Unload Me
    Form1.Show
End Sub
