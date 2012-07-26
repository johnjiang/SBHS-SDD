VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4755
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   2610
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSD.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   2610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Slider Slider1 
      Height          =   3375
      Left            =   1680
      TabIndex        =   3
      Top             =   840
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   5953
      _Version        =   393216
      Orientation     =   1
      Max             =   100
      TickStyle       =   1
      TickFrequency   =   100
      TextPosition    =   1
   End
   Begin VB.Shape Shape1 
      Height          =   1575
      Left            =   840
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label txtLow 
      BackStyle       =   0  'Transparent
      Caption         =   "Bottom"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   3840
      Width           =   615
   End
   Begin VB.Label txtTop 
      BackStyle       =   0  'Transparent
      Caption         =   "Top"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   615
   End
   Begin VB.Line lnMean 
      BorderStyle     =   3  'Dot
      X1              =   840
      X2              =   1560
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line3 
      X1              =   1080
      X2              =   1320
      Y1              =   4000
      Y2              =   4000
   End
   Begin VB.Line Line2 
      X1              =   1080
      X2              =   1320
      Y1              =   1000
      Y2              =   1000
   End
   Begin VB.Line Line1 
      X1              =   1200
      X2              =   1200
      Y1              =   1000
      Y2              =   4000
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   4815
      Left            =   0
      Picture         =   "frmSD.frx":000C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2655
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()
    Unload Me
End Sub

Private Sub Form_LostFocus()
    Unload Me
End Sub

Private Sub Image1_Click()
    Unload Me
End Sub

Private Sub Slider1_Click()
    Unload Me
End Sub

