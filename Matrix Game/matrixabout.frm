VERSION 5.00
Begin VB.Form matrixabout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   3285
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5235
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "matrixabout.frx":0000
   ScaleHeight     =   2267.365
   ScaleMode       =   0  'User
   ScaleWidth      =   4915.936
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   3840
      TabIndex        =   0
      Top             =   2760
      Width           =   1260
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "The Story Continues..."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   345
      Left            =   840
      TabIndex        =   4
      Top             =   480
      Width           =   3885
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "Special Thanks Goes to LancOme-PariS and Pianpupodoel for testing and support. Greatly appreciated."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   1485
      Left            =   840
      TabIndex        =   1
      Top             =   1320
      Width           =   3000
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Application Title"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   360
      Left            =   840
      TabIndex        =   2
      Top             =   120
      Width           =   3885
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   225
      Left            =   840
      TabIndex        =   3
      Top             =   900
      Width           =   3885
   End
End
Attribute VB_Name = "matrixabout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_ALPHA = &H2
Private Const WS_EX_TRANSPARENT = &H20&
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Option Explicit
Private Sub cmdOK_Click()
    matrixmain.Enabled = True
    Unload Me
End Sub

Private Sub Form_Load()
    Dim L As Long
    Me.Caption = "About " & App.Title
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
'set WS_EX_LAYERED on this window
    L = SetWindowLong(Me.hWnd, GWL_EXSTYLE, GetWindowLong(hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
'increase/decrease the 50 to make the window more or less "visible"
'ie a value of 1 is invisible, but still recieves the mouse click...
    L = SetLayeredWindowAttributes(Me.hWnd, 0, 150, LWA_ALPHA)
'0 to 255. 255 is 100% visible. 0 is 0% visible
    Me.Refresh
End Sub
