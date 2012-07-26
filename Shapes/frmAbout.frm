VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Jay Jay's Area and Volume Calculator"
   ClientHeight    =   4320
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5760
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   Picture         =   "frmAbout.frx":0000
   ScaleHeight     =   288
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   384
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   5280
      Top             =   0
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00F3E5CE&
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3360
      Width           =   1260
   End
   Begin VB.Label lbltime 
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label lbldate 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   0
      X2              =   370.933
      Y1              =   208
      Y2              =   208
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":2845
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1170
      Left            =   840
      TabIndex        =   1
      Top             =   1560
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Application Title"
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
      Height          =   480
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   5445
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   7
      X2              =   376.933
      Y1              =   164
      Y2              =   164
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1050
      TabIndex        =   4
      Top             =   960
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      BackStyle       =   0  'Transparent
      Caption         =   "Warning: This program should only be used for educational purposes. Suitable for children between the age of 5 - 12."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   825
      Left            =   240
      TabIndex        =   2
      Top             =   3240
      Width           =   3870
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_ALPHA = &H2
Private Const WS_EX_TRANSPARENT = &H20&
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Option Explicit

Private Sub cmdOK_Click()
    Unload Me
End Sub
Private Sub form_unload(cancel As Integer)
    Form1.Enabled = True
End Sub

Private Sub Form_Load()
    lbltime.Caption = Time
    lbldate.Caption = Format(Now, "mm-dd-yyyy")
    frmAbout.Icon = Form1.Icon
    Me.Caption = "About " & App.Title
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
    Dim L As Long
'set WS_EX_LAYERED on this window
    L = SetWindowLong(Me.hwnd, GWL_EXSTYLE, GetWindowLong(hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
'increase/decrease the 50 to make the window more or less "visible"
'ie a value of 1 is invisible, but still recieves the mouse click...
    L = SetLayeredWindowAttributes(Me.hwnd, 0, 200, LWA_ALPHA)
'0 to 255. 255 is 100% visible. 0 is 0% visible
    Me.Refresh
End Sub

Private Sub Timer1_Timer()
    If lbltime.Caption <> Time Then
        lbltime.Caption = Time
    End If
End Sub
