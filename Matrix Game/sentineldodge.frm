VERSION 5.00
Begin VB.Form sentdodge 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dodge The Sentinels"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   9810
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "sentineldodge.frx":0000
   ScaleHeight     =   6120
   ScaleWidth      =   9810
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdrestart 
      Caption         =   "Restart"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6960
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin MatrixGameRoom.cTransPictureBox imgdennis 
      Height          =   660
      Left            =   4680
      TabIndex        =   1
      Top             =   5400
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   1164
      PictureFile     =   "sentineldodge.frx":17D38
      TransparentColor=   0
      BackColor       =   0
      Style           =   1
      SyncTransColor  =   -1  'True
      BorderStyle     =   0
      MouseIcon       =   "sentineldodge.frx":17ED3
      MousePointer    =   0
      Enabled         =   -1  'True
      Object.Width           =   495
      Object.Height          =   660
   End
   Begin MatrixGameRoom.cTransPictureBox senti 
      Height          =   300
      Index           =   0
      Left            =   240
      TabIndex        =   12
      Top             =   1080
      Visible         =   0   'False
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   529
      PictureFile     =   "sentineldodge.frx":17EEF
      TransparentColor=   0
      BackColor       =   0
      Style           =   1
      SyncTransColor  =   -1  'True
      BorderStyle     =   0
      MouseIcon       =   "sentineldodge.frx":18330
      MousePointer    =   0
      Enabled         =   -1  'True
      Object.Width           =   450
      Object.Height          =   300
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Back"
      Height          =   375
      Left            =   8280
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   375
      Left            =   5640
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   9240
      Top             =   5520
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   8520
      Top             =   5520
   End
   Begin MatrixGameRoom.cTransPictureBox senti 
      Height          =   300
      Index           =   1
      Left            =   360
      TabIndex        =   13
      Top             =   1560
      Visible         =   0   'False
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   529
      PictureFile     =   "sentineldodge.frx":1834C
      TransparentColor=   0
      BackColor       =   0
      Style           =   1
      SyncTransColor  =   -1  'True
      BorderStyle     =   0
      MouseIcon       =   "sentineldodge.frx":1878D
      MousePointer    =   0
      Enabled         =   -1  'True
      Object.Width           =   450
      Object.Height          =   300
   End
   Begin MatrixGameRoom.cTransPictureBox senti 
      Height          =   300
      Index           =   2
      Left            =   360
      TabIndex        =   14
      Top             =   1560
      Visible         =   0   'False
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   529
      PictureFile     =   "sentineldodge.frx":187A9
      TransparentColor=   0
      BackColor       =   0
      Style           =   1
      SyncTransColor  =   -1  'True
      BorderStyle     =   0
      MouseIcon       =   "sentineldodge.frx":18BEA
      MousePointer    =   0
      Enabled         =   -1  'True
      Object.Width           =   450
      Object.Height          =   300
   End
   Begin MatrixGameRoom.cTransPictureBox senti 
      Height          =   300
      Index           =   3
      Left            =   360
      TabIndex        =   15
      Top             =   1560
      Visible         =   0   'False
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   529
      PictureFile     =   "sentineldodge.frx":18C06
      TransparentColor=   0
      BackColor       =   0
      Style           =   1
      SyncTransColor  =   -1  'True
      BorderStyle     =   0
      MouseIcon       =   "sentineldodge.frx":19047
      MousePointer    =   0
      Enabled         =   -1  'True
      Object.Width           =   450
      Object.Height          =   300
   End
   Begin MatrixGameRoom.cTransPictureBox senti 
      Height          =   300
      Index           =   4
      Left            =   360
      TabIndex        =   16
      Top             =   1560
      Visible         =   0   'False
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   529
      PictureFile     =   "sentineldodge.frx":19063
      TransparentColor=   0
      BackColor       =   0
      Style           =   1
      SyncTransColor  =   -1  'True
      BorderStyle     =   0
      MouseIcon       =   "sentineldodge.frx":194A4
      MousePointer    =   0
      Enabled         =   -1  'True
      Object.Width           =   450
      Object.Height          =   300
   End
   Begin MatrixGameRoom.cTransPictureBox senti 
      Height          =   300
      Index           =   5
      Left            =   1680
      TabIndex        =   17
      Top             =   1080
      Visible         =   0   'False
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   529
      PictureFile     =   "sentineldodge.frx":194C0
      TransparentColor=   0
      BackColor       =   0
      Style           =   1
      SyncTransColor  =   -1  'True
      BorderStyle     =   0
      MouseIcon       =   "sentineldodge.frx":198FE
      MousePointer    =   0
      Enabled         =   -1  'True
      Object.Width           =   450
      Object.Height          =   300
   End
   Begin MatrixGameRoom.cTransPictureBox senti 
      Height          =   300
      Index           =   6
      Left            =   1680
      TabIndex        =   18
      Top             =   1440
      Visible         =   0   'False
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   529
      PictureFile     =   "sentineldodge.frx":1991A
      TransparentColor=   0
      BackColor       =   0
      Style           =   1
      SyncTransColor  =   -1  'True
      BorderStyle     =   0
      MouseIcon       =   "sentineldodge.frx":19D58
      MousePointer    =   0
      Enabled         =   -1  'True
      Object.Width           =   450
      Object.Height          =   300
   End
   Begin MatrixGameRoom.cTransPictureBox senti 
      Height          =   300
      Index           =   7
      Left            =   1680
      TabIndex        =   19
      Top             =   1440
      Visible         =   0   'False
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   529
      PictureFile     =   "sentineldodge.frx":19D74
      TransparentColor=   0
      BackColor       =   0
      Style           =   1
      SyncTransColor  =   -1  'True
      BorderStyle     =   0
      MouseIcon       =   "sentineldodge.frx":1A1B2
      MousePointer    =   0
      Enabled         =   -1  'True
      Object.Width           =   450
      Object.Height          =   300
   End
   Begin MatrixGameRoom.cTransPictureBox senti 
      Height          =   300
      Index           =   8
      Left            =   1680
      TabIndex        =   20
      Top             =   1440
      Visible         =   0   'False
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   529
      PictureFile     =   "sentineldodge.frx":1A1CE
      TransparentColor=   0
      BackColor       =   0
      Style           =   1
      SyncTransColor  =   -1  'True
      BorderStyle     =   0
      MouseIcon       =   "sentineldodge.frx":1A60C
      MousePointer    =   0
      Enabled         =   -1  'True
      Object.Width           =   450
      Object.Height          =   300
   End
   Begin MatrixGameRoom.cTransPictureBox senti 
      Height          =   300
      Index           =   9
      Left            =   1680
      TabIndex        =   21
      Top             =   1440
      Visible         =   0   'False
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   529
      PictureFile     =   "sentineldodge.frx":1A628
      TransparentColor=   0
      BackColor       =   0
      Style           =   1
      SyncTransColor  =   -1  'True
      BorderStyle     =   0
      MouseIcon       =   "sentineldodge.frx":1AA66
      MousePointer    =   0
      Enabled         =   -1  'True
      Object.Width           =   450
      Object.Height          =   300
   End
   Begin MatrixGameRoom.cTransPictureBox senti 
      Height          =   300
      Index           =   10
      Left            =   2400
      TabIndex        =   22
      Top             =   960
      Visible         =   0   'False
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   529
      PictureFile     =   "sentineldodge.frx":1AA82
      TransparentColor=   0
      BackColor       =   0
      Style           =   1
      SyncTransColor  =   -1  'True
      BorderStyle     =   0
      MouseIcon       =   "sentineldodge.frx":1AEC3
      MousePointer    =   0
      Enabled         =   -1  'True
      Object.Width           =   450
      Object.Height          =   300
   End
   Begin MatrixGameRoom.cTransPictureBox senti 
      Height          =   300
      Index           =   11
      Left            =   2520
      TabIndex        =   23
      Top             =   1440
      Visible         =   0   'False
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   529
      PictureFile     =   "sentineldodge.frx":1AEDF
      TransparentColor=   0
      BackColor       =   0
      Style           =   1
      SyncTransColor  =   -1  'True
      BorderStyle     =   0
      MouseIcon       =   "sentineldodge.frx":1B320
      MousePointer    =   0
      Enabled         =   -1  'True
      Object.Width           =   450
      Object.Height          =   300
   End
   Begin MatrixGameRoom.cTransPictureBox senti 
      Height          =   300
      Index           =   12
      Left            =   2520
      TabIndex        =   24
      Top             =   1440
      Visible         =   0   'False
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   529
      PictureFile     =   "sentineldodge.frx":1B33C
      TransparentColor=   0
      BackColor       =   0
      Style           =   1
      SyncTransColor  =   -1  'True
      BorderStyle     =   0
      MouseIcon       =   "sentineldodge.frx":1B77D
      MousePointer    =   0
      Enabled         =   -1  'True
      Object.Width           =   450
      Object.Height          =   300
   End
   Begin MatrixGameRoom.cTransPictureBox senti 
      Height          =   300
      Index           =   13
      Left            =   2520
      TabIndex        =   25
      Top             =   1440
      Visible         =   0   'False
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   529
      PictureFile     =   "sentineldodge.frx":1B799
      TransparentColor=   0
      BackColor       =   0
      Style           =   1
      SyncTransColor  =   -1  'True
      BorderStyle     =   0
      MouseIcon       =   "sentineldodge.frx":1BBDA
      MousePointer    =   0
      Enabled         =   -1  'True
      Object.Width           =   450
      Object.Height          =   300
   End
   Begin MatrixGameRoom.cTransPictureBox senti 
      Height          =   300
      Index           =   14
      Left            =   2520
      TabIndex        =   26
      Top             =   1440
      Visible         =   0   'False
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   529
      PictureFile     =   "sentineldodge.frx":1BBF6
      TransparentColor=   0
      BackColor       =   0
      Style           =   1
      SyncTransColor  =   -1  'True
      BorderStyle     =   0
      MouseIcon       =   "sentineldodge.frx":1C037
      MousePointer    =   0
      Enabled         =   -1  'True
      Object.Width           =   450
      Object.Height          =   300
   End
   Begin MatrixGameRoom.cTransPictureBox senti 
      Height          =   300
      Index           =   15
      Left            =   3240
      TabIndex        =   27
      Top             =   960
      Visible         =   0   'False
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   529
      PictureFile     =   "sentineldodge.frx":1C053
      TransparentColor=   0
      BackColor       =   0
      Style           =   1
      SyncTransColor  =   -1  'True
      BorderStyle     =   0
      MouseIcon       =   "sentineldodge.frx":1C491
      MousePointer    =   0
      Enabled         =   -1  'True
      Object.Width           =   450
      Object.Height          =   300
   End
   Begin MatrixGameRoom.cTransPictureBox senti 
      Height          =   300
      Index           =   16
      Left            =   3240
      TabIndex        =   28
      Top             =   1320
      Visible         =   0   'False
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   529
      PictureFile     =   "sentineldodge.frx":1C4AD
      TransparentColor=   0
      BackColor       =   0
      Style           =   1
      SyncTransColor  =   -1  'True
      BorderStyle     =   0
      MouseIcon       =   "sentineldodge.frx":1C8EB
      MousePointer    =   0
      Enabled         =   -1  'True
      Object.Width           =   450
      Object.Height          =   300
   End
   Begin MatrixGameRoom.cTransPictureBox senti 
      Height          =   300
      Index           =   17
      Left            =   3240
      TabIndex        =   29
      Top             =   1320
      Visible         =   0   'False
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   529
      PictureFile     =   "sentineldodge.frx":1C907
      TransparentColor=   0
      BackColor       =   0
      Style           =   1
      SyncTransColor  =   -1  'True
      BorderStyle     =   0
      MouseIcon       =   "sentineldodge.frx":1CD45
      MousePointer    =   0
      Enabled         =   -1  'True
      Object.Width           =   450
      Object.Height          =   300
   End
   Begin MatrixGameRoom.cTransPictureBox senti 
      Height          =   300
      Index           =   18
      Left            =   3240
      TabIndex        =   30
      Top             =   1320
      Visible         =   0   'False
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   529
      PictureFile     =   "sentineldodge.frx":1CD61
      TransparentColor=   0
      BackColor       =   0
      Style           =   1
      SyncTransColor  =   -1  'True
      BorderStyle     =   0
      MouseIcon       =   "sentineldodge.frx":1D19F
      MousePointer    =   0
      Enabled         =   -1  'True
      Object.Width           =   450
      Object.Height          =   300
   End
   Begin MatrixGameRoom.cTransPictureBox senti 
      Height          =   300
      Index           =   19
      Left            =   3240
      TabIndex        =   31
      Top             =   1320
      Visible         =   0   'False
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   529
      PictureFile     =   "sentineldodge.frx":1D1BB
      TransparentColor=   0
      BackColor       =   0
      Style           =   1
      SyncTransColor  =   -1  'True
      BorderStyle     =   0
      MouseIcon       =   "sentineldodge.frx":1D5F9
      MousePointer    =   0
      Enabled         =   -1  'True
      Object.Width           =   450
      Object.Height          =   300
   End
   Begin MatrixGameRoom.cTransPictureBox senti 
      Height          =   300
      Index           =   20
      Left            =   3960
      TabIndex        =   32
      Top             =   960
      Visible         =   0   'False
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   529
      PictureFile     =   "sentineldodge.frx":1D615
      TransparentColor=   0
      BackColor       =   0
      Style           =   1
      SyncTransColor  =   -1  'True
      BorderStyle     =   0
      MouseIcon       =   "sentineldodge.frx":1DA56
      MousePointer    =   0
      Enabled         =   -1  'True
      Object.Width           =   450
      Object.Height          =   300
   End
   Begin MatrixGameRoom.cTransPictureBox senti 
      Height          =   300
      Index           =   21
      Left            =   4080
      TabIndex        =   33
      Top             =   1440
      Visible         =   0   'False
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   529
      PictureFile     =   "sentineldodge.frx":1DA72
      TransparentColor=   0
      BackColor       =   0
      Style           =   1
      SyncTransColor  =   -1  'True
      BorderStyle     =   0
      MouseIcon       =   "sentineldodge.frx":1DEB3
      MousePointer    =   0
      Enabled         =   -1  'True
      Object.Width           =   450
      Object.Height          =   300
   End
   Begin MatrixGameRoom.cTransPictureBox senti 
      Height          =   300
      Index           =   22
      Left            =   4080
      TabIndex        =   34
      Top             =   1440
      Visible         =   0   'False
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   529
      PictureFile     =   "sentineldodge.frx":1DECF
      TransparentColor=   0
      BackColor       =   0
      Style           =   1
      SyncTransColor  =   -1  'True
      BorderStyle     =   0
      MouseIcon       =   "sentineldodge.frx":1E310
      MousePointer    =   0
      Enabled         =   -1  'True
      Object.Width           =   450
      Object.Height          =   300
   End
   Begin MatrixGameRoom.cTransPictureBox senti 
      Height          =   300
      Index           =   23
      Left            =   4080
      TabIndex        =   35
      Top             =   1440
      Visible         =   0   'False
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   529
      PictureFile     =   "sentineldodge.frx":1E32C
      TransparentColor=   0
      BackColor       =   0
      Style           =   1
      SyncTransColor  =   -1  'True
      BorderStyle     =   0
      MouseIcon       =   "sentineldodge.frx":1E76D
      MousePointer    =   0
      Enabled         =   -1  'True
      Object.Width           =   450
      Object.Height          =   300
   End
   Begin MatrixGameRoom.cTransPictureBox senti 
      Height          =   300
      Index           =   24
      Left            =   4080
      TabIndex        =   36
      Top             =   1440
      Visible         =   0   'False
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   529
      PictureFile     =   "sentineldodge.frx":1E789
      TransparentColor=   0
      BackColor       =   0
      Style           =   1
      SyncTransColor  =   -1  'True
      BorderStyle     =   0
      MouseIcon       =   "sentineldodge.frx":1EBCA
      MousePointer    =   0
      Enabled         =   -1  'True
      Object.Width           =   450
      Object.Height          =   300
   End
   Begin MatrixGameRoom.cTransPictureBox senti 
      Height          =   300
      Index           =   25
      Left            =   4920
      TabIndex        =   37
      Top             =   960
      Visible         =   0   'False
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   529
      PictureFile     =   "sentineldodge.frx":1EBE6
      TransparentColor=   0
      BackColor       =   0
      Style           =   1
      SyncTransColor  =   -1  'True
      BorderStyle     =   0
      MouseIcon       =   "sentineldodge.frx":1F024
      MousePointer    =   0
      Enabled         =   -1  'True
      Object.Width           =   450
      Object.Height          =   300
   End
   Begin MatrixGameRoom.cTransPictureBox senti 
      Height          =   300
      Index           =   26
      Left            =   4920
      TabIndex        =   38
      Top             =   1320
      Visible         =   0   'False
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   529
      PictureFile     =   "sentineldodge.frx":1F040
      TransparentColor=   0
      BackColor       =   0
      Style           =   1
      SyncTransColor  =   -1  'True
      BorderStyle     =   0
      MouseIcon       =   "sentineldodge.frx":1F47E
      MousePointer    =   0
      Enabled         =   -1  'True
      Object.Width           =   450
      Object.Height          =   300
   End
   Begin MatrixGameRoom.cTransPictureBox senti 
      Height          =   300
      Index           =   27
      Left            =   4920
      TabIndex        =   39
      Top             =   1320
      Visible         =   0   'False
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   529
      PictureFile     =   "sentineldodge.frx":1F49A
      TransparentColor=   0
      BackColor       =   0
      Style           =   1
      SyncTransColor  =   -1  'True
      BorderStyle     =   0
      MouseIcon       =   "sentineldodge.frx":1F8D8
      MousePointer    =   0
      Enabled         =   -1  'True
      Object.Width           =   450
      Object.Height          =   300
   End
   Begin MatrixGameRoom.cTransPictureBox senti 
      Height          =   300
      Index           =   28
      Left            =   4920
      TabIndex        =   40
      Top             =   1320
      Visible         =   0   'False
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   529
      PictureFile     =   "sentineldodge.frx":1F8F4
      TransparentColor=   0
      BackColor       =   0
      Style           =   1
      SyncTransColor  =   -1  'True
      BorderStyle     =   0
      MouseIcon       =   "sentineldodge.frx":1FD32
      MousePointer    =   0
      Enabled         =   -1  'True
      Object.Width           =   450
      Object.Height          =   300
   End
   Begin MatrixGameRoom.cTransPictureBox senti 
      Height          =   300
      Index           =   29
      Left            =   4920
      TabIndex        =   41
      Top             =   1320
      Visible         =   0   'False
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   529
      PictureFile     =   "sentineldodge.frx":1FD4E
      TransparentColor=   0
      BackColor       =   0
      Style           =   1
      SyncTransColor  =   -1  'True
      BorderStyle     =   0
      MouseIcon       =   "sentineldodge.frx":2018C
      MousePointer    =   0
      Enabled         =   -1  'True
      Object.Width           =   450
      Object.Height          =   300
   End
   Begin VB.Label level 
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
      Left            =   720
      TabIndex        =   11
      Top             =   120
      Width           =   375
   End
   Begin VB.Label txttime 
      BackStyle       =   0  'Transparent
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
      Left            =   3120
      TabIndex        =   10
      Top             =   120
      Width           =   975
   End
   Begin VB.Label txthigh 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Left            =   3120
      TabIndex        =   9
      Top             =   360
      Width           =   975
   End
   Begin VB.Label txtscore 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      TabIndex        =   8
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "High Score    :"
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
      Left            =   1800
      TabIndex        =   7
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Score :"
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
      TabIndex        =   6
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Bonus Points :"
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
      Left            =   1800
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
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
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
   Begin VB.Shape circ 
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   27
      Left            =   480
      Shape           =   3  'Circle
      Top             =   6480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape circ 
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   5
      Left            =   120
      Shape           =   3  'Circle
      Top             =   6480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   9720
      Y1              =   720
      Y2              =   720
   End
End
Attribute VB_Name = "sentdodge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim up As Boolean
Dim down As Boolean
Dim lefty As Boolean
Dim right As Boolean
Dim a(29) As Boolean
Dim X As Integer
Dim Y As Integer
Dim i As Integer
Dim rate(29) As Integer
Dim Speed As Integer

Private Sub cmdrestart_Click()
    If MsgBox("Are you sure you would want to restart?", vbQuestion + vbYesNo, "Restart") = vbYes Then
       restart
    End If
    imgdennis.SetFocus
End Sub

Private Sub Command1_Click()
    Timer1.Enabled = True
    Timer2.Enabled = True
    imgdennis.Visible = True
    imgdennis.SetFocus
    Command1.Enabled = False
    cmdrestart.Enabled = True
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Click()
    If imgdennis.Visible = True Then
        imgdennis.SetFocus
    End If
End Sub

Private Sub imgdennis_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = 37
            lefty = True
        Case Is = 39
            right = True
        Case Is = 38
            up = True
        Case Is = 40
            down = True
    End Select
End Sub

Private Sub imgdennis_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = 37
            lefty = False
        Case Is = 39
            right = False
        Case Is = 38
            up = False
        Case Is = 40
            down = False
    End Select
End Sub

Private Sub Form_Load()
    Speed = 20
    shippos
    leftorright
    txttime.Caption = Speed * 10
    highscore
    changerate
End Sub
Private Sub Form_QueryUnLoad(cancel As Integer, UnloadMode As Integer)
    If MsgBox("Are you sure you want to unplug?", vbQuestion + vbYesNo, "Exit") = vbNo Then
        cancel = True
    Else
        matrixmain.Enabled = True
    End If
End Sub
Private Sub Timer1_Timer()
    ' movement for the pic
    If lefty = True Then X = X - 70
    If X < 0 Then
        X = 0
    End If
    If right = True Then X = X + 70
    If X >= Me.ScaleWidth - imgdennis.Width Then
        X = Me.ScaleWidth - imgdennis.Width
    End If
    If up = True Then Y = Y - 70
    If Y < 0 Then
        Y = 0
    End If
    If down = True Then Y = Y + 70
    If Y >= Me.ScaleHeight - imgdennis.Height Then
        Y = Me.ScaleHeight - imgdennis.Height
    End If
    imgdennis.Left = X
    imgdennis.Top = Y
End Sub

Private Sub Timer2_Timer()
    ' movement for the dudes
    win
    dead
    For i% = 0 To level.Caption - 1
        If a(i%) = False Then
            senti(i%).Left = senti(i%).Left + rate(i%)
            If senti(i%).Left >= Me.ScaleWidth Then
                senti(i%).Left = 0 - senti(0).Width
            End If
        Else
            senti(i%).Left = senti(i%).Left - rate(i%)
            If senti(i%).Left <= 0 - senti(0).Width Then
                senti(i%).Left = Me.ScaleWidth
            End If
        End If
    Next
    If txttime.Caption > 0 Then
        txttime.Caption = txttime.Caption - 1
    End If
End Sub

Public Sub KeyUp()
    lefty = False
    right = False
    up = False
    down = False
    For i% = 0 To level.Caption - 1
        senti(i%).Visible = False
    Next
End Sub

Public Sub SCORE()
    txtscore.Caption = txtscore.Caption + Val(txttime.Caption)
End Sub

Public Sub dead()
    For i% = 0 To level.Caption - 1
        If imgdennis.Left < senti(i%).Left + 300 And imgdennis.Left > senti(i%).Left - 300 And imgdennis.Top < senti(i%).Top + 250 And imgdennis.Top > senti(i%).Top - 400 Then
            Timer1.Enabled = False
            Timer2.Enabled = False
            highscore
            MsgBox "You've been destroyed by a sentinel! You've failed to reach Zion"
            restart
        End If
    Next
End Sub

Public Sub win()
    If imgdennis.Top < 720 Then
        If Not level.Caption = 30 Then
            KeyUp
            shippos
            SCORE
            txttime.Caption = Speed * 10
            Speed = Speed + 10
            changerate
            level.Caption = level.Caption + 1
        Else
            level.Caption = "You Win!"
            MsgBox "Congratulations you have finished the game!"
            Timer1.Enabled = False
            Timer2.Enabled = False
        End If
    End If
    For i% = 0 To level.Caption - 1
        senti(i%).Visible = True
    Next
End Sub

Public Sub highscore()
    Dim BestScore As String
    Open App.Path & "\Score.bestscorefile" For Binary Access Read As #1
    BestScore = Space$(LOF(1))
    Input #1, BestScore
    Close #1
    txthigh.Caption = BestScore
    If txtscore.Caption > Val(txthigh.Caption) Then
        MsgBox "You have beaten the highscore! You are the One!"
        Open App.Path & "\Score.bestscorefile" For Binary Access Write As #2
        Put #2, , txtscore.Caption
        Close #2
    End If
End Sub

Public Sub changerate()
    ' randomises the speed of each of the sentinels
    For i% = 0 To 29
        Randomize
        rate(i%) = Rnd * Speed + 30
    Next
End Sub

Public Sub position()
    ' randomises the position of each sentinel
    Randomize
    senti(i%).Left = Int(Rnd * Me.ScaleWidth)
    senti(i%).Top = Int(Rnd * (Me.ScaleHeight - 3 * imgdennis.Height) + 850)
End Sub

Public Sub shippos()
    ' set the position of the ship in the middle bottom of the screen
    X = Me.ScaleWidth / 2
    Y = Me.ScaleHeight - imgdennis.Height
End Sub

Public Sub restart()
    KeyUp
    Form_Load
    level.Caption = 1
    txtscore.Caption = 0
    imgdennis.Visible = False
    cmdrestart.Enabled = False
    Command1.Enabled = True
End Sub

Public Sub leftorright()
    For i% = 0 To 4
        position
    Next
    For i% = 10 To 14
        position
    Next
    For i% = 20 To 24
        position
    Next
    For i% = 5 To 9
        position
        a(i%) = True
    Next
    For i% = 15 To 19
        position
        a(i%) = True
    Next
    For i% = 25 To 29
        position
        a(i%) = True
    Next
End Sub
