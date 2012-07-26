VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form options 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   6630
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cancel 
      Cancel          =   -1  'True
      Caption         =   "Command1"
      Height          =   495
      Left            =   4920
      TabIndex        =   2
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Apply"
      Height          =   495
      Left            =   4920
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   6588
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "options.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Check1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Manage Users"
      TabPicture(1)   =   "options.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Update"
      TabPicture(2)   =   "options.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.CheckBox Check1 
         Caption         =   "Splash Screen"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   1455
      End
   End
End
Attribute VB_Name = "options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
