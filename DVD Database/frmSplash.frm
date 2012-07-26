VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form login 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6390
   ClientLeft      =   4275
   ClientTop       =   3105
   ClientWidth     =   9960
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   9960
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdanon 
      BackColor       =   &H00E3D4D2&
      Caption         =   "Log in as &Anonymous"
      Height          =   375
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5520
      Width           =   3015
   End
   Begin VB.CommandButton cmdexit 
      BackColor       =   &H00E3D4D2&
      Cancel          =   -1  'True
      Caption         =   "&Exit"
      Height          =   375
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5880
      Width           =   3015
   End
   Begin VB.CommandButton cmdlogin 
      BackColor       =   &H00E3D4D2&
      Caption         =   "&Log In"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5160
      Width           =   3015
   End
   Begin VB.TextBox txtlogin 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   7920
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   4680
      Width           =   1695
   End
   Begin VB.TextBox txtlogin 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   7920
      TabIndex        =   1
      Top             =   4200
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc dblogin 
      Height          =   330
      Left            =   0
      Top             =   5640
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=DVD.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=DVD.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Select * FROM Login ORDER by Username"
      Caption         =   "Grid"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc dblog 
      Height          =   330
      Left            =   0
      Top             =   6000
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=DVD.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=DVD.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from Log"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog picopen 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open"
      Filter          =   "All Images (*.jpeg)*.jpeg(*.gif)*.gif"
      FilterIndex     =   1
      InitDir         =   "app.path"
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Jexel Soft"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   9
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Version: 1.0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   8
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Easy DVD Database"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   7
      Top             =   360
      Width           =   3735
   End
   Begin VB.Image Image2 
      Height          =   3180
      Left            =   1320
      Picture         =   "frmSplash.frx":000C
      Top             =   2520
      Width           =   2220
   End
   Begin VB.Image Image1 
      Height          =   4050
      Left            =   5880
      Picture         =   "frmSplash.frx":1571
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   2
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Username :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   0
      Top             =   4200
      Width           =   1215
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    
    picopen.InitDir = App.Path
    
    ' loads the view and log forms
    Load view
    Load log
    
End Sub

Private Sub cmdanon_Click()
    
    'shows the view form
    view.Show
    
    'sets the constant to string
    Username = "Anonymous"
    
    'displays the username in statusbar
    check_username
    
    'adds rank to statusbar
    view.dvdstatus.Panels(2).Text = "Rank: Administrator"
    
    'hides login screen
    Me.Hide
    
    cleartext
    
End Sub

Private Sub cmdlogin_Click()
    
    'searches for the username typed in the textbox
    dblogin.RecordSource = "select * from Login where Username LIKE '%" & txtlogin(0).Text & "%' order by Username"
    dblogin.Refresh
    
    'if no username is found display message
    If dblogin.Recordset.RecordCount = 0 Then
        
        MsgBox "Invalid Username or Password", vbInformation
        Exit Sub
        
    End If
    
    'if user is banned display message
    If dblogin.Recordset.Fields(2) = "Banned" Then
        
        MsgBox "You have been banned!", vbInformation
        Exit Sub
    
    End If
    
    If txtlogin(0).Text = dblogin.Recordset.Fields(0) And txtlogin(1).Text = dblogin.Recordset.Fields(1) Then
        
        'sets username string
        Username = dblogin.Recordset.Fields(0)
        
        'displays view form
        view.Show
        
        'sets the username into the statusbar
        check_username
        
        'add rank to statusbar
        view.dvdstatus.Panels(2).Text = "Rank: " & dblogin.Recordset.Fields(2)
        
        'hides login form
        Me.Hide
        
        cleartext

    Else
    
        'displays message if password or username incorrect
        MsgBox "Invalid Username or Password"
        
    End If
    
End Sub

Private Sub cmdexit_Click()
    
    'confirm user exit
    If MsgBox("Are you sure you want to Exit?", vbQuestion + vbYesNo, "Exit") = vbYes Then
        End
    End If
    
End Sub

Private Sub txtlogin_Change(Index As Integer)
    
    'validates textbox to ensure data is entered
    If Not txtlogin(0).Text = "" And Not txtlogin(1).Text = "" Then
        cmdlogin.Enabled = True
    Else
        cmdlogin.Enabled = False
    End If
    
End Sub

Public Sub cleartext()
    
    'clears the textbox
    txtlogin(0).Text = ""
    txtlogin(1).Text = ""
    
End Sub
