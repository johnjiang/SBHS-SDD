VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form login 
   BorderStyle     =   0  'None
   Caption         =   "Sydney Boys High School"
   ClientHeight    =   5040
   ClientLeft      =   0
   ClientTop       =   -15
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "login.frx":0000
   ScaleHeight     =   5040
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_exit 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Exit"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.OptionButton opt_log 
      BackColor       =   &H8000000E&
      Caption         =   "Student"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   2760
      Picture         =   "login.frx":E1042
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2640
      Value           =   -1  'True
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc db_login 
      Height          =   330
      Left            =   0
      Top             =   600
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
      CommandType     =   8
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=studentdb.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=studentdb.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from Student"
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
   Begin VB.OptionButton opt_log 
      BackColor       =   &H8000000E&
      Caption         =   "Teacher"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   4080
      Picture         =   "login.frx":E311D
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Timer timer1 
      Interval        =   500
      Left            =   120
      Top             =   0
   End
   Begin VB.CommandButton cmd_sign 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2760
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3600
      Width           =   2535
   End
   Begin VB.TextBox txtlogin 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   3960
      PasswordChar    =   "*"
      TabIndex        =   2
      Text            =   "rish"
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox txtlogin 
      Height          =   375
      Index           =   0
      Left            =   3960
      TabIndex        =   1
      Text            =   "bob_dowdell"
      Top             =   1560
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc absences 
      Height          =   330
      Left            =   0
      Top             =   960
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
      CommandType     =   8
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=studentdb.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=studentdb.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from absences where student_id LIKE "" & 0 & "" order by date"
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
   Begin MSAdodcLib.Adodc dblog 
      Height          =   330
      Left            =   0
      Top             =   1680
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
      CommandType     =   8
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=studentdb.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=studentdb.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from log"
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
   Begin VB.Label Label4 
      Caption         =   "Label3"
      DataField       =   "student_id"
      DataSource      =   "absences"
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   2400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   1695
      Left            =   0
      Picture         =   "login.frx":E51F8
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   1650
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      DataField       =   "student_id"
      DataSource      =   "db_login"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lbl_display 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Top             =   4440
      Width           =   4455
   End
   Begin VB.Label lbl_time 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   2880
      TabIndex        =   5
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   2040
      Width           =   1095
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_exit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Load log
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    If MsgBox("Are you sure you want to exit? Make sure school has closed before exiting", vbInformation + vbOKCancel, "Exit") = vbCancel Then
        Cancel = True
    Else
        If Timer < "55200" Then
            check_reset
        End If
    End If
    
End Sub

Private Sub opt_log_Click(Index As Integer)
    txtlogin(0).SetFocus
End Sub

Private Sub Timer1_Timer()
    
    lbl_time.Caption = Time
    
    If Timer > "55200" Then
        cmd_exit.Visible = True
    End If
    
End Sub

Private Sub cmd_reset_Click()
    
    txtlogin(0).Text = ""
    txtlogin(1).Text = ""
    txtlogin(0).SetFocus
    
End Sub

Private Sub cmd_sign_Click()
    
    If opt_log(0).Value = True Then
        'searches for the username typed in the textbox
        db_login.RecordSource = "select * from student where student_id like '%" & txtlogin(0).Text & "%' "
    Else
        db_login.RecordSource = "select * from teacher where username like '%" & txtlogin(0).Text & "%' "
    End If
    
    db_login.Refresh
    
    'if no username is found display message
    If db_login.Recordset.RecordCount = 0 Then
        
        MsgBox "Invalid Username or Password", vbInformation
        Exit Sub
        
    End If
    
    'validates password
    If txtlogin(0).Text = db_login.Recordset.Fields(0) And txtlogin(1).Text = db_login.Recordset.Fields("password") Then
        
        If opt_log(1).Value = True Then
            
            cmd_reset_Click
            
            'main.show
            
            main.Show
            main.Label1.Caption = db_login.Recordset.Fields(0)
            Me.Hide
            
        Else

            register
            
            cmd_reset_Click
            
        End If

    Else
    
        'displays message if password or username incorrect
        MsgBox "Invalid Username or Password"
        
    End If

End Sub

Public Sub register()
    
    If db_login.Recordset.Fields("Present") = "True" Then
         
        lbl_display.Caption = txtlogin(0) & " has already logged in"
    
    Else
        
        If Timer < "55200" Then
            mark_present
        
            check_late
        Else
            MsgBox "You cannot login after school has ended!", vbExclamation, "Invalid Entry"
        End If
        
    End If
    
End Sub

Public Sub mark_present()
    
    lbl_display.Caption = txtlogin(0) & " has successfully logged in at " & Time

    With db_login.Recordset
        !present = "True"
        .Update
        .Requery
    End With
    
End Sub

Public Sub check_late()
    
    
    If Timer > "32400" Then
        
        With absences.Recordset
            
            .AddNew
            !student_id = txtlogin(0).Text
            !Time = Time
            !Date = Date
            !Type = "Partial"
            .Update
            .Requery
            
        End With
        
        'If stu_exp.Visible = True Then
        '    Set stu_exp.grid_absence.DataSource = absences
        'End If
    
    End If
    
End Sub

Public Sub check_reset()
           
    db_login.RecordSource = "select * from student"
    db_login.Recordset.AbsolutePosition = 1
    
    Do While Not db_login.Recordset.EOF
        
        With db_login.Recordset
            !present = "False"
            .Update
        End With
        
        db_login.Recordset.MoveNext
        
    Loop
    
End Sub

