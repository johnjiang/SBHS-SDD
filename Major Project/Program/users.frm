VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form users 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manage Users"
   ClientHeight    =   5355
   ClientLeft      =   3150
   ClientTop       =   3345
   ClientWidth     =   8745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   8745
   Begin VB.ListBox list_user 
      DataSource      =   "dblogin"
      Height          =   5130
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   2415
   End
   Begin MSAdodcLib.Adodc dblogin 
      Height          =   330
      Left            =   7440
      Top             =   4800
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
      RecordSource    =   "select * from teacher"
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
   Begin VB.CommandButton cmdcan 
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4800
      Width           =   1935
   End
   Begin TabDlg.SSTab usertab 
      Height          =   4575
      Left            =   2760
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   8070
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Add User"
      TabPicture(0)   =   "users.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "chk_admin"
      Tab(0).Control(1)=   "cmdadduser"
      Tab(0).Control(2)=   "txtuser(2)"
      Tab(0).Control(3)=   "txtuser(1)"
      Tab(0).Control(4)=   "txtuser(0)"
      Tab(0).Control(5)=   "Label4"
      Tab(0).Control(6)=   "label1"
      Tab(0).Control(7)=   "label2"
      Tab(0).Control(8)=   "label3"
      Tab(0).Control(9)=   "Image1"
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Edit User"
      TabPicture(1)   =   "users.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Image2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label5"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label6"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label7"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "txtedituser(1)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "txtedituser(0)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cmdapply"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cmddel"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "chk_admin_edit"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).ControlCount=   9
      Begin VB.CheckBox chk_admin 
         BackColor       =   &H80000009&
         Caption         =   "Administrator"
         Height          =   255
         Left            =   -72720
         TabIndex        =   7
         ToolTipText     =   "Check if user is administrator"
         Top             =   2280
         Width           =   1935
      End
      Begin VB.CheckBox chk_admin_edit 
         BackColor       =   &H80000009&
         Caption         =   "Administrator"
         Height          =   255
         Left            =   2280
         TabIndex        =   13
         ToolTipText     =   "Check if user is administrator"
         Top             =   1800
         Width           =   1935
      End
      Begin VB.CommandButton cmddel 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Delete User"
         Height          =   495
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   3840
         Width           =   1935
      End
      Begin VB.CommandButton cmdadduser 
         BackColor       =   &H00C0C0C0&
         Caption         =   "A&dd User"
         Default         =   -1  'True
         Height          =   495
         Left            =   -74760
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   3840
         Width           =   1935
      End
      Begin VB.CommandButton cmdapply 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Apply Changes"
         Height          =   495
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   3840
         Width           =   1935
      End
      Begin VB.TextBox txtedituser 
         DataSource      =   "dvddb1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   2280
         TabIndex        =   10
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox txtedituser 
         DataSource      =   "dvddb1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   2280
         PasswordChar    =   "*"
         TabIndex        =   12
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox txtuser 
         DataSource      =   "dvddb1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   -72720
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox txtuser 
         DataSource      =   "dvddb1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   -72720
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox txtuser 
         DataSource      =   "dvddb1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   -72720
         TabIndex        =   4
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Change Password :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "User :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Level :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Level :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   9
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label label1 
         BackStyle       =   0  'Transparent
         Caption         =   "User :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   3
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Password :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   2
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Reenter Password :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   1
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Image Image2 
         Height          =   4200
         Left            =   0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   5880
      End
      Begin VB.Image Image1 
         Height          =   4200
         Left            =   -75000
         Stretch         =   -1  'True
         Top             =   360
         Width           =   5880
      End
   End
End
Attribute VB_Name = "users"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim level As String
Option Explicit

Private Sub Form_Load()
    
    user_disp
      
End Sub

Private Sub cmdadduser_Click()
    
    'validates fields
    For i = 0 To 2
        If txtuser(i).Text = "" Then
            MsgBox "Please fill in all fields"
            Exit Sub
        End If
    Next
    
    'checks if passwords are identical
    If txtuser(1).Text <> txtuser(2).Text Then
        MsgBox "The password you entered do not match", vbInformation
        Exit Sub
    End If
    
    adduser
    clearfields
    user_disp
    
End Sub

Public Sub adduser()
    
    addlog
       
    'adds user
    With dblogin.Recordset
        .AddNew
        !UserName = txtuser(0).Text
        !Password = txtuser(1).Text
        !admin = chk_admin.Value
        .Update
        .Requery
    End With

End Sub

Private Sub cmdapply_Click()
    
    'confirm and validate fields
    If MsgBox("Are you sure you want to change the settings?", vbQuestion + vbYesNo, "Apply Changes") = vbYes Then
        For i = 0 To 1
            If txtedituser(i).Text = "" Then
                MsgBox "Please fill in all fields"
                Exit Sub
            End If
        Next
        edituser
        user_disp
    End If
    
End Sub

Public Sub edituser()
    
    editlog
    
    'make changes to record
    With dblogin.Recordset
        !UserName = txtedituser(0).Text
        !Password = txtedituser(1).Text
        !admin = chk_admin_edit.Value
        .Update
        .Requery
    End With
    
    dblogin.Refresh
    
End Sub

Private Sub cmddel_Click()
    
    deletelog
    'confirm deletion
    If MsgBox("Are you sure you want to delete " & "'" & list_user.Text & "' ?", vbQuestion + vbYesNo, "Are you sure you want to delete?") = vbYes Then
        
        ' delete an entry from the database
        With dblogin.Recordset
            .Delete
            .Requery
        End With
        
        user_disp
            
    End If
End Sub

Public Sub deletelog()
    
    'add delete log
    With login.dblog.Recordset
        .AddNew
        !Time = Time
        !Date = Date
        !Event = "User " & list_user.Text & " was deleted"
        !User = login.db_login.Recordset.Fields(0)
        .Update
        .Requery
    End With
    
    
    login.dblog.Refresh
    
End Sub

Public Sub editlog()
    
    'add editlog
    With login.dblog.Recordset
        .AddNew
        !Time = Time
        !Date = Date
        !Event = "User " & list_user.Text & " was edited"
        !User = login.db_login.Recordset.Fields(0)
        .Update
        .Requery
    End With
    
    user_disp
    
    login.dblog.Refresh
    
End Sub

Public Sub addlog()
    
    'add add log
    With login.dblog.Recordset
        .AddNew
        !Time = Time
        !Date = Date
        !Event = "User " & txtuser(0).Text & " was added"
        !User = login.db_login.Recordset.Fields(0)
        .Update
        .Requery
    End With
    
    login.dblog.Refresh
    
End Sub

Private Sub cmdcan_Click()
    Unload Me
    main.Show
End Sub

Public Sub clearfields()
    
    'clears all fields
    For i = 0 To 2
        txtuser(i).Text = ""
    Next
End Sub

Public Sub user_disp()
    list_user.Clear
    
    dblogin.Recordset.AbsolutePosition = 1
    
    'displays the first and last names into the list_name listbox
    Do While Not dblogin.Recordset.EOF
        list_user.AddItem dblogin.Recordset!UserName
        dblogin.Recordset.MoveNext
    Loop
    
    dblogin.Recordset.AbsolutePosition = 1
    
    list_user.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    main.Show
End Sub

Private Sub list_user_Click()
    Dim tru As Integer
    
    dblogin.Recordset.AbsolutePosition = list_user.ListIndex + 1
    txtedituser(0).Text = dblogin.Recordset.Fields("Username")
    txtedituser(1).Text = dblogin.Recordset.Fields("Password")
    If dblogin.Recordset.Fields("Admin").Value = True Then
        tru = 1
    Else
        tru = 0
    End If
    
    chk_admin_edit.Value = tru
    
    
End Sub
