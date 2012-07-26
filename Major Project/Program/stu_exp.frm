VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form stu_exp 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Student Explorer"
   ClientHeight    =   9585
   ClientLeft      =   1230
   ClientTop       =   1515
   ClientWidth     =   12840
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9585
   ScaleWidth      =   12840
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   63
      Top             =   0
      Width           =   12840
      _ExtentX        =   22648
      _ExtentY        =   1111
      ButtonWidth     =   1032
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Back"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Edit"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Add"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Delete"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Apply"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Cancel"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stat 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   62
      Top             =   9300
      Width           =   12840
      _ExtentX        =   22648
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8819
            MinWidth        =   8819
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   30
      Left            =   0
      TabIndex        =   61
      Top             =   9555
      Width           =   12840
      _ExtentX        =   22648
      _ExtentY        =   53
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.ListBox list_name 
      DataSource      =   "student"
      Height          =   7665
      ItemData        =   "stu_exp.frx":0000
      Left            =   120
      List            =   "stu_exp.frx":0002
      TabIndex        =   30
      Top             =   1320
      Width           =   2415
   End
   Begin MSAdodcLib.Adodc student 
      Height          =   330
      Left            =   240
      Top             =   9480
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
      RecordSource    =   "select * from student order by last_name"
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
   Begin TabDlg.SSTab SSTab2 
      Height          =   8535
      Left            =   0
      TabIndex        =   31
      Top             =   720
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   15055
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Student Info"
      TabPicture(0)   =   "stu_exp.frx":0004
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Image2(0)"
      Tab(0).Control(1)=   "Label10"
      Tab(0).Control(2)=   "Label9"
      Tab(0).Control(3)=   "lbl_date"
      Tab(0).Control(4)=   "lbl_dob"
      Tab(0).Control(5)=   "Label7"
      Tab(0).Control(6)=   "img_pic"
      Tab(0).Control(7)=   "Label8"
      Tab(0).Control(8)=   "Label6"
      Tab(0).Control(9)=   "Label5"
      Tab(0).Control(10)=   "Label3"
      Tab(0).Control(11)=   "Label2"
      Tab(0).Control(12)=   "Label1"
      Tab(0).Control(13)=   "Label4"
      Tab(0).Control(14)=   "lbl_pword(0)"
      Tab(0).Control(15)=   "lbl_pword(1)"
      Tab(0).Control(16)=   "Image2(6)"
      Tab(0).Control(17)=   "Label16"
      Tab(0).Control(18)=   "Picture1"
      Tab(0).Control(19)=   "dt_date"
      Tab(0).Control(20)=   "cmb_state"
      Tab(0).Control(21)=   "txt_field(5)"
      Tab(0).Control(22)=   "txt_field(4)"
      Tab(0).Control(23)=   "txt_field(3)"
      Tab(0).Control(24)=   "txt_field(2)"
      Tab(0).Control(25)=   "txt_field(1)"
      Tab(0).Control(26)=   "txt_field(0)"
      Tab(0).Control(27)=   "txt_field(6)"
      Tab(0).Control(28)=   "txt_field(7)"
      Tab(0).Control(29)=   "txt_pword(0)"
      Tab(0).Control(30)=   "txt_pword(1)"
      Tab(0).Control(31)=   "browse"
      Tab(0).Control(32)=   "pic_open"
      Tab(0).Control(33)=   "Command1"
      Tab(0).Control(34)=   "dbmark"
      Tab(0).ControlCount=   35
      TabCaption(1)   =   "Advanced Search and Sort"
      TabPicture(1)   =   "stu_exp.frx":0020
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Image2(1)"
      Tab(1).Control(1)=   "Image2(4)"
      Tab(1).Control(2)=   "cmdsearch"
      Tab(1).Control(3)=   "cmdreset"
      Tab(1).Control(4)=   "Frame1"
      Tab(1).Control(5)=   "Frame2"
      Tab(1).Control(6)=   "Frame3"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Print"
      TabPicture(2)   =   "stu_exp.frx":003C
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Image2(2)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label15"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Image2(5)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Check3"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Check2"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Check1"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "cmd_print"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Check4"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).ControlCount=   8
      Begin MSAdodcLib.Adodc dbmark 
         Height          =   330
         Left            =   -72240
         Top             =   4440
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
         RecordSource    =   "select * from mark"
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Browse"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72240
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   3960
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.PictureBox pic_open 
         BackColor       =   &H80000009&
         Height          =   3495
         Left            =   -71040
         ScaleHeight     =   3435
         ScaleWidth      =   4755
         TabIndex        =   64
         Top             =   2160
         Visible         =   0   'False
         Width           =   4815
         Begin VB.CommandButton cmd_cancel 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Cancel"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3120
            Style           =   1  'Graphical
            TabIndex        =   69
            Top             =   2760
            Width           =   1095
         End
         Begin VB.CommandButton cmd_ok 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Okay"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3120
            Style           =   1  'Graphical
            TabIndex        =   68
            Top             =   2160
            Width           =   1095
         End
         Begin VB.DriveListBox Drive1 
            Height          =   315
            Left            =   0
            TabIndex        =   67
            Top             =   0
            Width           =   2415
         End
         Begin VB.DirListBox Dir1 
            Height          =   3015
            Left            =   0
            TabIndex        =   66
            Top             =   360
            Width           =   2415
         End
         Begin VB.FileListBox File1 
            Height          =   1260
            Left            =   2400
            Pattern         =   "*.jpg;*.jpeg;*.gif;*.bmp"
            TabIndex        =   65
            Top             =   360
            Width           =   2295
         End
      End
      Begin MSComDlg.CommonDialog browse 
         Left            =   -63120
         Top             =   840
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox txt_pword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   -63840
         PasswordChar    =   "*"
         TabIndex        =   60
         Top             =   4440
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txt_pword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   -63840
         PasswordChar    =   "*"
         TabIndex        =   59
         Top             =   3960
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txt_field 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   7
         Left            =   -67560
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   56
         Text            =   "text"
         Top             =   3960
         Width           =   855
      End
      Begin VB.TextBox txt_field 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   6
         Left            =   -67560
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   55
         Text            =   "text"
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3735
         Left            =   -72360
         TabIndex        =   50
         Top             =   480
         Width           =   9975
         Begin VB.TextBox txtfirst 
            Height          =   285
            Left            =   1320
            TabIndex        =   13
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox txtlast 
            Height          =   285
            Left            =   3720
            TabIndex        =   14
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox txt_no 
            Height          =   285
            Index           =   0
            Left            =   6360
            MaxLength       =   5
            TabIndex        =   15
            Top             =   480
            Width           =   1095
         End
         Begin VB.TextBox txt_no 
            Height          =   285
            Index           =   1
            Left            =   8520
            MaxLength       =   8
            TabIndex        =   16
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Student No:"
            Height          =   255
            Left            =   5280
            TabIndex        =   54
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "BOS No:"
            Height          =   255
            Left            =   7680
            TabIndex        =   53
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Last Name:"
            Height          =   255
            Left            =   2640
            TabIndex        =   52
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "First Name:"
            Height          =   375
            Left            =   120
            TabIndex        =   51
            Top             =   480
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sort Ascending/Descending"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   -69480
         TabIndex        =   49
         Top             =   4320
         Width           =   2775
         Begin VB.OptionButton opt_ad 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Descending"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   22
            Top             =   720
            Width           =   1695
         End
         Begin VB.OptionButton opt_ad 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Ascending"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   21
            Top             =   360
            Value           =   -1  'True
            Width           =   1695
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sort By"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   -72360
         TabIndex        =   48
         Top             =   4320
         Width           =   2775
         Begin VB.OptionButton opt_sort 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Last Name"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   17
            Top             =   360
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton opt_sort 
            BackColor       =   &H00E0E0E0&
            Caption         =   "BOS Number"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   18
            Top             =   720
            Width           =   1455
         End
         Begin VB.OptionButton opt_sort 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Student Number"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   19
            Top             =   1080
            Width           =   1935
         End
         Begin VB.OptionButton opt_sort 
            BackColor       =   &H00E0E0E0&
            Caption         =   "DOB"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   20
            Top             =   1440
            Width           =   855
         End
      End
      Begin VB.CommandButton cmdreset 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Reset"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -71160
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   7800
         Width           =   855
      End
      Begin VB.CommandButton cmdsearch 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Apply"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72240
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   7800
         Width           =   855
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "All Student Personal Information"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   26
         Top             =   2520
         Width           =   2655
      End
      Begin VB.TextBox txt_field 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   0
         Left            =   -67560
         Locked          =   -1  'True
         TabIndex        =   0
         Text            =   "text"
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txt_field 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   1
         Left            =   -64440
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "text"
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txt_field 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   2
         Left            =   -67560
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "text"
         Top             =   1560
         Width           =   3375
      End
      Begin VB.TextBox txt_field 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   3
         Left            =   -67560
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "text"
         Top             =   2040
         Width           =   1695
      End
      Begin VB.TextBox txt_field 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   4
         Left            =   -67560
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   5
         Text            =   "text"
         Top             =   2520
         Width           =   495
      End
      Begin VB.TextBox txt_field 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   5
         Left            =   -67560
         MaxLength       =   8
         TabIndex        =   7
         Text            =   "text"
         Top             =   4440
         Width           =   855
      End
      Begin VB.ComboBox cmb_state 
         Height          =   315
         ItemData        =   "stu_exp.frx":0058
         Left            =   -67560
         List            =   "stu_exp.frx":0074
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   3000
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmd_print 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   3240
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Absences"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   23
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Personal Information"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   24
         Top             =   1560
         Width           =   1815
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Student List"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   25
         Top             =   2040
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker dt_date 
         Height          =   375
         Left            =   -67560
         TabIndex        =   28
         Top             =   1080
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   24969217
         CurrentDate     =   38916
         MaxDate         =   2958464
         MinDate         =   29221
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3615
         Left            =   -72360
         ScaleHeight     =   3615
         ScaleWidth      =   9975
         TabIndex        =   32
         Top             =   4920
         Width           =   9975
         Begin VB.OptionButton opt_show 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Show Unexplained"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   6120
            TabIndex        =   11
            Top             =   360
            Width           =   1935
         End
         Begin VB.OptionButton opt_show 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Show Partial Absences"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   1200
            TabIndex        =   9
            Top             =   360
            Width           =   2055
         End
         Begin VB.OptionButton opt_show 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Show Whole Day Absences"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   3480
            TabIndex        =   10
            Top             =   360
            Width           =   2415
         End
         Begin VB.OptionButton opt_show 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Show All"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   8
            Top             =   360
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.CommandButton cmd_reason 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Add/Edit Reason"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8400
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   360
            Width           =   1455
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid_absence 
            Height          =   2775
            Left            =   0
            TabIndex        =   33
            Top             =   840
            Width           =   9975
            _ExtentX        =   17595
            _ExtentY        =   4895
            _Version        =   393216
            BackColor       =   16777215
            Rows            =   3
            Cols            =   4
            FixedRows       =   0
            FixedCols       =   0
            BackColorFixed  =   8421504
            BackColorBkg    =   12632256
            GridColor       =   8421504
            AllowBigSelection=   0   'False
            FocusRect       =   0
            HighLight       =   0
            SelectionMode   =   1
            BorderStyle     =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   4
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
            _Band(0).ColHeader=   1
            _Band(0)._NumMapCols=   6
            _Band(0)._MapCol(0)._Name=   "ID"
            _Band(0)._MapCol(0)._RSIndex=   0
            _Band(0)._MapCol(0)._Alignment=   7
            _Band(0)._MapCol(0)._Hidden=   -1  'True
            _Band(0)._MapCol(1)._Name=   "student_id"
            _Band(0)._MapCol(1)._RSIndex=   1
            _Band(0)._MapCol(1)._Alignment=   7
            _Band(0)._MapCol(1)._Hidden=   -1  'True
            _Band(0)._MapCol(2)._Name=   "time"
            _Band(0)._MapCol(2)._Caption=   "Time"
            _Band(0)._MapCol(2)._RSIndex=   2
            _Band(0)._MapCol(3)._Name=   "date"
            _Band(0)._MapCol(3)._Caption=   "Date"
            _Band(0)._MapCol(3)._RSIndex=   3
            _Band(0)._MapCol(4)._Name=   "Type"
            _Band(0)._MapCol(4)._RSIndex=   4
            _Band(0)._MapCol(5)._Name=   "Reason"
            _Band(0)._MapCol(5)._RSIndex=   5
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Student Absences:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   45
            Top             =   0
            Width           =   2535
         End
      End
      Begin VB.Label Label16 
         Caption         =   "Label16"
         DataField       =   "student_id"
         DataSource      =   "dbmark"
         Height          =   255
         Left            =   -70800
         TabIndex        =   71
         Top             =   4440
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Image Image2 
         Height          =   1695
         Index           =   6
         Left            =   -64320
         Picture         =   "stu_exp.frx":009D
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   1650
      End
      Begin VB.Image Image2 
         Height          =   1695
         Index           =   5
         Left            =   10680
         Picture         =   "stu_exp.frx":0531
         Stretch         =   -1  'True
         Top             =   6480
         Width           =   1650
      End
      Begin VB.Image Image2 
         Height          =   1695
         Index           =   4
         Left            =   -64320
         Picture         =   "stu_exp.frx":09C5
         Stretch         =   -1  'True
         Top             =   6480
         Width           =   1650
      End
      Begin VB.Label lbl_pword 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Retype Password:"
         Height          =   195
         Index           =   1
         Left            =   -65520
         TabIndex        =   58
         Top             =   4440
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.Label lbl_pword 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         Height          =   195
         Index           =   0
         Left            =   -65520
         TabIndex        =   57
         Top             =   3960
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address:"
         Height          =   195
         Left            =   -69120
         TabIndex        =   41
         Top             =   1560
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "First Name:"
         Height          =   195
         Left            =   -69120
         TabIndex        =   44
         Top             =   600
         Width           =   990
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Surname:"
         Height          =   195
         Left            =   -65640
         TabIndex        =   43
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Birth:"
         Height          =   195
         Left            =   -69120
         TabIndex        =   42
         Top             =   1080
         Width           =   1170
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "City/Town:"
         Height          =   195
         Left            =   -69120
         TabIndex        =   40
         Top             =   2040
         Width           =   945
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "State/Territory:"
         Height          =   195
         Left            =   -69120
         TabIndex        =   39
         Top             =   3000
         Width           =   1350
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone Number:"
         Height          =   195
         Left            =   -69120
         TabIndex        =   38
         Top             =   4440
         Width           =   1335
      End
      Begin VB.Image img_pic 
         Height          =   3225
         Left            =   -72240
         Stretch         =   -1  'True
         Top             =   600
         Width           =   2520
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Postcode:"
         Height          =   195
         Left            =   -69120
         TabIndex        =   37
         Top             =   2520
         Width           =   840
      End
      Begin VB.Label lbl_dob 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label15"
         DataField       =   "dob"
         DataSource      =   "student"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -67560
         TabIndex        =   2
         Top             =   1080
         Width           =   570
      End
      Begin VB.Label lbl_date 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label15"
         DataField       =   "state"
         DataSource      =   "student"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -67560
         TabIndex        =   6
         Top             =   3000
         Width           =   570
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Student Number:"
         Height          =   195
         Left            =   -69120
         TabIndex        =   36
         Top             =   3480
         Width           =   1470
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BOS Number:"
         Height          =   195
         Left            =   -69120
         TabIndex        =   35
         Top             =   3960
         Width           =   1185
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "What would you like to print?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   34
         Top             =   600
         Width           =   3855
      End
      Begin VB.Image Image2 
         Height          =   10080
         Index           =   1
         Left            =   -75000
         Picture         =   "stu_exp.frx":0E59
         Stretch         =   -1  'True
         Top             =   360
         Width           =   12960
      End
      Begin VB.Image Image2 
         Height          =   10080
         Index           =   2
         Left            =   0
         Picture         =   "stu_exp.frx":2F34
         Stretch         =   -1  'True
         Top             =   360
         Width           =   12960
      End
      Begin VB.Image Image2 
         Height          =   10080
         Index           =   0
         Left            =   -75000
         Picture         =   "stu_exp.frx":500F
         Stretch         =   -1  'True
         Top             =   360
         Width           =   12960
      End
   End
   Begin VB.Menu mnu_file 
      Caption         =   "&File"
      Begin VB.Menu mnu_back 
         Caption         =   "Go &Back"
      End
      Begin VB.Menu sep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_out 
         Caption         =   "&Log Out"
      End
   End
   Begin VB.Menu mnu_tools 
      Caption         =   "&Tools"
      Begin VB.Menu mnu_edit 
         Caption         =   "&Edit"
      End
      Begin VB.Menu mnu_del 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnu_add 
         Caption         =   "&Add"
      End
   End
   Begin VB.Menu mnu_help 
      Caption         =   "&Help"
      Begin VB.Menu mnu_use 
         Caption         =   "Terms of &Use"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_feed 
         Caption         =   "Send &Feedback"
      End
      Begin VB.Menu mnu_support 
         Caption         =   "Help and &Support"
      End
      Begin VB.Menu mnu_www 
         Caption         =   "Official &Website"
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_about 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "stu_exp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer
Dim x As String
Dim filter As String
Dim add_edit As String

Private Sub Check4_Click()
    
    If Check4.Value = 1 Then
        Check2.Enabled = False
    Else
        Check2.Enabled = True
    End If
    
End Sub

Private Sub cmdsearch_Click()
    
    If opt_ad(0) Then
        If opt_sort(0).Value Then
            student.RecordSource = "select * from student where first_name LIKE '%" & txtfirst.Text & "%' AND last_name LIKE '%" & txtlast.Text & "%' AND student_id LIKE '%" & txt_no(0).Text & "%' AND bos like '%" & txt_no(1).Text & "%' order by last_name asc"
        End If
        If opt_sort(1).Value Then
            student.RecordSource = "select * from student where first_name LIKE '%" & txtfirst.Text & "%' AND last_name LIKE '%" & txtlast.Text & "%' AND student_id LIKE '%" & txt_no(0).Text & "%' AND bos like '%" & txt_no(1).Text & "%' order by bos asc"
        End If
        If opt_sort(2).Value Then
            student.RecordSource = "select * from student where first_name LIKE '%" & txtfirst.Text & "%' AND last_name LIKE '%" & txtlast.Text & "%' AND student_id LIKE '%" & txt_no(0).Text & "%' AND bos like '%" & txt_no(1).Text & "%' order by student_id asc"
        End If
        If opt_sort(3).Value Then
            student.RecordSource = "select * from student where first_name LIKE '%" & txtfirst.Text & "%' AND last_name LIKE '%" & txtlast.Text & "%' AND student_id LIKE '%" & txt_no(0).Text & "%' AND bos like '%" & txt_no(1).Text & "%' order by dob asc"
        End If
    Else
        If opt_sort(0).Value Then
            student.RecordSource = "select * from student where first_name LIKE '%" & txtfirst.Text & "%' AND last_name LIKE '%" & txtlast.Text & "%' AND student_id LIKE '%" & txt_no(0).Text & "%' AND bos like '%" & txt_no(1).Text & "%' order by last_name desc"
        End If
        If opt_sort(1).Value Then
            student.RecordSource = "select * from student where first_name LIKE '%" & txtfirst.Text & "%' AND last_name LIKE '%" & txtlast.Text & "%' AND student_id LIKE '%" & txt_no(0).Text & "%' AND bos like '%" & txt_no(1).Text & "%' order by bos desc"
        End If
        If opt_sort(2).Value Then
            student.RecordSource = "select * from student where first_name LIKE '%" & txtfirst.Text & "%' AND last_name LIKE '%" & txtlast.Text & "%' AND student_id LIKE '%" & txt_no(0).Text & "%' AND bos like '%" & txt_no(1).Text & "%' order by student_id desc"
        End If
        If opt_sort(3).Value Then
            student.RecordSource = "select * from student where first_name LIKE '%" & txtfirst.Text & "%' AND last_name LIKE '%" & txtlast.Text & "%' AND student_id LIKE '%" & txt_no(0).Text & "%' AND bos like '%" & txt_no(1).Text & "%' order by dob desc"
        End If
    End If
    student.Refresh
    
    If student.Recordset.RecordCount <> 0 Then
        list_disp
    Else
        list_name.Clear
        MsgBox "Sorry no results found, please try a different query", vbInformation, "Sorry"
    End If
    
End Sub


Private Sub Form_Load()

    Set grid_absence.DataSource = login.absences
    list_disp
    
    organise_grid
    
    check_permission
    check_status

End Sub

Public Sub list_disp()

    student.Refresh
    list_name.Clear
    
    student.Recordset.AbsolutePosition = 1
    
    'displays the first and last names into the list_name listbox
    Do While Not student.Recordset.EOF
        list_name.AddItem student.Recordset!first_name & " " & student.Recordset!last_name
        student.Recordset.MoveNext
    Loop
    
    student.Recordset.AbsolutePosition = 1
    
    list_name.ListIndex = 0
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    If Toolbar.Buttons.Item(5).Enabled = True Then
        
        If MsgBox("You have unsaved work do you want to continue?", vbYesNo, "Exit") = vbNo Then
            Cancel = True
            Exit Sub
        End If
                
    End If
    
    main.Show
    Unload Me
    
End Sub

Private Sub grid_absence_Click()
    
    grid_absence.HighLight = flexHighlightWithFocus 'highlights selected grid
    
    'checks whether there are any entries
    If login.absences.Recordset.RecordCount <> 0 Then
        'selects row
        login.absences.Recordset.AbsolutePosition = grid_absence.Row
    End If
    
End Sub

Private Sub cmd_reason_Click()
    
    'checks to see if row selected is valid
    If grid_absence.Row = 0 Then
        MsgBox "Please select a valide entry.", vbInformation, "Error!"
        Exit Sub
    End If
    
    'creates input box
    x = InputBox("Please enter a reason for the " & login.absences.Recordset.Fields("Type") & " absence on " & login.absences.Recordset.Fields("Date") & " at " & login.absences.Recordset.Fields("Time") & " for the student " & list_name.Text & ".", "Enter Reason", login.absences.Recordset.Fields("Reason") & "")
    
    'updates record if input is not blank
    If x <> "" Then
        With login.absences.Recordset
            !Reason = x
            .Update
            .Requery
        End With
    
        login.absences.Refresh
    
        Set grid_absence.DataSource = login.absences
    
    End If
    
End Sub

Private Sub cmdreset_Click()
    
    'resets everything to original state
    student.RecordSource = "select * from student order by last_name"
    student.Refresh
    list_disp
    
End Sub

Private Sub list_name_Click()
    
    'selects record from database to be displayed
    student.Recordset.AbsolutePosition = list_name.ListIndex + 1
    
    bind_txt
    
    login.absences.RecordSource = "select * from absences where student_id LIKE " & txt_field(6).Text & " order by date"
    login.absences.Refresh

    opt_show(0).Value = True
    
End Sub

Private Sub mnu_about_Click()
frmAbout.Show
End Sub

Private Sub mnu_add_Click()
    
    'initate adds new student
    add
    add_edit = "Add"
    img_pic.Picture = LoadPicture(App.Path & "\Images\noimage.gif")
    MsgBox "Do Not Search or Print whilst adding or data may be lost!" & vbCrLf & "Press cancel or apply to search or print", vbExclamation, "Warning"
End Sub

Private Sub mnu_back_Click()
    'closes form
    Unload Me
    main.Show
    
End Sub

Private Sub mnu_del_Click()
    delete_student
End Sub

Private Sub mnu_edit_Click()
    enable_txt
    add_edit = "Edit"
    
    txt_pword(0).Text = student.Recordset.Fields("password")
    txt_pword(1).Text = student.Recordset.Fields("password")
    
    MsgBox "Do Not Search or Print whilst editing or data may be lost!" & vbCrLf & "Press cancel or apply to search or print", vbExclamation, "Warning"
    
End Sub

Private Sub mnu_feed_Click()
    Call RunBrowser("http://pcbeef.pc.funpic.org/Forum/index.php?showforum=19", 10, 1)
End Sub

Private Sub mnu_out_Click()
    Unload Me
    Unload main
    login.Show
End Sub

Private Sub mnu_support_Click()
    Call RunBrowser("http://www.users.tpg.com.au/tttran", 10, 1)
    Call RunBrowser("http://pcbeef.pc.funpic.org/Forum/index.php?showforum=22", 10, 1)
End Sub

Private Sub mnu_use_Click()
frmTerms.Show
End Sub

Private Sub mnu_www_Click()
    Call RunBrowser("http://pcbeef.pc.funpic.org/", 10, 1)
End Sub

Private Sub opt_show_Click(Index As Integer)
          
    Select Case Index
        Case Is = 0
            login.absences.RecordSource = "select * from absences where student_id LIKE " & txt_field(6).Text & " order by date"
        Case Is = 1
            login.absences.RecordSource = "select * from absences where student_id LIKE " & txt_field(6).Text & " AND Type = 'partial' order by date"
        Case Is = 2
            login.absences.RecordSource = "select * from absences where student_id LIKE " & txt_field(6).Text & " AND Type = 'Whole Day' order by date"
        Case Is = 3
            login.absences.RecordSource = "select * from absences where student_id LIKE " & txt_field(6).Text & " AND Reason = 'None' order by date"
    End Select
    
    login.absences.Refresh

End Sub

Private Sub toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case Button
        Case "Back"
            mnu_back_Click
        Case "Edit"
            mnu_edit_Click
        Case "Add"
            mnu_add_Click
        Case "Delete"
            delete_student
        Case "Apply"
            apply
        Case "Cancel"
            Cancel
    End Select
    
End Sub

Public Sub organise_grid()

    grid_absence.ColWidth(0) = grid_absence.width * 0.1
    grid_absence.ColWidth(1) = grid_absence.width * 0.1
    grid_absence.ColWidth(2) = grid_absence.width * 0.1
    grid_absence.ColWidth(3) = grid_absence.width * 0.69
    
End Sub

Public Sub delete_student()
       
    'confirm deletion
    If MsgBox("Are you sure you want to delete " & "'" & list_name.Text & "' ?", vbQuestion + vbYesNo, "Are you sure you want to delete?") = vbYes Then
        
        delete_mark
        ' delete an entry from the database
        With student.Recordset
            .Delete
            .Requery
        End With
        
        list_disp
        
        stu_exp.stat.Panels(1).Text = stu_exp.list_name.ListCount & " students found"
            
    End If
    
End Sub

Public Sub bind_txt()
    
    txt_field(0).Text = student.Recordset.Fields("first_name")
    txt_field(1).Text = student.Recordset.Fields("last_name")
    txt_field(2).Text = student.Recordset.Fields("address")
    txt_field(3).Text = student.Recordset.Fields("city")
    txt_field(4).Text = student.Recordset.Fields("postcode")
    txt_field(5).Text = student.Recordset.Fields("home_phone")
    txt_field(6).Text = student.Recordset.Fields("student_id")
    txt_field(7).Text = student.Recordset.Fields("bos")
    cmb_state.Text = student.Recordset.Fields("state")
    dt_date.Value = student.Recordset.Fields("dob")
    
    If student.Recordset.Fields("image") <> "\Images\" Then
        img_pic.Picture = LoadPicture(App.Path & student.Recordset.Fields("image"))
    Else
        img_pic.Picture = LoadPicture(App.Path & "\Images\noimage.gif")
    End If
    
End Sub

Private Sub txt_no_KeyPress(Index As Integer, KeyAscii As Integer)
    
    'integer only textbox
    If (Not IsNumeric(ChrW(KeyAscii)) And Not (KeyAscii = vbKeyBack)) Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub txt_field_KeyPress(Index As Integer, KeyAscii As Integer)
    
    'integer only textbox
    If Index = 4 Or Index = 5 Or Index = 7 Then
        If (Not IsNumeric(ChrW(KeyAscii)) And Not (KeyAscii = vbKeyBack)) Then
            KeyAscii = 0
        End If
    End If
    
End Sub

Public Sub enable_txt()
    
    For i = 0 To 5
        With txt_field(i)
            .Locked = False
            .BorderStyle = 1
        End With
    Next
    
    Command1.Visible = True
    lbl_pword(0).Visible = True
    txt_pword(0).Visible = True
    lbl_pword(1).Visible = True
    txt_pword(1).Visible = True
      
    lbl_date.Visible = False
    lbl_dob.Visible = False
    dt_date.Visible = True
    cmb_state.Visible = True
    
    list_name.Enabled = False
    
    Toolbar.Buttons(2).Enabled = False
    Toolbar.Buttons(3).Enabled = False
    Toolbar.Buttons(4).Enabled = False
    Toolbar.Buttons(5).Enabled = True
    Toolbar.Buttons(6).Enabled = True
    
End Sub

Public Sub disable_txt()
    
    i = 0
    For i = 0 To 7
        With txt_field(i)
            .Locked = True
            .BorderStyle = 0
        End With
    Next
    
    i = 0
    For i = 0 To 1
        lbl_pword(i).Visible = False
        txt_pword(i).Visible = False
        txt_pword(i).Text = ""
    Next
    
    Picture1.Visible = True
    lbl_date.Visible = True
    lbl_dob.Visible = True
    dt_date.Visible = False
    cmb_state.Visible = False
        
    Command1.Visible = False
        
    bind_txt
    
    list_name.Enabled = True
    
    Toolbar.Buttons(2).Enabled = True
    Toolbar.Buttons(3).Enabled = True
    Toolbar.Buttons(4).Enabled = True
    Toolbar.Buttons(5).Enabled = False
    Toolbar.Buttons(6).Enabled = False
    
End Sub


Public Function validate() As Boolean
    
    validate = False
    
    For i = 0 To 7
        If txt_field(i).Text = "" Then
            MsgBox "Not all fields have been filled in, please go back and fill in all available fields.", vbInformation, "Validation Prompt"
            Exit Function
        End If
    Next
    
    If txt_pword(0).Text <> txt_pword(1).Text Then
        MsgBox "Please verify the password you have entered", vbInformation, "Validation Prompt"
        Exit Function
    End If
    
    validate = True
    
End Function

Public Sub apply()

    If add_edit = "Edit" Then
        If MsgBox("Are you sure would like to save these changes?", vbInformation + vbYesNo, "Apply") = vbYes Then
        
            If validate = True Then
                editlog
                edit_student
                disable_txt
            End If
            
        End If
    
    Else
        
        If MsgBox("Are you sure you want to add this student?" & vbCrLf & "Note: BOS number cannot be changed once student is added." & vbCrLf & "See system admin for alteration", vbInformation + vbYesNo, "Apply") = vbYes Then
        
            If validate = True Then
                addlog
                add_student
                disable_txt
            End If
            
        End If
    
    End If

End Sub

Public Sub Cancel()
    If MsgBox("Are you sure you want to cancel? ALL unsaved work will be discarded.", vbInformation + vbOKCancel, "Cancel") = vbOK Then
        disable_txt
    End If
End Sub

Private Sub cmd_print_Click()
    joe3
    JOE
    joe1
End Sub
Private Sub joe1()
    Dim e As Integer
    Dim a As Integer
    Dim Ypos As Single
    Dim c As Integer
    Dim d As Integer
    Dim fnt1 As New StdFont
    Dim fnt As New StdFont
    
    c = list_name.ListCount
    a = 0
    e = 0
    fnt1.Name = "MS Sans Serif": fnt1.Size = 10: fnt1.Bold = False
    
    fnt.Name = "MS Sans Serif": fnt.Size = 10: fnt.Bold = True
    
    If Check3.Value = 1 Then
        list_name.ListIndex = 0
        Set Printer.Font = fnt
        Printer.Print "Student List " & _
            Date + Time
            Printer.CurrentY = Printer.CurrentY + 300
            Printer.CurrentX = 0
            Ypos = Printer.CurrentY
            
            Printer.Print "Full Name"
            Printer.CurrentY = Ypos
            Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth("Student Number")) / 2
            Printer.Print "Student Number"
            Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth("BOS Number"))
            Printer.CurrentY = Ypos
            Printer.Print "BOS Number"
            For a = 0 To c - 1
    
            Printer.CurrentY = Printer.CurrentY + 300
            Printer.CurrentX = 0
            Ypos = Printer.CurrentY
            Printer.Print list_name
            Set Printer.Font = fnt1
            Printer.CurrentY = Ypos
            Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth("Student Number")) / 2
            Printer.Print txt_field(6)
            Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth("BOS Number"))
            Printer.CurrentY = Ypos
            Printer.Print txt_field(7)
            If e = 28 Then
            Printer.NewPage
            Printer.CurrentY = 0
            e = 0
            End If
            
            If list_name.ListIndex < c - 1 Then
                list_name.ListIndex = list_name.ListIndex + 1
            End If
            Set Printer.Font = fnt
            e = e + 1
        Next a
    Printer.EndDoc
    
    End If
    
End Sub

Private Sub joe3()
    Dim dob As String
    Dim a As Integer
    Dim Ypos As Single
    Dim c As Integer
    Dim d As Integer
    c = list_name.ListCount
    a = 0
    Dim fnt1 As New StdFont
    fnt1.Name = "MS Sans Serif": fnt1.Size = 10: fnt1.Bold = False
    Dim fnt As New StdFont
    fnt.Name = "MS Sans Serif": fnt.Size = 10: fnt.Bold = True
    
    If Check4.Value = 1 Then
        list_name.ListIndex = 0
        For a = 0 To c - 1
     Set Printer.Font = fnt
            Printer.Print list_name + "'s Personal Information " & _
            Date + Time
            Printer.CurrentY = 500
            Printer.CurrentX = 0
            Printer.Print "Full Name: "
            Printer.CurrentY = 500
            Printer.CurrentX = 1700
            Set Printer.Font = fnt1
            Printer.Print stu_exp.txt_field(0) + " " + txt_field(1)
            Set Printer.Font = fnt
            
            Printer.PaintPicture img_pic.Picture, 5000, 500
            
            Printer.CurrentY = 1000
            Printer.CurrentX = 0
            Printer.Print "First Name: "
            Printer.CurrentY = 1000
            Printer.CurrentX = 1700
            Set Printer.Font = fnt1
            Printer.Print stu_exp.txt_field(0)
            Set Printer.Font = fnt
        
            Printer.CurrentY = 1500
            Printer.CurrentX = 0
            Printer.Print "Surname: "
            Printer.CurrentY = 1500
            Printer.CurrentX = 1700
            Set Printer.Font = fnt1
            Printer.Print stu_exp.txt_field(1)
            Set Printer.Font = fnt
        
            Printer.CurrentY = 2000
            Printer.CurrentX = 0
            Printer.Print "Date of Birth: "
            Printer.CurrentY = 2000
            Printer.CurrentX = 1700
            Set Printer.Font = fnt1
            Printer.Print dob
            Set Printer.Font = fnt
        
            Printer.CurrentY = 2500
            Printer.CurrentX = 0
            Printer.Print "Address: "
            Printer.CurrentY = 2500
            Printer.CurrentX = 1700
            Set Printer.Font = fnt1
            Printer.Print txt_field(2)
            Set Printer.Font = fnt
            
            Printer.CurrentY = 3000
            Printer.CurrentX = 0
            Printer.Print "Postcode: "
            Printer.CurrentY = 3000
            Printer.CurrentX = 1700
            Set Printer.Font = fnt1
            Printer.Print txt_field(4)
            Set Printer.Font = fnt
        
            Printer.CurrentY = 3500
            Printer.CurrentX = 0
            Printer.Print "City/Town: "
            Printer.CurrentY = 3500
            Printer.CurrentX = 1700
            Set Printer.Font = fnt1
            Printer.Print stu_exp.txt_field(3)
            Set Printer.Font = fnt
        
            Printer.CurrentY = 4000
            Printer.CurrentX = 0
            Printer.Print "State/Territory: "
            Printer.CurrentY = 4000
            Printer.CurrentX = 1700
            Set Printer.Font = fnt1
            Printer.Print stu_exp.cmb_state
            Set Printer.Font = fnt
        
            Printer.CurrentY = 4500
            Printer.CurrentX = 0
            Printer.Print "Student Number: "
            Printer.CurrentY = 4500
            Printer.CurrentX = 1700
            Set Printer.Font = fnt1
            Printer.Print stu_exp.txt_field(6)
            Set Printer.Font = fnt
        
            Printer.CurrentY = 5000
            Printer.CurrentX = 0
            Printer.Print "BOS Number: "
            Printer.CurrentY = 5000
            Printer.CurrentX = 1700
            Set Printer.Font = fnt1
            Printer.Print stu_exp.txt_field(7)
            Set Printer.Font = fnt
        
            Printer.CurrentY = 5500
            Printer.CurrentX = 0
            Printer.Print "Phone Number: "
            Printer.CurrentY = 5500
            Printer.CurrentX = 1700
            Set Printer.Font = fnt1
            Printer.Print stu_exp.txt_field(5)
            Set Printer.Font = fnt
                If list_name.ListIndex < c - 1 Then
                    list_name.ListIndex = list_name.ListIndex + 1
                End If
            Printer.NewPage
            Next a
            Printer.EndDoc
        End If
End Sub

Private Sub JOE()
Dim i As Integer
    Dim dob As String
    dob = stu_exp.dt_date.Value
    If stu_exp.Check2.Value = 1 Then
    Dim fnt1 As New StdFont
    fnt1.Name = "MS Sans Serif": fnt1.Size = 10: fnt1.Bold = False
    Dim fnt As New StdFont
    fnt.Name = "MS Sans Serif": fnt.Size = 10: fnt.Bold = True
    Set Printer.Font = fnt
                       Printer.Print list_name + "'s Personal Information " & _
            Date + Time
            Printer.CurrentY = 500
            Printer.CurrentX = 0
            Printer.Print "Full Name: "
            Printer.CurrentY = 500
            Printer.CurrentX = 1700
            Set Printer.Font = fnt1
            Printer.Print stu_exp.txt_field(0) + " " + txt_field(1)
            Set Printer.Font = fnt
            
            Printer.PaintPicture img_pic.Picture, 5000, 500
            
            Printer.CurrentY = 1000
            Printer.CurrentX = 0
            Printer.Print "First Name: "
            Printer.CurrentY = 1000
            Printer.CurrentX = 1700
            Set Printer.Font = fnt1
            Printer.Print stu_exp.txt_field(0)
            Set Printer.Font = fnt
        
            Printer.CurrentY = 1500
            Printer.CurrentX = 0
            Printer.Print "Surname: "
            Printer.CurrentY = 1500
            Printer.CurrentX = 1700
            Set Printer.Font = fnt1
            Printer.Print stu_exp.txt_field(1)
            Set Printer.Font = fnt
        
            Printer.CurrentY = 2000
            Printer.CurrentX = 0
            Printer.Print "Date of Birth: "
            Printer.CurrentY = 2000
            Printer.CurrentX = 1700
            Set Printer.Font = fnt1
            Printer.Print dob
            Set Printer.Font = fnt
        
            Printer.CurrentY = 2500
            Printer.CurrentX = 0
            Printer.Print "Address: "
            Printer.CurrentY = 2500
            Printer.CurrentX = 1700
            Set Printer.Font = fnt1
            Printer.Print txt_field(2)
            Set Printer.Font = fnt
            
            Printer.CurrentY = 3000
            Printer.CurrentX = 0
            Printer.Print "Postcode: "
            Printer.CurrentY = 3000
            Printer.CurrentX = 1700
            Set Printer.Font = fnt1
            Printer.Print txt_field(4)
            Set Printer.Font = fnt
        
            Printer.CurrentY = 3500
            Printer.CurrentX = 0
            Printer.Print "City/Town: "
            Printer.CurrentY = 3500
            Printer.CurrentX = 1700
            Set Printer.Font = fnt1
            Printer.Print stu_exp.txt_field(3)
            Set Printer.Font = fnt
        
            Printer.CurrentY = 4000
            Printer.CurrentX = 0
            Printer.Print "State/Territory: "
            Printer.CurrentY = 4000
            Printer.CurrentX = 1700
            Set Printer.Font = fnt1
            Printer.Print stu_exp.cmb_state
            Set Printer.Font = fnt
        
            Printer.CurrentY = 4500
            Printer.CurrentX = 0
            Printer.Print "Student Number: "
            Printer.CurrentY = 4500
            Printer.CurrentX = 1700
            Set Printer.Font = fnt1
            Printer.Print stu_exp.txt_field(6)
            Set Printer.Font = fnt
        
            Printer.CurrentY = 5000
            Printer.CurrentX = 0
            Printer.Print "BOS Number: "
            Printer.CurrentY = 5000
            Printer.CurrentX = 1700
            Set Printer.Font = fnt1
            Printer.Print stu_exp.txt_field(7)
            Set Printer.Font = fnt
        
            Printer.CurrentY = 5500
            Printer.CurrentX = 0
            Printer.Print "Phone Number: "
            Printer.CurrentY = 5500
            Printer.CurrentX = 1700
            Set Printer.Font = fnt1
            Printer.Print stu_exp.txt_field(5)
            Set Printer.Font = fnt
    
    'Printer.Print "Surname: " + txt_field(1)
   
    'Printer.Print "Date of Birth: " + dob
    
    'Printer.Print "City/Town: " + txt_field(2)

    'Printer.Print "Postcode: " + txt_field(3)
  
    'Printer.Print "State/Territory: " + cmb_state
  
    'Printer.Print "Student Number: " + lbl_info(0)
   
    'Printer.Print "BOS Number: " + lbl_info(1)
  
    'Printer.Print "Phone Number: " + txt_field(5)
    
    Printer.EndDoc
    End If
    If stu_exp.Check1.Value = 1 Then
        If grid_absence.Text = "" Then
            If MsgBox("Sorry, there are no absences to print", , "Error") Then
                i = i
            End If
        
        Else
                P001.GridToPrint = grid_absence
                P001.Text1(3).FontSize = 10
                P001.Text1(3).FontBold = True
                P001.ReportTitle = list_name + "'s Absences " & _
                    Date + Time

                P001.AllowDialogue = False  'let the user change some of the settings at run time
                P001.TitlePages = TitleOption.USER_MAY_SET + TitleOption.TITLE_ALL_PAGES
                P001.SetEffects = Effects.EFFECTS_NORMAL
                'plus any more property settings you want to control from the calling program
                'then just
                P001.PrintGridAPI   'call the PrintGrid routines to display the P001 dialogue
        End If
        
    End If


End Sub

Public Sub add()
    
    Picture1.Visible = False
    
    enable_txt
    
    With txt_field(7)
        .Locked = False
        .BorderStyle = 1
        .BackColor = &HFFFFFF
    End With
    
    i = 0
    For i = 0 To 7
        txt_field(i).Text = ""
    Next
    
    i = 0
    For i = 0 To 1
        lbl_pword(i).Visible = True
        txt_pword(i).Visible = True
    Next
    
    txt_field(6).Text = "Auto Genereated"
    
End Sub

Public Sub edit_student()
    
    'If File1.Path & "\" & File1.FileName <> App.Path & "\Images\" & File1.FileName Then
    '    FileCopy File1.Path & "\" & File1.FileName, App.Path & "\Images\" & File1.FileName
    'End If
    
    With student.Recordset
        !first_name = txt_field(0)
        !last_name = txt_field(1)
        !address = txt_field(2)
        !city = txt_field(3)
        !postcode = txt_field(4)
        !home_phone = txt_field(5)
        !dob = dt_date.Value
        !State = cmb_state.Text
        !Password = txt_pword(0).Text
        !Image = "\Images\" & File1.FileName
        .Update
        .Requery
    End With
    
End Sub

Public Sub add_student()
    
    If File1.Path & "\" & File1.FileName <> App.Path & "\Images\" & File1.FileName Then
        FileCopy File1.Path & "\" & File1.FileName, App.Path & "\Images\" & File1.FileName
    End If
    With student.Recordset
        .AddNew
        !first_name = txt_field(0)
        !last_name = txt_field(1)
        !address = txt_field(2)
        !city = txt_field(3)
        !postcode = txt_field(4)
        !home_phone = txt_field(5)
        !bos = txt_field(7)
        !dob = dt_date.Value
        !State = cmb_state.Text
        !Password = txt_pword(1).Text
        !Image = "\Images\" & File1.FileName
        .Update
    End With
    
    list_disp
    
    Picture1.Visible = True
    
End Sub

Public Sub deletelog()
    
    'add delete log
    With login.dblog.Recordset
        .AddNew
        !Time = Time
        !Date = Date
        !Event = "Student " & list_name.Text & " was deleted"
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
        !Event = "Student " & list_name.Text & " was edited"
        !User = login.db_login.Recordset.Fields(0)
        .Update
        .Requery
    End With
    
    login.dblog.Refresh
    
End Sub

Public Sub addlog()
    
    'add add log
    With login.dblog.Recordset
        .AddNew
        !Time = Time
        !Date = Date
        !Event = "Student " & txt_field(0).Text & " was added"
        !User = login.db_login.Recordset.Fields(0)
        .Update
        .Requery
    End With
    
    login.dblog.Refresh
    
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub
Private Sub Command1_Click()
    pic_open.Visible = True
    Dir1.Path = App.Path & "\Images\"
    File1.Path = App.Path & "\Images\"
End Sub
Private Sub cmd_cancel_Click()
    pic_open.Visible = False
    Dir1.Path = App.Path & "\Images\"
    File1.Path = App.Path & "\Images\"
    File1.FileName = ""
End Sub

Private Sub cmd_ok_Click()
    img_pic.Picture = LoadPicture(File1.Path & "\" & File1.FileName)
    pic_open.Visible = False
End Sub

Public Sub delete_mark()
    
    dbmark.RecordSource = "select * from mark where student_id = LIKE " & txt_field(6).Text & ""
    dbmark.Refresh
    Do While Not dbmark.Recordset.EOF
        dbmark.Recordset.Delete
        dbmark.Recordset.MoveNext
    Loop
    dbmark.Recordset.Update
    dbmark.Recordset.Requery
    
End Sub
