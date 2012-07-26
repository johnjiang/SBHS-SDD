VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form view 
   BackColor       =   &H00FFFFFF&
   Caption         =   " Fort Minor"
   ClientHeight    =   10410
   ClientLeft      =   1140
   ClientTop       =   1845
   ClientWidth     =   15000
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   10410
   ScaleWidth      =   15000
   Begin VB.TextBox txtimage 
      DataField       =   "Cover"
      DataSource      =   "dvddb1"
      Height          =   330
      Left            =   5160
      TabIndex        =   49
      Top             =   9600
      Visible         =   0   'False
      Width           =   1455
   End
   Begin TabDlg.SSTab dvdtab 
      Height          =   4695
      Left            =   120
      TabIndex        =   15
      Top             =   5280
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   8281
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Browse"
      TabPicture(0)   =   "view.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "datatree"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Search"
      TabPicture(1)   =   "view.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label14"
      Tab(1).Control(1)=   "Label15"
      Tab(1).Control(2)=   "Image1"
      Tab(1).Control(3)=   "Label24"
      Tab(1).Control(4)=   "Label25"
      Tab(1).Control(5)=   "Label26"
      Tab(1).Control(6)=   "Line2"
      Tab(1).Control(7)=   "Label27"
      Tab(1).Control(8)=   "txtsearch"
      Tab(1).Control(9)=   "cmbgen"
      Tab(1).Control(10)=   "cmbclass"
      Tab(1).Control(11)=   "txtrel"
      Tab(1).Control(12)=   "cmdgo"
      Tab(1).Control(13)=   "txtcountry"
      Tab(1).Control(14)=   "txtlang"
      Tab(1).ControlCount=   15
      Begin VB.TextBox txtlang 
         Height          =   375
         Left            =   -73440
         TabIndex        =   55
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox txtcountry 
         Height          =   375
         Left            =   -73440
         TabIndex        =   54
         Top             =   3480
         Width           =   1215
      End
      Begin VB.CommandButton cmdgo 
         Caption         =   "Go"
         Default         =   -1  'True
         Height          =   360
         Left            =   -72360
         TabIndex        =   51
         Top             =   570
         Width           =   480
      End
      Begin VB.TextBox txtrel 
         Height          =   375
         Left            =   -73440
         MaxLength       =   4
         TabIndex        =   50
         Top             =   1560
         Width           =   1215
      End
      Begin ComctlLib.TreeView datatree 
         Height          =   4095
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   7223
         _Version        =   327682
         LabelEdit       =   1
         Style           =   7
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.ComboBox cmbclass 
         Height          =   360
         ItemData        =   "view.frx":0038
         Left            =   -73440
         List            =   "view.frx":0051
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   2520
         Width           =   1215
      End
      Begin VB.ComboBox cmbgen 
         Height          =   360
         ItemData        =   "view.frx":0075
         Left            =   -73440
         List            =   "view.frx":00AC
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox txtsearch 
         Height          =   360
         Left            =   -74450
         MaxLength       =   30
         TabIndex        =   17
         Top             =   570
         Width           =   2130
      End
      Begin VB.Label Label27 
         Caption         =   "Advanced Search"
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
         TabIndex        =   57
         Top             =   1200
         Width           =   2535
      End
      Begin VB.Line Line2 
         X1              =   -74880
         X2              =   -71880
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label26 
         Caption         =   "Language :"
         Height          =   375
         Left            =   -74760
         TabIndex        =   56
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label Label25 
         Caption         =   "Country :"
         Height          =   375
         Left            =   -74760
         TabIndex        =   53
         Top             =   3480
         Width           =   1335
      End
      Begin VB.Label Label24 
         Caption         =   "Release Date :"
         Height          =   375
         Left            =   -74760
         TabIndex        =   52
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Image Image1 
         Height          =   300
         Left            =   -74760
         Picture         =   "view.frx":014C
         Top             =   600
         Width           =   300
      End
      Begin VB.Label Label15 
         Caption         =   "Classification :"
         Height          =   255
         Left            =   -74760
         TabIndex        =   21
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label14 
         Caption         =   "Genre :"
         Height          =   255
         Left            =   -74760
         TabIndex        =   20
         Top             =   2040
         Width           =   855
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dvdgrid1 
      Bindings        =   "view.frx":075F
      Height          =   4695
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   8281
      _Version        =   393216
      Rows            =   16
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      BackColorFixed  =   15984078
      BackColorSel    =   15521715
      BackColorBkg    =   16777215
      BackColorUnpopulated=   16777215
      GridColorFixed  =   0
      GridColorUnpopulated=   0
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   1
      _Band(0).GridLinesBand=   1
      _Band(0).GridLineWidthBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
      _Band(0).ColHeader=   1
      _Band(0)._NumMapCols=   14
      _Band(0)._MapCol(0)._Name=   "ID"
      _Band(0)._MapCol(0)._RSIndex=   0
      _Band(0)._MapCol(0)._Alignment=   7
      _Band(0)._MapCol(0)._Hidden=   -1  'True
      _Band(0)._MapCol(1)._Name=   "Title"
      _Band(0)._MapCol(1)._RSIndex=   1
      _Band(0)._MapCol(2)._Name=   "Release_Date"
      _Band(0)._MapCol(2)._Caption=   "Release Date"
      _Band(0)._MapCol(2)._RSIndex=   2
      _Band(0)._MapCol(2)._Alignment=   7
      _Band(0)._MapCol(2)._Hidden=   -1  'True
      _Band(0)._MapCol(3)._Name=   "Director"
      _Band(0)._MapCol(3)._RSIndex=   3
      _Band(0)._MapCol(3)._Hidden=   -1  'True
      _Band(0)._MapCol(4)._Name=   "Producer"
      _Band(0)._MapCol(4)._RSIndex=   4
      _Band(0)._MapCol(4)._Hidden=   -1  'True
      _Band(0)._MapCol(5)._Name=   "Writer"
      _Band(0)._MapCol(5)._RSIndex=   5
      _Band(0)._MapCol(5)._Hidden=   -1  'True
      _Band(0)._MapCol(6)._Name=   "Genre"
      _Band(0)._MapCol(6)._RSIndex=   6
      _Band(0)._MapCol(6)._Hidden=   -1  'True
      _Band(0)._MapCol(7)._Name=   "Classification"
      _Band(0)._MapCol(7)._RSIndex=   7
      _Band(0)._MapCol(7)._Hidden=   -1  'True
      _Band(0)._MapCol(8)._Name=   "Runtime"
      _Band(0)._MapCol(8)._RSIndex=   8
      _Band(0)._MapCol(8)._Hidden=   -1  'True
      _Band(0)._MapCol(9)._Name=   "Country"
      _Band(0)._MapCol(9)._RSIndex=   9
      _Band(0)._MapCol(9)._Hidden=   -1  'True
      _Band(0)._MapCol(10)._Name=   "Language"
      _Band(0)._MapCol(10)._RSIndex=   10
      _Band(0)._MapCol(10)._Hidden=   -1  'True
      _Band(0)._MapCol(11)._Name=   "Awards"
      _Band(0)._MapCol(11)._RSIndex=   11
      _Band(0)._MapCol(11)._Hidden=   -1  'True
      _Band(0)._MapCol(12)._Name=   "Official_Site"
      _Band(0)._MapCol(12)._RSIndex=   12
      _Band(0)._MapCol(12)._Hidden=   -1  'True
      _Band(0)._MapCol(13)._Name=   "Plot"
      _Band(0)._MapCol(13)._RSIndex=   13
      _Band(0)._MapCol(13)._Hidden=   -1  'True
   End
   Begin ComctlLib.StatusBar dvdstatus 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   10050
      Width           =   15000
      _ExtentX        =   26458
      _ExtentY        =   635
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   4410
            MinWidth        =   4410
            Text            =   "Logged in as "
            TextSave        =   "Logged in as "
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   3351
            MinWidth        =   3351
            Text            =   "Rank:"
            TextSave        =   "Rank:"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   2646
            MinWidth        =   2646
            Text            =   "Total DVDs:"
            TextSave        =   "Total DVDs:"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   2822
            MinWidth        =   2822
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc dvddb1 
      Height          =   330
      Left            =   3840
      Top             =   9600
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
      RecordSource    =   "Select * FROM DVDs ORDER by Title"
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
   Begin VB.Label txtfield 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "15"
      Height          =   240
      Index           =   15
      Left            =   12840
      TabIndex        =   48
      Top             =   3000
      Width           =   1995
      WordWrap        =   -1  'True
   End
   Begin VB.Label txtfield 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "14"
      Height          =   240
      Index           =   14
      Left            =   12840
      TabIndex        =   47
      Top             =   2400
      Width           =   1875
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbladd 
      BackStyle       =   0  'Transparent
      Caption         =   "Main Audio :"
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
      Index           =   12
      Left            =   11280
      TabIndex        =   46
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label lbladd 
      BackStyle       =   0  'Transparent
      Caption         =   "Screen Ratio :"
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
      Index           =   13
      Left            =   11280
      TabIndex        =   45
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "Rating :"
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
      Left            =   3720
      TabIndex        =   44
      Top             =   5760
      Width           =   855
   End
   Begin VB.Label txtfield 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "16"
      Height          =   240
      Index           =   16
      Left            =   12840
      TabIndex        =   43
      Top             =   3600
      Width           =   675
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "No of Disks :"
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
      Left            =   11280
      TabIndex        =   42
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label txtfield 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "12"
      Height          =   240
      Index           =   12
      Left            =   8880
      TabIndex        =   41
      Top             =   2400
      Width           =   2235
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Studio :"
      DataSource      =   "dvddb1"
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
      Left            =   7320
      TabIndex        =   40
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Miscellaneous :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   39
      Top             =   5280
      Width           =   2415
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Details :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11280
      TabIndex        =   38
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Browser :"
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
      Left            =   120
      TabIndex        =   37
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "DVD Cover"
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
      Left            =   3720
      TabIndex        =   36
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label txtfield 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      Height          =   3855
      Index           =   10
      Left            =   7320
      TabIndex        =   35
      Top             =   5760
      Width           =   7335
   End
   Begin VB.Label txtfield 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "9"
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   9
      Left            =   5280
      TabIndex        =   34
      Top             =   7160
      Width           =   1035
      WordWrap        =   -1  'True
   End
   Begin VB.Label txtfield 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      Height          =   480
      Index           =   8
      Left            =   5280
      TabIndex        =   33
      Top             =   6460
      Width           =   1635
      WordWrap        =   -1  'True
   End
   Begin VB.Label txtfield 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      Height          =   240
      Index           =   7
      Left            =   12840
      TabIndex        =   32
      Top             =   1800
      Width           =   675
      WordWrap        =   -1  'True
   End
   Begin VB.Label txtfield 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      Height          =   240
      Index           =   6
      Left            =   12840
      TabIndex        =   31
      Top             =   4200
      Width           =   675
      WordWrap        =   -1  'True
   End
   Begin VB.Label txtfield 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      Height          =   240
      Index           =   5
      Left            =   12840
      TabIndex        =   30
      Top             =   1200
      Width           =   675
      WordWrap        =   -1  'True
   End
   Begin VB.Label txtfield 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "13"
      Height          =   240
      Index           =   13
      Left            =   12840
      TabIndex        =   29
      Top             =   600
      Width           =   675
      WordWrap        =   -1  'True
   End
   Begin VB.Label txtfield 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "11"
      Height          =   240
      Index           =   11
      Left            =   8880
      TabIndex        =   28
      Top             =   1800
      Width           =   2235
      WordWrap        =   -1  'True
   End
   Begin VB.Label txtfield 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      Height          =   240
      Index           =   4
      Left            =   8880
      TabIndex        =   27
      Top             =   4200
      Width           =   2235
      WordWrap        =   -1  'True
   End
   Begin VB.Label txtfield 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      Height          =   240
      Index           =   3
      Left            =   8880
      TabIndex        =   26
      Top             =   3600
      Width           =   2235
      WordWrap        =   -1  'True
   End
   Begin VB.Label txtfield 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      Height          =   240
      Index           =   2
      Left            =   8880
      TabIndex        =   25
      Top             =   3000
      Width           =   2235
      WordWrap        =   -1  'True
   End
   Begin VB.Label txtfield 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   240
      Index           =   1
      Left            =   8880
      TabIndex        =   24
      Top             =   1200
      Width           =   2235
      WordWrap        =   -1  'True
   End
   Begin VB.Label txtfield 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   240
      Index           =   0
      Left            =   8880
      TabIndex        =   23
      Top             =   600
      Width           =   2235
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Movie Details :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   22
      Top             =   120
      Width           =   2415
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   14760
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Synoposis :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   14
      Top             =   5280
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Director :"
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
      Left            =   7320
      TabIndex        =   7
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Producer :"
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
      Left            =   7320
      TabIndex        =   6
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Writer :"
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
      Left            =   7320
      TabIndex        =   5
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Image imgcover 
      BorderStyle     =   1  'Fixed Single
      DataSource      =   "dvddb1"
      Height          =   4695
      Left            =   3720
      Stretch         =   -1  'True
      Top             =   360
      Width           =   3255
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Official Site :"
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
      Left            =   3720
      TabIndex        =   13
      Top             =   7160
      Width           =   1335
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Awards :"
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
      Left            =   3720
      TabIndex        =   12
      Top             =   6460
      Width           =   1215
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Runtime :"
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
      Left            =   11280
      TabIndex        =   11
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Country :"
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
      Left            =   11280
      TabIndex        =   10
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Language :"
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
      Left            =   11280
      TabIndex        =   9
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Genre :"
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
      Left            =   7320
      TabIndex        =   8
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Classification :"
      DataSource      =   "dvddb1"
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
      Left            =   11280
      TabIndex        =   4
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Release Date :"
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
      Left            =   7320
      TabIndex        =   3
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Title :"
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
      Left            =   7320
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnuprint 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuout 
         Caption         =   "Log &Out"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuexit 
         Caption         =   "&Exit"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnutools 
      Caption         =   "&Tools"
      Begin VB.Menu mnu_db 
         Caption         =   "Manage &Database"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuusers 
         Caption         =   "Manage &Users"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnu_log 
         Caption         =   "View &Log"
         Shortcut        =   ^L
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuabout 
         Caption         =   "&About"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnucontext 
      Caption         =   "Context Menu"
      Visible         =   0   'False
      Begin VB.Menu mnucopy 
         Caption         =   "Copy"
      End
   End
End
Attribute VB_Name = "view"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim selmenu As Integer
Option Explicit


Private Sub Form_Load()
    
    'calls on individual functions
    set_grid
    
    dvdcount
    
    treeview
    
    dvdgrid1.RowSel = 1
    
    dvdgrid1_RowColChange
    
    load_imageview
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    'confirm user exit
    If MsgBox("Are you sure you want to Exit?", vbQuestion + vbYesNo, "Exit") = vbNo Then
        Cancel = True
    Else
        End
    End If
    
End Sub

Private Sub dvdgrid1_RowColChange()
    
    'fetches info depending on clicked row
    If dvdgrid1.Row > 0 Then
        dvddb1.Recordset.AbsolutePosition = dvdgrid1.Row
    End If
    
    load_imageview
    
    'binds fields to db
    For i = 0 To 16
        txtfield(i) = dvddb1.Recordset.Fields(i)
    Next
    txtfield(9).Caption = "Click Here"
    
End Sub


Private Sub dvdgrid1_Click()
    dvdgrid1_RowColChange
End Sub

Private Sub dvdgrid1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    ' If this is not row 0, do nothing.
    If dvdgrid1.MouseRow <> 0 Then Exit Sub

    ' Sort by the clicked column.
    If dvddb1.Recordset.Sort = dvddb1.Recordset.Fields(dvdgrid1.MouseCol).Name & " ASC" Then
        dvddb1.Recordset.Sort = dvddb1.Recordset.Fields(dvdgrid1.MouseCol).Name & " DESC"
    Else
        dvddb1.Recordset.Sort = dvddb1.Recordset.Fields(dvdgrid1.MouseCol).Name & " ASC"
    End If
    
End Sub

Private Sub mnuout_Click()
    
    'confirm user log out
    If MsgBox("Are you sure you want to Log Out?", vbQuestion + vbYesNo, "Exit") = vbYes Then
        Me.Hide
        login.Show
    End If
    
End Sub

Private Sub txtfield_Click(Index As Integer)
    
    'opens browser if url field is clicked
    If Index = 9 Then
        Call RunBrowser(view.dvddb1.Recordset.Fields(9), 10, 1)
    End If
End Sub

Private Sub txtfield_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'if mouse is over url field, then the tooltip is the url
    If Index = 9 Then
        txtfield(9).ToolTipText = view.dvddb1.Recordset.Fields(9)
    End If
End Sub

Private Sub txtrel_KeyPress(KeyAscii As Integer)
    'text only textbox
    If (Not IsNumeric(ChrW(KeyAscii)) And Not (KeyAscii = vbKeyBack)) Then
        KeyAscii = 0
    End If
End Sub

Private Sub cmdgo_Click()
    dvdstatus.Panels(4).Text = "Searching..."
    
    'searches the db
    dvddb1.RecordSource = "select * from DVDs where title LIKE '%" & txtsearch.Text & "%' AND release_date Like '%" & txtrel.Text & "%' AND genre Like '%" & cmbgen.List(cmbgen.ListIndex) & "%' AND Classification Like '%" & cmbclass.List(cmbclass.ListIndex) & "%' AND Country LIKE '%" & txtcountry.Text & "%' AND Lang LIKE '%" & txtlang.Text & "%' order by Title"
    
    Set_datasource
    
    status_search
End Sub

Public Sub treeview()
    
    ' adds the treelists
    datatree.Nodes.Add , , "root", "Genre"
    datatree.Nodes.Add "root", tvwChild, "child1_root", "All"
    datatree.Nodes.Add "root", tvwChild, "child2_root", "Action"
    datatree.Nodes.Add "root", tvwChild, "child3_root", "Adult"
    datatree.Nodes.Add "root", tvwChild, "child4_root", "Adventure"
    datatree.Nodes.Add "root", tvwChild, "child5_root", "Anime"
    datatree.Nodes.Add "root", tvwChild, "child6_root", "Children"
    datatree.Nodes.Add "root", tvwChild, "child7_root", "Comedy"
    datatree.Nodes.Add "root", tvwChild, "child8_root", "Documentary"
    datatree.Nodes.Add "root", tvwChild, "child9_root", "Drama"
    datatree.Nodes.Add "root", tvwChild, "child10_root", "Family"
    datatree.Nodes.Add "root", tvwChild, "child11_root", "Horror"
    datatree.Nodes.Add "root", tvwChild, "child12_root", "Music"
    datatree.Nodes.Add "root", tvwChild, "child13_root", "Musical"
    datatree.Nodes.Add "root", tvwChild, "child14_root", "Mystery"
    datatree.Nodes.Add "root", tvwChild, "child15_root", "Romance"
    datatree.Nodes.Add "root", tvwChild, "child16_root", "Science Fiction"
    datatree.Nodes.Add "root", tvwChild, "child17_root", "Sports"
    datatree.Nodes.Add "root", tvwChild, "child18_root", "Thriller"
    
    ' make the lists visible
    datatree.Nodes("child2_root").EnsureVisible
    
End Sub

Private Sub datatree_Click()
    
    dvdstatus.Panels(4).Text = "Searching..."
    
    'sorts db depending on clicked item
    Select Case datatree.SelectedItem
        Case "Genre"
            Exit Sub
        Case "All"
            dvddb1.RecordSource = "select * from DVDs order by Title"
        Case Else
            dvddb1.RecordSource = "select * from DVDs where genre LIKE '%" & datatree.SelectedItem & "%' order by Title"
    End Select

    Set_datasource
    
    status_search
    
End Sub

Public Sub status_search()
    
    'displays info in status bar
    If dvdgrid1.Rows = 0 Then
        dvdstatus.Panels(4).Text = "No DVDs Found"
    Else
        dvdstatus.Panels(4).Text = dvdgrid1.Rows - 1 & " DVDs Found"
    End If
End Sub

Private Sub txtfield_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    selmenu = Index
    
    'displays context menu on right click
    If Button = vbRightButton Then
        PopupMenu mnucontext, vbPopupMenuRightButton
    End If
    
End Sub

Private Sub mnucopy_Click()
    
    'copy's info into clipboard
    Clipboard.Clear
    Clipboard.SetText txtfield(selmenu).Caption
    
End Sub

Private Sub mnuexit_Click()
    Unload Me
End Sub

Private Sub mnu_db_Click()
    
    'validates rights to access manage database form
    If dvdstatus.Panels(2).Text = "Rank: User" Then
        MsgBox "You do not have priviledges to access this", vbInformation
    Else
        Me.Enabled = False
        addentry.Show
    End If
    
End Sub

Private Sub mnuusers_Click()
    
    'validates rights to access manage users form
    If dvdstatus.Panels(2).Text = "Rank: Administrator" Then
        Me.Enabled = False
        users.Show
    Else
        MsgBox "You do not have priviledges to access this", vbInformation
    End If
End Sub

Private Sub mnu_log_Click()
    
    'logs out user
    Me.Enabled = False
    log.Show
End Sub

Private Sub mnuprint_Click()
    If MsgBox("Are you sure you want to Print?", vbQuestion + vbYesNo, "Print") = vbYes Then
        Me.PrintForm
    End If
End Sub
