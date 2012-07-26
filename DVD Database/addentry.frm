VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form addentry 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manage Database"
   ClientHeight    =   6825
   ClientLeft      =   2835
   ClientTop       =   3045
   ClientWidth     =   11310
   ClipControls    =   0   'False
   Icon            =   "addentry.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   11310
   Begin ComctlLib.Toolbar barDVD 
      Align           =   1  'Align Top
      Height          =   675
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11310
      _ExtentX        =   19950
      _ExtentY        =   1191
      ButtonWidth     =   1852
      ButtonHeight    =   1032
      Appearance      =   1
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   4
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Go Back"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Add Entry"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Edit Entry"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Delete Entry"
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dvdgrid1 
      Bindings        =   "addentry.frx":000C
      Height          =   6015
      Left            =   0
      TabIndex        =   5
      Top             =   720
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   10610
      _Version        =   393216
      Rows            =   11
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      BackColorFixed  =   15984078
      BackColorSel    =   14398577
      BackColorBkg    =   16777215
      BackColorUnpopulated=   16777215
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      RowSizingMode   =   1
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
   Begin VB.Frame frmadd 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Add Entry"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6135
      Left            =   3120
      TabIndex        =   1
      Top             =   600
      Width           =   8055
      Begin VB.CommandButton cmdclear 
         Caption         =   "&Clear All Fields"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   5640
         Width           =   1575
      End
      Begin VB.CommandButton cmdadd 
         Caption         =   "&Add Entry"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   5160
         Width           =   1575
      End
      Begin TabDlg.SSTab addtab 
         Height          =   4695
         Left            =   120
         TabIndex        =   30
         Top             =   360
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   8281
         _Version        =   393216
         Tabs            =   5
         TabsPerRow      =   5
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
         TabCaption(0)   =   "Movie Details"
         TabPicture(0)   =   "addentry.frx":0021
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lbladd(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lbladd(1)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "lblcmb(0)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "lbladd(2)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "lbladd(3)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "lbladd(4)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "lblcmb(1)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "txtadd(0)"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "txtadd(1)"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "txtadd(2)"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "txtadd(3)"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "txtadd(4)"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "cmbadd(0)"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "cmbadd(1)"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).ControlCount=   14
         TabCaption(1)   =   "Product Details"
         TabPicture(1)   =   "addentry.frx":003D
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "lbladd(5)"
         Tab(1).Control(1)=   "lblcmb(2)"
         Tab(1).Control(2)=   "lblcmb(5)"
         Tab(1).Control(3)=   "lblcmb(3)"
         Tab(1).Control(4)=   "lblcmb(4)"
         Tab(1).Control(5)=   "lbladd(7)"
         Tab(1).Control(6)=   "lbladd(6)"
         Tab(1).Control(7)=   "txtadd(5)"
         Tab(1).Control(8)=   "cmbadd(2)"
         Tab(1).Control(9)=   "txtadd(7)"
         Tab(1).Control(10)=   "txtadd(6)"
         Tab(1).Control(11)=   "cmbadd(3)"
         Tab(1).Control(12)=   "cmbadd(4)"
         Tab(1).Control(13)=   "cmbadd(5)"
         Tab(1).ControlCount=   14
         TabCaption(2)   =   "Misc."
         TabPicture(2)   =   "addentry.frx":0059
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "lbladd(9)"
         Tab(2).Control(1)=   "lbladd(8)"
         Tab(2).Control(2)=   "lbladd(14)"
         Tab(2).Control(3)=   "txtadd(8)"
         Tab(2).Control(4)=   "txtadd(9)"
         Tab(2).ControlCount=   5
         TabCaption(3)   =   "Plot"
         TabPicture(3)   =   "addentry.frx":0075
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "lbladd(10)"
         Tab(3).Control(1)=   "txtadd(10)"
         Tab(3).ControlCount=   2
         TabCaption(4)   =   "Cover"
         TabPicture(4)   =   "addentry.frx":0091
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "cmdbrowseadd"
         Tab(4).Control(1)=   "txtpicpathadd"
         Tab(4).Control(2)=   "Label1"
         Tab(4).Control(3)=   "imgaddcover"
         Tab(4).ControlCount=   4
         Begin VB.CommandButton cmdbrowseadd 
            Caption         =   "Browse"
            Height          =   375
            Left            =   -68640
            TabIndex        =   81
            Top             =   840
            Width           =   1095
         End
         Begin VB.TextBox txtpicpathadd 
            Height          =   375
            Left            =   -72720
            Locked          =   -1  'True
            TabIndex        =   80
            Top             =   840
            Width           =   3975
         End
         Begin VB.ComboBox cmbadd 
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
            Index           =   5
            ItemData        =   "addentry.frx":00AD
            Left            =   -73200
            List            =   "addentry.frx":00ED
            Style           =   2  'Dropdown List
            TabIndex        =   78
            Top             =   1680
            Width           =   735
         End
         Begin VB.ComboBox cmbadd 
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
            Index           =   4
            ItemData        =   "addentry.frx":0138
            Left            =   -73200
            List            =   "addentry.frx":015A
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   2640
            Width           =   2295
         End
         Begin VB.ComboBox cmbadd 
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
            Index           =   3
            ItemData        =   "addentry.frx":0219
            Left            =   -73200
            List            =   "addentry.frx":022F
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   2160
            Width           =   2295
         End
         Begin VB.TextBox txtadd 
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
            Index           =   6
            Left            =   -73200
            TabIndex        =   36
            Top             =   3600
            Width           =   1695
         End
         Begin VB.TextBox txtadd 
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
            Index           =   7
            Left            =   -73200
            TabIndex        =   35
            Top             =   3120
            Width           =   1695
         End
         Begin VB.TextBox txtadd 
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
            Index           =   9
            Left            =   -73320
            TabIndex        =   59
            Top             =   1680
            Width           =   2895
         End
         Begin VB.TextBox txtadd 
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
            Index           =   8
            Left            =   -73320
            TabIndex        =   58
            Top             =   1200
            Width           =   2895
         End
         Begin VB.TextBox txtadd 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3615
            Index           =   10
            Left            =   -74760
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   56
            Top             =   840
            Width           =   7455
         End
         Begin VB.ComboBox cmbadd 
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
            Index           =   1
            ItemData        =   "addentry.frx":029E
            Left            =   1680
            List            =   "addentry.frx":02D2
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   2160
            Width           =   1695
         End
         Begin VB.ComboBox cmbadd 
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
            Index           =   2
            ItemData        =   "addentry.frx":03CD
            Left            =   -73200
            List            =   "addentry.frx":03E6
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   720
            Width           =   855
         End
         Begin VB.ComboBox cmbadd 
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
            ItemData        =   "addentry.frx":040A
            Left            =   1680
            List            =   "addentry.frx":0441
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   1680
            Width           =   1695
         End
         Begin VB.TextBox txtadd 
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
            Index           =   4
            Left            =   1680
            TabIndex        =   43
            Top             =   3600
            Width           =   2895
         End
         Begin VB.TextBox txtadd 
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
            Index           =   3
            Left            =   1680
            TabIndex        =   42
            Top             =   3120
            Width           =   2895
         End
         Begin VB.TextBox txtadd 
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
            Index           =   2
            Left            =   1680
            TabIndex        =   41
            Top             =   2640
            Width           =   2895
         End
         Begin VB.TextBox txtadd 
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
            Index           =   1
            Left            =   1680
            MaxLength       =   4
            TabIndex        =   38
            Top             =   1200
            Width           =   1095
         End
         Begin VB.TextBox txtadd 
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
            Left            =   1680
            TabIndex        =   37
            Top             =   720
            Width           =   2895
         End
         Begin VB.TextBox txtadd 
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
            Index           =   5
            Left            =   -73200
            MaxLength       =   3
            TabIndex        =   32
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Picture Path :"
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
            Left            =   -72720
            TabIndex        =   79
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label lbladd 
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
            Index           =   14
            Left            =   -74760
            TabIndex        =   75
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label lbladd 
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
            Index           =   6
            Left            =   -74760
            TabIndex        =   65
            Top             =   3600
            Width           =   1215
         End
         Begin VB.Label lbladd 
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
            Index           =   7
            Left            =   -74760
            TabIndex        =   64
            Top             =   3120
            Width           =   1215
         End
         Begin VB.Label lblcmb 
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
            Index           =   4
            Left            =   -74760
            TabIndex        =   63
            Top             =   2640
            Width           =   1455
         End
         Begin VB.Label lblcmb 
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
            Index           =   3
            Left            =   -74760
            TabIndex        =   62
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label lbladd 
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
            Index           =   8
            Left            =   -74760
            TabIndex        =   61
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label lbladd 
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
            Index           =   9
            Left            =   -74760
            TabIndex        =   60
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Image imgaddcover 
            BorderStyle     =   1  'Fixed Single
            Height          =   2835
            Left            =   -74880
            Stretch         =   -1  'True
            Top             =   480
            Width           =   1995
         End
         Begin VB.Label lbladd 
            BackStyle       =   0  'Transparent
            Caption         =   "Plot :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   10
            Left            =   -74760
            TabIndex        =   57
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label lblcmb 
            BackStyle       =   0  'Transparent
            Caption         =   "No. of Disks :"
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
            Index           =   5
            Left            =   -74760
            TabIndex        =   54
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label lblcmb 
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
            Index           =   1
            Left            =   240
            TabIndex        =   53
            Top             =   2160
            Width           =   975
         End
         Begin VB.Label lbladd 
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
            Index           =   4
            Left            =   240
            TabIndex        =   51
            Top             =   3600
            Width           =   855
         End
         Begin VB.Label lbladd 
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
            Index           =   3
            Left            =   240
            TabIndex        =   50
            Top             =   3120
            Width           =   1095
         End
         Begin VB.Label lbladd 
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
            Index           =   2
            Left            =   240
            TabIndex        =   49
            Top             =   2640
            Width           =   975
         End
         Begin VB.Label lblcmb 
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
            Index           =   0
            Left            =   240
            TabIndex        =   48
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label lbladd 
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
            Index           =   1
            Left            =   240
            TabIndex        =   47
            Top             =   1200
            Width           =   1455
         End
         Begin VB.Label lbladd 
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
            Index           =   0
            Left            =   240
            TabIndex        =   46
            Top             =   720
            Width           =   855
         End
         Begin VB.Label lblcmb 
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
            Index           =   2
            Left            =   -74760
            TabIndex        =   45
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label lbladd 
            BackStyle       =   0  'Transparent
            Caption         =   "Runtime (Min) :"
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
            Index           =   5
            Left            =   -74760
            TabIndex        =   44
            Top             =   1200
            Width           =   1455
         End
      End
   End
   Begin VB.Frame frmedit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Edit Entry"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6135
      Left            =   3120
      TabIndex        =   4
      Top             =   600
      Visible         =   0   'False
      Width           =   8055
      Begin VB.CommandButton cmdapply 
         Caption         =   "&Apply"
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   5160
         Width           =   1575
      End
      Begin TabDlg.SSTab edittab 
         Height          =   4695
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   8281
         _Version        =   393216
         Tabs            =   5
         TabsPerRow      =   5
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
         TabCaption(0)   =   "Movie Details"
         TabPicture(0)   =   "addentry.frx":04E1
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lblfield(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lblfield(1)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "lbledit(0)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "lblfield(4)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "lblfield(3)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "lblfield(2)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "lbledit(1)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "txtfield(0)"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "txtfield(1)"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "txtfield(2)"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "txtfield(3)"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "txtfield(4)"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "cmbedit(0)"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "cmbedit(1)"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).ControlCount=   14
         TabCaption(1)   =   "Product Details"
         TabPicture(1)   =   "addentry.frx":04FD
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "lbledit(2)"
         Tab(1).Control(1)=   "lblfield(7)"
         Tab(1).Control(2)=   "lblfield(6)"
         Tab(1).Control(3)=   "lblfield(5)"
         Tab(1).Control(4)=   "lblfield(11)"
         Tab(1).Control(5)=   "lbledit(3)"
         Tab(1).Control(6)=   "lbledit(4)"
         Tab(1).Control(7)=   "txtfield(5)"
         Tab(1).Control(8)=   "txtfield(6)"
         Tab(1).Control(9)=   "txtfield(7)"
         Tab(1).Control(10)=   "cmbedit(2)"
         Tab(1).Control(11)=   "cmbedit(3)"
         Tab(1).Control(12)=   "cmbedit(4)"
         Tab(1).Control(13)=   "cmbedit(5)"
         Tab(1).ControlCount=   14
         TabCaption(2)   =   "Misc."
         TabPicture(2)   =   "addentry.frx":0519
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "txtfield(9)"
         Tab(2).Control(1)=   "txtfield(8)"
         Tab(2).Control(2)=   "lblfield(14)"
         Tab(2).Control(3)=   "lblfield(9)"
         Tab(2).Control(4)=   "lblfield(8)"
         Tab(2).ControlCount=   5
         TabCaption(3)   =   "Plot"
         TabPicture(3)   =   "addentry.frx":0535
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "txtfield(10)"
         Tab(3).Control(1)=   "lblfield(10)"
         Tab(3).ControlCount=   2
         TabCaption(4)   =   "Cover"
         TabPicture(4)   =   "addentry.frx":0551
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "imgeditcover"
         Tab(4).Control(1)=   "Label2"
         Tab(4).Control(2)=   "cmdbrowse"
         Tab(4).Control(3)=   "txtpicpath"
         Tab(4).ControlCount=   4
         Begin VB.TextBox txtpicpath 
            Height          =   375
            Left            =   -72720
            TabIndex        =   83
            Top             =   840
            Width           =   3975
         End
         Begin VB.CommandButton cmdbrowse 
            Caption         =   "Browse"
            Height          =   375
            Left            =   -68640
            TabIndex        =   82
            Top             =   840
            Width           =   1095
         End
         Begin VB.ComboBox cmbedit 
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
            Index           =   5
            ItemData        =   "addentry.frx":056D
            Left            =   -73200
            List            =   "addentry.frx":05AD
            Style           =   2  'Dropdown List
            TabIndex        =   77
            Top             =   1680
            Width           =   735
         End
         Begin VB.ComboBox cmbedit 
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
            Index           =   1
            ItemData        =   "addentry.frx":05F8
            Left            =   1680
            List            =   "addentry.frx":062C
            Style           =   2  'Dropdown List
            TabIndex        =   76
            Top             =   2160
            Width           =   1695
         End
         Begin VB.ComboBox cmbedit 
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
            Index           =   4
            ItemData        =   "addentry.frx":0727
            Left            =   -73200
            List            =   "addentry.frx":0749
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   2640
            Width           =   2295
         End
         Begin VB.ComboBox cmbedit 
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
            Index           =   3
            ItemData        =   "addentry.frx":0808
            Left            =   -73200
            List            =   "addentry.frx":081E
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   2160
            Width           =   2295
         End
         Begin VB.TextBox txtfield 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3735
            Index           =   10
            Left            =   -74760
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   70
            Top             =   840
            Width           =   7455
         End
         Begin VB.TextBox txtfield 
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
            Index           =   9
            Left            =   -73320
            TabIndex        =   67
            Top             =   1680
            Width           =   2535
         End
         Begin VB.TextBox txtfield 
            DataField       =   "Awards"
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
            Index           =   8
            Left            =   -73320
            TabIndex        =   66
            Top             =   1200
            Width           =   2535
         End
         Begin VB.ComboBox cmbedit 
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
            Index           =   2
            ItemData        =   "addentry.frx":088D
            Left            =   -73200
            List            =   "addentry.frx":08A6
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   720
            Width           =   855
         End
         Begin VB.ComboBox cmbedit 
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
            ItemData        =   "addentry.frx":08CA
            Left            =   1680
            List            =   "addentry.frx":0901
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   1680
            Width           =   1695
         End
         Begin VB.TextBox txtfield 
            DataField       =   "Writer"
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
            Index           =   4
            Left            =   1680
            MultiLine       =   -1  'True
            TabIndex        =   18
            Top             =   3600
            Width           =   2895
         End
         Begin VB.TextBox txtfield 
            DataField       =   "Producer"
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
            Index           =   3
            Left            =   1680
            MultiLine       =   -1  'True
            TabIndex        =   17
            Top             =   3120
            Width           =   2895
         End
         Begin VB.TextBox txtfield 
            DataField       =   "Director"
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
            Index           =   2
            Left            =   1680
            MultiLine       =   -1  'True
            TabIndex        =   16
            Top             =   2640
            Width           =   2895
         End
         Begin VB.TextBox txtfield 
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
            Index           =   1
            Left            =   1680
            MaxLength       =   4
            TabIndex        =   14
            Top             =   1200
            Width           =   1095
         End
         Begin VB.TextBox txtfield 
            DataField       =   "Title"
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
            Left            =   1680
            TabIndex        =   13
            Top             =   720
            Width           =   2895
         End
         Begin VB.TextBox txtfield 
            DataField       =   "Language"
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
            Index           =   7
            Left            =   -73200
            TabIndex        =   11
            Top             =   3120
            Width           =   1695
         End
         Begin VB.TextBox txtfield 
            DataField       =   "Country"
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
            Index           =   6
            Left            =   -73200
            TabIndex        =   12
            Top             =   3600
            Width           =   1695
         End
         Begin VB.TextBox txtfield 
            DataField       =   "Runtime"
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
            Index           =   5
            Left            =   -73200
            MaxLength       =   4
            TabIndex        =   8
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "Picture Path :"
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
            Left            =   -72720
            TabIndex        =   84
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label lblfield 
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
            Index           =   14
            Left            =   -74760
            TabIndex        =   74
            Top             =   720
            Width           =   855
         End
         Begin VB.Label lbledit 
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
            Index           =   4
            Left            =   -74760
            TabIndex        =   73
            Top             =   2640
            Width           =   1455
         End
         Begin VB.Label lbledit 
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
            Index           =   3
            Left            =   -74760
            TabIndex        =   72
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label lblfield 
            BackStyle       =   0  'Transparent
            Caption         =   "Plot :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   10
            Left            =   -74760
            TabIndex        =   71
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label lblfield 
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
            Index           =   9
            Left            =   -74760
            TabIndex        =   69
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label lblfield 
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
            Index           =   8
            Left            =   -74760
            TabIndex        =   68
            Top             =   1200
            Width           =   855
         End
         Begin VB.Image imgeditcover 
            BorderStyle     =   1  'Fixed Single
            Height          =   2835
            Left            =   -74880
            Stretch         =   -1  'True
            Top             =   480
            Width           =   1995
         End
         Begin VB.Label lblfield 
            BackStyle       =   0  'Transparent
            Caption         =   "No. of Disks :"
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
            Index           =   11
            Left            =   -74760
            TabIndex        =   55
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label lbledit 
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
            Index           =   1
            Left            =   240
            TabIndex        =   52
            Top             =   2160
            Width           =   975
         End
         Begin VB.Label lblfield 
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
            Index           =   2
            Left            =   240
            TabIndex        =   28
            Top             =   2640
            Width           =   1215
         End
         Begin VB.Label lblfield 
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
            Index           =   3
            Left            =   240
            TabIndex        =   27
            Top             =   3120
            Width           =   1215
         End
         Begin VB.Label lblfield 
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
            Index           =   4
            Left            =   240
            TabIndex        =   26
            Top             =   3600
            Width           =   735
         End
         Begin VB.Label lbledit 
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
            Index           =   0
            Left            =   240
            TabIndex        =   25
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label lblfield 
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
            Index           =   1
            Left            =   240
            TabIndex        =   24
            Top             =   1200
            Width           =   1455
         End
         Begin VB.Label lblfield 
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
            Index           =   0
            Left            =   240
            TabIndex        =   23
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label lblfield 
            BackStyle       =   0  'Transparent
            Caption         =   "Runtime (Min) :"
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
            Index           =   5
            Left            =   -74760
            TabIndex        =   22
            Top             =   1200
            Width           =   1455
         End
         Begin VB.Label lblfield 
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
            Index           =   6
            Left            =   -74760
            TabIndex        =   21
            Top             =   3600
            Width           =   855
         End
         Begin VB.Label lblfield 
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
            Index           =   7
            Left            =   -74760
            TabIndex        =   20
            Top             =   3120
            Width           =   1095
         End
         Begin VB.Label lbledit 
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
            Index           =   2
            Left            =   -74760
            TabIndex        =   19
            Top             =   720
            Width           =   1455
         End
      End
   End
End
Attribute VB_Name = "addentry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Option Explicit

Private Sub Form_Load()
        
    'binds grid to db
    Set dvdgrid1.DataSource = view.dvddb1
    
    'selects first row
    dvdgrid1.RowSel = 1
    
    'calls sub to sort out grid
    field_length
    
    dvdgrid1_RowColChange
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    view.Enabled = True
End Sub

Private Sub barDVD_ButtonClick(ByVal Button As ComctlLib.Button)
    
    'Launches action depending on button pressed
    Select Case Button
        Case "Go Back"
            Unload Me
        Case "Add Entry":
            add_entry
        Case "Edit Entry"
            edit_entry
        Case "Delete Entry":
            delete_entry
    End Select
End Sub
Public Sub add_entry()
    
    'shows and hides specific frames
    frmadd.Visible = True
    frmedit.Visible = False
    
    'makes the add button default
    cmdadd.Default = True
End Sub

Private Sub cmdAdd_Click()
    
    'calls upon the validate function
    validate_add
    
End Sub
Private Sub cmdclear_Click()
    
    'clears all fields
    If MsgBox("Are you sure you would like to clear all fields?", vbQuestion + vbYesNo, "Clear Field") = vbYes Then
        clear_fields
    End If
    
End Sub

Public Sub edit_entry()
    
    'hides and show specific frames
    frmadd.Visible = False
    frmedit.Visible = True
    
    cmdapply.Default = True
    
End Sub

Private Sub cmdapply_Click()
    
    'applies changes made
    If MsgBox("Are you sure you want to change the DVD?", vbQuestion + vbYesNo, "Apply Changes") = vbYes Then
        validate_edit
        editlog
    End If
End Sub

Public Sub delete_entry()
    
    'checkes to see if row selected is valid
    If dvdgrid1.Row = 0 Then
        MsgBox "You cannot delete the header.", vbInformation, "Cannot Delete"
        Exit Sub
    End If
    
    'confirm deletion
    If MsgBox("Are you sure you want to delete " & "'" & dvdgrid1.TextMatrix(dvdgrid1.Row, 0) & "' ?", vbQuestion + vbYesNo, "Are you sure you want to delete?") = vbYes Then
        deletelog
        
        ' delete an entry from the database
        With view.dvddb1.Recordset
            .Delete
            .Move (dvdgrid1.RowSel - 1) ' we minus one because row zero is the header row
            .Requery
        End With
        Set_datasource
        dvdcount
    End If
End Sub

Private Sub dvdgrid1_RowColChange()

    'loads data from db depending on clicked row
    If dvdgrid1.Row > 0 Then
        view.dvddb1.Recordset.AbsolutePosition = dvdgrid1.Row
    End If
    
    'calls on image
    load_imageedit
    
    'binds the textboxes
    For i = 0 To 10
        txtfield(i).Text = view.dvddb1.Recordset.Fields(i)
    Next
    
    'binds the comboboxes
    For i = 0 To 5
        cmbedit(i).Text = view.dvddb1.Recordset.Fields(i + 11)
    Next
    
    txtpicpath.Text = App.Path & view.dvddb1.Recordset.Fields(18)
    
    login.picopen.FileName = App.Path & view.dvddb1.Recordset.Fields(18)
    
End Sub

Private Sub dvdgrid1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If this is not row 0, do nothing.
    If dvdgrid1.MouseRow <> 0 Then Exit Sub

    ' Sort by the clicked column.
    If view.dvddb1.Recordset.Sort = view.dvddb1.Recordset.Fields(dvdgrid1.MouseRow).Name & " ASC" Then
        view.dvddb1.Recordset.Sort = view.dvddb1.Recordset.Fields(dvdgrid1.MouseRow).Name & " DESC"
    Else
        view.dvddb1.Recordset.Sort = view.dvddb1.Recordset.Fields(dvdgrid1.MouseRow).Name & " ASC"
    End If
End Sub

Private Sub txtadd_KeyPress(Index As Integer, KeyAscii As Integer)
    
    'integer only textbox
    If Index = 1 Or Index = 5 Then
        If (Not IsNumeric(ChrW(KeyAscii)) And Not (KeyAscii = vbKeyBack)) Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtfield_KeyPress(Index As Integer, KeyAscii As Integer)
    
    'integer only textbox
    If Index = 1 Or Index = 5 Then
        If (Not IsNumeric(ChrW(KeyAscii)) And Not (KeyAscii = vbKeyBack)) Then
            KeyAscii = 0
        End If
    End If
End Sub

Public Sub validate_add()
    
    'validates add textboxes
    For i = 0 To 10
        If txtadd(i).Text = "" Then
            MsgBox "Please enter the " & lbladd(i).Caption
            Exit Sub
        End If
    Next
    
    'valides add comboboxes
    For i = 0 To 5
        If cmbadd(i).ListIndex = -1 Then
            MsgBox "Please select the " & lblcmb(i).Caption
            Exit Sub
        End If
    Next
    
    addlog
    add_record
    Set_datasource
    clear_fields
    dvdcount
    
    
End Sub

Public Sub add_record()
    
    ' add a new entry to our table.
    With view.dvddb1.Recordset
        .AddNew
        !Title = txtadd(0)
        !Release_Date = txtadd(1)
        !Genre = cmbadd(0).List(cmbadd(0).ListIndex)
        !Studio = cmbadd(1).List(cmbadd(1).ListIndex)
        !Director = txtadd(2)
        !Producer = txtadd(3)
        !Writer = txtadd(4)
        !Classification = cmbadd(2).List(cmbadd(2).ListIndex)
        !Runtime = txtadd(5)
        !Disks = cmbadd(5)
        !Audio = cmbadd(3).List(cmbadd(3).ListIndex)
        !Ratio = cmbadd(4).List(cmbadd(4).ListIndex)
        !Country = txtadd(6)
        !Language = txtadd(7)
        !Awards = txtadd(8)
        !Official_Site = txtadd(9)
        !Plot = txtadd(10)
        '!Cover = "\Images\" & login.picopen.FileTitle
        .Update
        .Requery
    End With
    
    'FileCopy login.picopen.FileName, App.Path & "\Images\" & login.picopen.FileTitle
    
    
End Sub
Public Sub clear_fields()
    
    'clears all field in add screen
    For i = 0 To 10
        txtadd(i).Text = ""
    Next
    For i = 0 To 5
        cmbadd(i).ListIndex = -1
    Next
    addtab.Tab = 0
End Sub


Public Sub validate_edit()
    
    'validates add fields to check there's data in all fields
    For i = 0 To 10
        If txtfield(i).Text = "" Then
            MsgBox "Please enter the " & lblfield(i).Caption
            Exit Sub
        End If
    Next
    
    edit_fields
    Set_datasource
    
End Sub

Public Sub edit_fields()
    
    ' edits entry in the record
    With view.dvddb1.Recordset
        !Title = txtfield(0)
        !Release_Date = txtfield(1)
        !Genre = cmbedit(0).List(cmbedit(0).ListIndex)
        !Studio = cmbedit(1).List(cmbedit(1).ListIndex)
        !Director = txtfield(2)
        !Producer = txtfield(3)
        !Writer = txtfield(4)
        !Classification = cmbedit(2).List(cmbedit(2).ListIndex)
        !Runtime = txtfield(5)
        !Audio = cmbedit(3).List(cmbedit(3).ListIndex)
        !Ratio = cmbedit(4).List(cmbedit(4).ListIndex)
        !Country = txtfield(6)
        !Language = txtfield(7)
        !Awards = txtfield(8)
        !Official_Site = txtfield(9)
        !Plot = txtfield(10)
        '!Cover = "\Image\" & login.picopen.FileTitle
        .Update
        .Requery
    End With
    
    
End Sub

Public Sub field_length()
    dvdgrid1.ColWidth(0) = dvdgrid1.Width * 0.99
End Sub

Public Sub editlog()
    
    'add edit log
    With login.dblog.Recordset
        .AddNew
        !Time = Time
        !Date = Date
        !Event = "DVD '" & dvdgrid1.TextMatrix(dvdgrid1.Row, 0) & "' was edited"
        !User = Username
        .Update
        .Requery
    End With
    
    rebind_log
    
End Sub

Public Sub deletelog()

    'add delete log
    With login.dblog.Recordset
        .AddNew
        !Time = Time
        !Date = Date
        !Event = "DVD '" & dvdgrid1.TextMatrix(dvdgrid1.Row, 0) & "' was deleted"
        !User = Username
        .Update
        .Requery
    End With
    
    rebind_log
    
End Sub

Public Sub addlog()

    'add add log
    With login.dblog.Recordset
        .AddNew
        !Time = Time
        !Date = Date
        !Event = "DVD '" & txtadd(0).Text & "' was added"
        !User = Username
        .Update
        .Requery
    End With
    
    rebind_log
    
End Sub

Private Sub cmdbrowseadd_Click()
    
    login.picopen.ShowOpen
    
    txtpicpathadd.Text = login.picopen.FileName
    
    imgaddcover.Picture = LoadPicture(login.picopen.FileName)
    
    
End Sub

Private Sub cmdbrowse_Click()
    login.picopen.ShowOpen
End Sub
