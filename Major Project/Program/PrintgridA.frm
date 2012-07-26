VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form P001 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Options"
   ClientHeight    =   9810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12750
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9810
   ScaleWidth      =   12750
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
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
      Left            =   2805
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   77
      TabIndex        =   90
      Top             =   2010
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3165
      Index           =   6
      Left            =   8040
      TabIndex        =   27
      Top             =   6120
      Visible         =   0   'False
      Width           =   4665
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Page Numbering"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2715
         Index           =   10
         Left            =   1800
         TabIndex        =   43
         Top             =   240
         Width           =   2775
         Begin VB.CheckBox Check1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Print above top margin"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   10
            Left            =   225
            TabIndex        =   75
            Top             =   1965
            Width           =   2310
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Print below bottom margin"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   225
            TabIndex        =   49
            Top             =   2295
            Width           =   2250
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Format as Page n of n"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   225
            TabIndex        =   48
            Top             =   1635
            Width           =   2055
         End
         Begin VB.ComboBox Combo1 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   900
            Style           =   2  'Dropdown List
            TabIndex        =   47
            Top             =   1230
            Width           =   1695
         End
         Begin VB.OptionButton Option4 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Number all pages"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   225
            TabIndex        =   46
            Top             =   840
            Width           =   1830
         End
         Begin VB.OptionButton Option4 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Number pages after first"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   225
            TabIndex        =   45
            Top             =   585
            Value           =   -1  'True
            Width           =   2175
         End
         Begin VB.OptionButton Option4 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Omit page numbers"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   225
            TabIndex        =   44
            Top             =   285
            Width           =   1830
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Location"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   14
            Left            =   225
            TabIndex        =   50
            Top             =   1275
            Width           =   780
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Page Margins"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2715
         Index           =   9
         Left            =   60
         TabIndex        =   28
         Top             =   315
         Width           =   1620
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   7
            Left            =   690
            TabIndex        =   33
            Text            =   "1"
            Top             =   1545
            Width           =   450
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   6
            Left            =   690
            TabIndex        =   32
            Text            =   "1"
            Top             =   1155
            Width           =   450
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   5
            Left            =   690
            TabIndex        =   31
            Text            =   "1"
            Top             =   795
            Width           =   450
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Default"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   675
            Style           =   1  'Graphical
            TabIndex        =   30
            ToolTipText     =   "Click to set margins to the printer default values"
            Top             =   1950
            Width           =   750
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   4
            Left            =   675
            TabIndex        =   29
            Text            =   "1"
            Top             =   420
            Width           =   450
         End
         Begin MSComCtl2.UpDown UpDown1 
            Height          =   345
            Index           =   3
            Left            =   1185
            TabIndex        =   34
            Top             =   405
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   609
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UpDown1 
            Height          =   345
            Index           =   4
            Left            =   1185
            TabIndex        =   35
            Top             =   810
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   609
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UpDown1 
            Height          =   345
            Index           =   5
            Left            =   1185
            TabIndex        =   36
            Top             =   1185
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   609
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UpDown1 
            Height          =   345
            Index           =   6
            Left            =   1185
            TabIndex        =   37
            Top             =   1560
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   609
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Right"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   315
            Index           =   12
            Left            =   120
            TabIndex        =   41
            Top             =   1590
            Width           =   435
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Left"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   315
            Index           =   11
            Left            =   120
            TabIndex        =   40
            Top             =   1200
            Width           =   435
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Bottom"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   315
            Index           =   10
            Left            =   90
            TabIndex        =   39
            Top             =   825
            Width           =   495
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Top"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   315
            Index           =   9
            Left            =   105
            TabIndex        =   38
            Top             =   435
            Width           =   435
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "mm"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   240
            Index           =   13
            Left            =   765
            TabIndex        =   42
            Top             =   195
            Width           =   375
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Index           =   11
      Left            =   60
      TabIndex        =   62
      Top             =   3195
      Width           =   4560
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Font..."
         Height          =   345
         Index           =   4
         Left            =   3540
         Style           =   1  'Graphical
         TabIndex        =   89
         Top             =   1035
         Width           =   945
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Print &Row Captions on all pages"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   12
         Left            =   2640
         TabIndex        =   86
         ToolTipText     =   "Select to repeat the left hand column on all additional pages"
         Top             =   2610
         Width           =   1755
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Print Column &Headings on all pages"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   4
         Left            =   165
         TabIndex        =   74
         Top             =   2625
         Width           =   1875
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Height          =   1575
         Index           =   13
         Left            =   105
         TabIndex        =   68
         Top             =   1035
         Width           =   1920
         Begin VB.OptionButton Option3 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Once each row"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   4
            Left            =   105
            TabIndex        =   73
            Top             =   1200
            Width           =   1470
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00E0E0E0&
            Caption         =   "First row only"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   105
            TabIndex        =   72
            Top             =   951
            Width           =   1470
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Omit"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   105
            TabIndex        =   71
            Top             =   210
            Width           =   1215
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00E0E0E0&
            Caption         =   "First page only"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   105
            TabIndex        =   70
            Top             =   457
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Print on all pages"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   105
            TabIndex        =   69
            Top             =   704
            Width           =   1785
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Height          =   1095
         Index           =   12
         Left            =   2640
         TabIndex        =   64
         Top             =   1485
         Width           =   1830
         Begin VB.OptionButton Option5 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Left Justify"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   0
            Left            =   195
            TabIndex        =   67
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton Option5 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Centre Justify"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   1
            Left            =   195
            TabIndex        =   66
            Top             =   480
            Width           =   1425
         End
         Begin VB.OptionButton Option5 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Right Justify"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   2
            Left            =   195
            TabIndex        =   65
            Top             =   720
            Width           =   1455
         End
      End
      Begin VB.TextBox Text1 
         Height          =   780
         Index           =   3
         Left            =   90
         TabIndex        =   63
         Top             =   195
         Width           =   4395
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   3225
      Index           =   0
      Left            =   8010
      TabIndex        =   4
      Top             =   2730
      Width           =   4680
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Height          =   3105
         Index           =   2
         Left            =   60
         TabIndex        =   10
         Top             =   30
         Width           =   1815
         Begin VB.OptionButton Option1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "&Print 3 up on 2 pages"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   2
            Left            =   120
            TabIndex        =   87
            Top             =   2415
            Width           =   1485
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Print &Full Size"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   105
            TabIndex        =   14
            Top             =   210
            Value           =   -1  'True
            Width           =   1395
         End
         Begin VB.TextBox Text1 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   390
            TabIndex        =   13
            Text            =   "1"
            Top             =   1920
            Width           =   315
         End
         Begin VB.TextBox Text1 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   390
            TabIndex        =   12
            Text            =   "1"
            Top             =   1440
            Width           =   315
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "&Scale To Fit"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   120
            TabIndex        =   11
            Top             =   1080
            Width           =   1275
         End
         Begin MSComCtl2.UpDown UpDown1 
            Height          =   345
            Index           =   1
            Left            =   735
            TabIndex        =   15
            Top             =   1890
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   609
            _Version        =   393216
            Enabled         =   0   'False
         End
         Begin MSComCtl2.UpDown UpDown1 
            Height          =   345
            Index           =   0
            Left            =   735
            TabIndex        =   76
            Top             =   1425
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   609
            _Version        =   393216
            Enabled         =   0   'False
         End
         Begin VB.Line Line1 
            Index           =   1
            X1              =   105
            X2              =   1680
            Y1              =   2325
            Y2              =   2325
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "nn Sheets High"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   2
            Left            =   360
            TabIndex        =   19
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "nn Sheets Wide"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   1
            Left            =   360
            TabIndex        =   18
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "High"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   315
            Index           =   6
            Left            =   1080
            TabIndex        =   17
            Top             =   1920
            Width           =   480
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Wide"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   315
            Index           =   5
            Left            =   1050
            TabIndex        =   16
            Top             =   1440
            Width           =   480
         End
         Begin VB.Line Line1 
            Index           =   0
            X1              =   120
            X2              =   1680
            Y1              =   960
            Y2              =   960
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Paper Size"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1605
         Index           =   4
         Left            =   1860
         TabIndex        =   20
         Top             =   30
         Width           =   2805
         Begin VB.ComboBox cbPaperSize 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   420
            Style           =   2  'Dropdown List
            TabIndex        =   82
            Top             =   1230
            Width           =   2295
         End
         Begin VB.OptionButton optPPS 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Landscape"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   150
            TabIndex        =   81
            Top             =   990
            Width           =   1335
         End
         Begin VB.OptionButton optPPS 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Portrait"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   80
            Top             =   750
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "Size N"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   360
            Index           =   8
            Left            =   1680
            TabIndex        =   21
            Top             =   420
            Width           =   915
            WordWrap        =   -1  'True
         End
         Begin VB.Image Image1 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1950
            Picture         =   "PrintgridA.frx":0000
            Stretch         =   -1  'True
            ToolTipText     =   "Click to change paper orietation."
            Top             =   300
            Width           =   615
         End
         Begin VB.Label Label1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "NN mm High"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   4
            Left            =   180
            TabIndex        =   23
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "NN mm Wide"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   3
            Left            =   180
            TabIndex        =   22
            Top             =   255
            Width           =   1260
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Print Range"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1485
         Index           =   3
         Left            =   1860
         TabIndex        =   6
         Top             =   1650
         Width           =   2805
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   8
            Left            =   945
            TabIndex        =   55
            Top             =   810
            Width           =   1200
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00E0E0E0&
            Caption         =   "&All"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   9
            Top             =   330
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Pages"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   150
            TabIndex        =   8
            Top             =   825
            Width           =   975
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Selected &Range"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   150
            TabIndex        =   7
            Top             =   585
            Width           =   1935
         End
      End
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Show Print Progress"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   11
      Left            =   60
      TabIndex        =   85
      Top             =   1740
      Value           =   1  'Checked
      Width           =   1890
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3000
      Index           =   8
      Left            =   30
      TabIndex        =   56
      Top             =   6345
      Width           =   4575
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Print Cell Graphics"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   150
         TabIndex        =   88
         Top             =   1200
         Width           =   2130
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Merged Cells"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1170
         Index           =   14
         Left            =   15
         TabIndex        =   77
         Top             =   1815
         Width           =   4515
         Begin VB.OptionButton Option6 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Across horizontal page breaks"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   330
            TabIndex        =   79
            Top             =   735
            Width           =   3150
         End
         Begin VB.OptionButton Option6 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Within each page only"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   360
            TabIndex        =   78
            Top             =   390
            Value           =   -1  'True
            Width           =   2325
         End
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Print &Background Colours"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   2310
         TabIndex        =   61
         Top             =   720
         Width           =   2205
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Print &Text Colours"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2310
         TabIndex        =   60
         Top             =   465
         Width           =   1935
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Print &Grid Lines"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   150
         TabIndex        =   59
         Top             =   465
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Print &Grid Line Colours"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   150
         TabIndex        =   58
         Top             =   840
         Width           =   2205
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Shade Column Headings"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   2310
         TabIndex        =   57
         Top             =   1200
         Value           =   1  'Checked
         Width           =   2220
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   840
      Index           =   1
      Left            =   60
      TabIndex        =   5
      Top             =   795
      Visible         =   0   'False
      Width           =   2220
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   285
         Left            =   75
         TabIndex        =   84
         Top             =   480
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Page nn of nn"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   105
         TabIndex        =   83
         Top             =   135
         Width           =   2010
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   2505
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   960
      Width           =   945
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   3765
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1095
      Width           =   945
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   9390
      Width           =   12750
      _ExtentX        =   22490
      _ExtentY        =   741
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Picture         =   "PrintgridA.frx":0052
            Object.ToolTipText     =   "The currently selected printer"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Pages"
            TextSave        =   "Pages"
            Object.ToolTipText     =   "The number of pages you have decided to print"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Text            =   "Copies"
            TextSave        =   "Copies"
            Object.ToolTipText     =   "The number of copies of the document that will be printed"
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   75
      Top             =   2115
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3165
      Index           =   5
      Left            =   5505
      TabIndex        =   24
      Top             =   -75
      Visible         =   0   'False
      Width           =   4395
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
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
         Index           =   7
         Left            =   105
         TabIndex        =   51
         Top             =   2385
         Width           =   2775
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   1080
            TabIndex        =   52
            Text            =   "1"
            Top             =   240
            Width           =   555
         End
         Begin MSComCtl2.UpDown UpDown1 
            Height          =   345
            Index           =   2
            Left            =   1665
            TabIndex        =   53
            Top             =   225
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   609
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Number Of Copies"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   54
            Top             =   120
            Width           =   975
         End
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Set-Up..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   3180
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Use this button to access the Windows printer set-up"
         Top             =   1140
         Width           =   1080
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   150
         TabIndex        =   25
         Text            =   "Combo1"
         Top             =   750
         Width           =   4110
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   660
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   4950
      _ExtentX        =   8731
      _ExtentY        =   1164
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Selection"
            Key             =   "select"
            Object.ToolTipText     =   "Select pages to print"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Title"
            Key             =   "title"
            Object.ToolTipText     =   "Control the titles shown on the printed pages"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Printer"
            Key             =   "print"
            Object.ToolTipText     =   "Select the printer and change printer settings"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Effects"
            Key             =   "effects"
            Object.ToolTipText     =   "Set colour and other presentation options"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Page"
            Key             =   "page"
            Object.ToolTipText     =   "Set Page Margins and Page Numbering"
            ImageVarType    =   2
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
   Begin VB.Image Image2 
      Height          =   10080
      Index           =   1
      Left            =   0
      Picture         =   "PrintgridA.frx":0164
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12960
   End
End
Attribute VB_Name = "P001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'© Copyright Adit Limited 1998 to 2005
'Version 2.09 (beta of 2.1 release) - please see README.TXT for new features and bug fixes.
'This is intended to be the final version - although further bug fixes may be applied
'The full API documentation is contained in an accompanying document

'TEMPORARY WORKING NOTES FOR BETA
'Fixed Alignment (column headers)error FIXED
'Faster processing of Merged Cells

'General Declarations
Private FormDisplayMode As Long
Private POOrientation As Long, LastLineColour As Long
Private CurrentFrame As Long, MCells() As Long, CCells() As Long
Private PrintPageWidth As Single, PrintPageHeight As Single
Private GridWidth As Double, GridHeight As Double, MarginAdjust As Double
Private ColWidths() As Double, RowHeights() As Double, WideCol As Double, HighRow As Double
Private DefSavePosition As Variant, ColAlign() As Variant, Selectedpages As Variant
Private TitleFont As New StdFont

Private Msg As String, title As String
Private UserSelection As Boolean, PrintStarted As Boolean, InitialisingForm As Boolean
Private CancelSheetPrint As Boolean, ScalePrint As Boolean, CheckNeeded As Boolean
Private AddHeadingsToP1 As Boolean, GridFont As Boolean, changingPPS As Boolean
Private ShowGridColour As Boolean, ShowLines As Boolean, ShadeFixedRows As Boolean
Private ShowTextColour As Boolean, ShowCellColour As Boolean, InvisibleGrid As Boolean
Private PrintMultiGrids As Boolean, ProgressBarDisplayed As Boolean
Private UserSelectedPagesCount As Long, RowsToPrint As Long, ColsToPrint As Long
Private ColSpace As Single

Private defPagesWide As Long, defPagesHigh As Long
Private userPagesWide As Long, userPagesHigh As Long
Private CurrentPageTop As Long

'Error recovery subsytem declarations
Private Const MyE2999 = 2999
Private ClassErrorCode As Long
Private ClassErrorSource As String
Private ClassErrorMessage As String
Private errorResponse As Long

'Form Constants
Private Const TwipsMM = 56.7
Private Const TwipsTI = 144
Private Const MYBLACK = -2147483630
Private Const COLSPACEMM = 1 'sets column spacing in mm - could become a variable
Private Const DefFixB = &H8000000F
Private Const MaxImageSize = 956
Private Const DISPLAY_DIALOGUE = 0
Private Const DISPLAY_PROGRESS = 1
Private Const SECTION_GAP = 373 'This can be adjusted to change the gap between
                                'Grid sections printed on the same page

'Enums
Public Enum PrintSettings   'retained for backwards compatibility
    SHOW_LINES = 16         'Sets Grid line Printing
    TITLE_FIRSTPAGE = 64   'Print a title on the first page
    TITLE_ALLPAGES = 128   'Print a title on all pages after first row
    REPEAT_HEADINGS = 256  'Repeat Fixed grid rows on subsequent pages
    GRID_FONT = 512        'Supports bold, italic and underlined cells in Grid
    ALLOW_SELECTED = 1024  'Allows the user to elect to print selected area of grid only
    ALLOW_DIALOGUE = 1     'Set to show the dialogue form to the user
    ALLOW_COLOURS = 2      'Allows Grid colours to be reproduced (background colours require grid to be printed as well
    ALLOW_SETLINES = 4     'allows the user to select to print Grid lines
    ALLOW_TITLE = 8        'Allows the user to change the report title settings

    GRID_NORMAL = 1 + 4 + 8 + 64 + 16 + 1024   'Default settings
                                                'Allow User to see dialogue
                                                'Print the title on the first page
                                                'Allow the user to change the report title
                                                'allow the user to choose to print the grid lines
                                                'set the default to show the grid lines
                                                'lets the user elect to print just the selected area (if there is one)
End Enum

Public Enum TitleOption
    USER_MAY_SET = 1
    NO_TITLES = 2
    FIRST_PAGE_ONLY = 4
    TITLE_ALL_PAGES = 8
    TITLE_ALL_ROWS = 16
    TITLE_FIRST_ROW_ONLY = 32
    JUSTIFY_LEFT = 64
    JUSTIFY_CENTRE = 128
    JUSTIFY_RIGHT = 256
    REPEAT_COL_HEADINGS = 512
    FONT_BOLD = 1024
    FONT_UNDERLINE = 2048
    REPEAT_FIXED_COLS = 4096
    TITLE_NORMAL = 1 + 4 + 64 + 1024 + 2048
End Enum
Public Enum Effects
    USER_CAN_CHANGE = 1
    GRID_LINES = 2
    GRID_COLOURS = 4
    TEXT_COLOURS = 8
    CELL_BACK_COLOURS = 16
    COLUMN_HEAD_SHADING = 32
    CELL_FONT_EFFECTS = 64
    CELL_GRAPHICS = 128
    EFFECTS_NORMAL = 1 + 2 + 32 + 64    'produces a monochrome report with shaded column headings
End Enum
Public Enum PageNumbers
    USER_CAN_SET = 1
    AFTER_FIRST = 2
    ALL_PAGES = 4
    TOP_LEFT = 8
    TOP_RIGHT = 16
    BOTTOM_LEFT = 32
    BOTTOM_CENTRE = 64
    BOTTOM_RIGHT = 128
    INCLUDE_PAGE_COUNT = 256
    OVER_TOP_MARGIN = 512
    UNDER_BOTTOM_MARGIN = 1024
    TOP_CENTRE = 2048
    NUMBER_NORMAL = 1 + 2 + 16
End Enum
Public Enum CellMerge
    USER_HAS_CONTROL = 1
    MERGE_SAME_PAGE_ONLY = 2
    MERGE_ACROSS_ROWS = 4
    MERGE_NORMAL = 1 + 2
End Enum
Private Enum CellAlignment
    CellAligndef = 1
    CellAlignleft = 2
    CellAlignCentre = 4
    CellAlignRight = 8
    CellAlignTop = 16
    CellAlignMiddle = 32
    cellalignbottom = 64
End Enum
Public Enum PageOrientation
    PGPortrait = 1
    PGLandscape = 2
End Enum
Public Enum PrintSelect
    PS_MULTICELL = 1    'Enables Print Selection if more than one cell selected (default)
    PS_MULTICOL = 2     'Enables Print Selection if more than one Column is selected
    PS_MULTIROW = 4     'Enables Print Selection if more than one row selected
    PS_DEF_PRINT_SELECT = 8    'Change default action to print selection rather than all (subject to above options)
    PS_NOT_AVAILABLE = 16       'Disables the option to print a selected area
End Enum
Public Enum MultiColumn
    MC_SET_AS_DEFAULT = 1   'Default to print multiple grids on the same page if they will fit
    MC_DENY_MULTI_COL = 2   'Do not let the user have access to the option on the dialogue
End Enum

Public Enum JustifySubTitle
    LEFT_JUSTIFY = 1
    CENTER_JUSTIFY = 2
    RIGHT_JUSTIFY = 3
End Enum

Public Enum SubTitleUsage
    SUB_AS_MAIN_TITLE = 0 'always print under the main title and repeat as main title (default)
    SUB_FIRST_PAGE_ONLY = 1
    SUB_ALL_PAGES = 2
    SUB_ONCE_EACH_ROW = 4
End Enum


'Type declarations
Private Type PSize
    Name As String
    height As Single
    width As Single
    MSystem As Long
End Type
Private UserPSize As PSize
Private Type Margin
    top As Double
    bottom As Double
    left As Double
    right As Double
    height As Double
    width As Double
    hasbeenset As Boolean
End Type
Private defMargin As Margin, setMargin As Margin
'defMargin holds the current printable area of the page/printer
'setMargin holds a working copy of margins for the user to adjust

Private Type pNum
    numOption As Integer
    numPos As Integer
    incPCount As Boolean
    underMargin As Boolean
    overMargin As Boolean
    minimumTop As Single
    footHeight As Single
End Type
Private PageNumbs As pNum
Private Type page
    GridStart As Long
    GridEnd As Long
    Size As Double
End Type
Private PagesW() As page, PagesH() As page, PagesS() As page
Private Type Heads
    Page1Height As Double
    ColHeadHeight As Double
    NextPage As Double
    NextTitle As Double
    TitleWidth As Double
    FixedCols As Long
    Justify As Integer
    FontScale As Single
    Bold As Boolean
    Underlined As Boolean
    IncludeColHeadings As Boolean
    CHinPage1 As Boolean
    PrintPage As Long
End Type
Private Header As Heads
Private Type GColr
    Back As Long
    Front As Long
    Grid As Long
End Type
Private Type mergeprule
    MergeRows As Boolean
    MergeCols As Boolean
    MergeRule As Long
End Type
Private PrintMerge As mergeprule
Private Type CellC
    CellHeight As Single
    CellWidth As Single
    PrintHeight As Single
    PrintWidth As Single
    AlreadyUsed As Single
    WrapCount As Integer
    PrintLine As String
    GridText As String
    WrapLines() As Variant
End Type
Private Type PaperType
    Name As String
    Index As Integer
    width As Single
    height As Single
End Type
Private PaperTypes() As PaperType
Private Type FixedColDat
    HighCol As Long
    FixedWidth As Single
    FixedPage1 As Single
    RepeatFixed As Boolean
End Type
Private FixedColData As FixedColDat

'API Declarations
Private Declare Function GetLocaleInfo Lib "KERNEL32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Const LOCALE_IMEASURE = &HD         '  0 = metric, 1 = US
Private Const USER_DEFAULT = &H400
Private Declare Function SetBkMode Lib "gdi32" _
      (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private iBKMode As Long
Private Const TRANSPARENT = 1
Private Const OPAQUE = 2
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Const PHYSICALOFFSETX As Long = 112
Private Const PHYSICALOFFSETY As Long = 113

Private Declare Function DeviceCapabilities Lib "winspool.drv" Alias "DeviceCapabilitiesA" (ByVal lpDeviceName As String, ByVal lpPort As String, ByVal iIndex As Long, lpOutput As Any, ByVal Dev As Long) As Long

Private Const DC_PAPERNAMES = 16
Private Const DC_PAPERS = 2
Private Const DC_PAPERSIZE = 3

Private Type POINTAPI
        x As Long
        y As Long
End Type
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Const SW_SHOWNOACTIVATE = 4, SW_HIDE = 0

Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Private Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest

'Private copies of PrintGrid API Public Property Values
Private pFlexgrid As MSHFlexGrid
Private mReportTitle As String
Private mAllowDialogue As Boolean, mShowGridLines As Boolean
Private mPrintSelected As Boolean, mPrintComplete As Boolean, mSetPages As Boolean
Private mShowProgress As Boolean
Private mPrintCopies As Long
Private mTitlePages As TitleOption
Private mEffects As Effects
Private mPageNumbering As PageNumbers
Private mAPIMarginUnit As ScaleModeConstants
Private mSelectionPrintRule As PrintSelect
Private mMergeRule As CellMerge
Private mPrinterOrientation As PageOrientation
Private mMarginTop As Single, mMarginBottom As Single, mMarginLeft As Single, mMarginRight As Single
Private mSetPrinter As Variant
Private mRequestPagesWide As Long, mRequestPagesHigh As Long
Private mRequestSectionsWide As Long, mRequestSectionsHigh
Private mPrintProgress As Long
Private mMultiColumnPrint As MultiColumn
Private mySubHeadings As Variant, nullHeading As Variant
Private mySubHeadFont As New StdFont
Private mySubJustify As Integer
Private mPLColStart As Long, mPLColEnd As Long, mPLRowStart As Long, mPLRowEnd As Long
Private mRepeatSubTitle As SubTitleUsage
Private mPrintCellImages As Boolean, mProportionalCompression As Boolean


Public Sub CountGridPages()
    'This routine just counts the pages that would be output if the PrintGridAPI sub was called
    Dim savefont As Single
    
    On Error GoTo NoCGPrinterErr
    PrintStarted = False
    mPrintComplete = False
    Printer.ScaleMode = vbTwips 'Checks there is a printer somewhere - otherwise not much point in all this
    DoEvents
    On Error GoTo CountGridPagesErr
    Screen.MousePointer = vbHourglass
    savefont = Printer.FontSize
    Printer.Font = pFlexgrid.Font
    SaveGridPosition
    ReSetForm
    GridWidth = GetColWidths
    GridHeight = GetRowHeights
    CountPages False
    mPrintProgress = 0
    mRequestPagesHigh = defPagesHigh
    mRequestPagesWide = defPagesWide
    mSetPages = False
    Screen.MousePointer = vbDefault
    Exit Sub
NoCGPrinterErr:
    mRequestPagesHigh = 0
    mRequestPagesWide = 0
    Exit Sub
CountGridPagesErr:
   Select Case Err
        Case 484
            Resume Next 'printer driver does not support what we just tried
        Case vbObjectError + MyE2999
            'Abort selected in a function called from this sub or function
        Case Else
            Screen.MousePointer = vbDefault
            errorResponse = MsgBox("Unexpected Error Occured." & vbLf & "Please note details." & vbLf & Error(Err) & vbLf & "in Sub:CountGridPages", vbExclamation + vbAbortRetryIgnore, "Unexpected Error")
            Select Case errorResponse
                Case vbAbort
                Case vbRetry
                    Resume
                Case vbIgnore
                    Resume Next
            End Select
    End Select
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub
Public Property Let BottomMargin(ByVal sNewValue As Single)
    Dim workscale As Integer
    
    Select Case mAPIMarginUnit
        Case 1 To 8
            workscale = mAPIMarginUnit
        Case Else   'if not set via API use PC default units
            If LocaleUnits = 1 Then
                workscale = vbInches
            Else
                workscale = vbMillimeters
            End If
    End Select
    sNewValue = Printer.ScaleY(sNewValue, workscale, vbTwips)
    mMarginBottom = sNewValue
End Property
Public Sub CancelPrint()
    On Error Resume Next
    Printer.KillDoc
End Sub

Private Function CalcScale(SetPrint As Boolean) As Double
    'Returns the FontFactor, If SetPrint then sets the printer ScaleMode
    'Sub re-written for version 2.1 to improve the guesswork
    Dim PScaleWidth As Double, PScaleHeight As Double, PWorkWidth As Double, PWorkHeight As Double
    Dim PrintSWidth As Double, PrintSHeight As Double, FontFactor As Double
    Dim FudgeHFactor As Double, FudgeWFactor As Double, Headheight As Double
    
    'then for testing
    Dim oldpagesWide As Long, oldPagesHigh As Long
    
    oldpagesWide = userPagesWide
    oldPagesHigh = userPagesHigh
    
    PScaleWidth = Printer.ScaleWidth
    PScaleHeight = Printer.ScaleHeight
    PWorkWidth = (Printer.ScaleWidth - (setMargin.left + setMargin.right))
    PWorkHeight = (Printer.ScaleHeight - (setMargin.top + setMargin.bottom))
    
    If userPagesWide > 1 Then
        FudgeWFactor = GridWidth / ColsToPrint
        FudgeWFactor = FudgeWFactor * userPagesWide
        If WideCol > FudgeWFactor Then
            FudgeWFactor = WideCol
        End If
    Else
        FudgeWFactor = TwipsMM
    End If
    Headheight = Header.Page1Height + TwipsMM
    If userPagesHigh > 1 Then
        FudgeHFactor = GridHeight / RowsToPrint 'average row height
        FudgeHFactor = FudgeHFactor * userPagesHigh
        If FudgeHFactor < HighRow Then
            FudgeHFactor = HighRow
        End If
        If Header.NextPage > Headheight Then
            Headheight = Header.NextPage + TwipsMM
        End If
    Else
        FudgeHFactor = TwipsMM
    End If
    
    
    PrintSWidth = (GridWidth + FudgeWFactor) / userPagesWide
    PrintSHeight = (GridHeight + FudgeHFactor) / userPagesHigh
    PrintSHeight = PrintSHeight + Headheight
    
    PrintSWidth = PrintSWidth * (PScaleWidth / PWorkWidth)
    PrintSHeight = PrintSHeight * (PScaleHeight / PWorkHeight)
    
    'we have a choice now (Version 2.1) about any required compression
    'for text only output it is generally best to just re-scale the printer graphics object
    'to the minimum required extent - as previous version
    'where cell graphics are included it can look better if the compression is symetrical
    'the new ProportionalCompression boolean property can be used to select the latter option
    If PScaleWidth > PrintSWidth Then
        PrintSWidth = PScaleWidth
        FudgeWFactor = 1
    Else
        FudgeWFactor = PrintSWidth / PScaleWidth
    End If
    If PScaleHeight > PrintSHeight Then
        PrintSHeight = PScaleHeight
        FudgeHFactor = 1
    Else
        FudgeHFactor = PrintSHeight / PScaleHeight
    End If
    If mProportionalCompression Then
        If FudgeWFactor <> 1 Or FudgeHFactor <> 1 Then
            If FudgeWFactor > FudgeHFactor Then
                PrintSHeight = PrintSHeight * FudgeWFactor
            Else
                PrintSWidth = PrintSWidth * FudgeHFactor
            End If
        End If
    End If
    'sort out the grid font size to match any compression
    FontFactor = CalcFontFactor2(PrintSWidth, PrintSHeight, Printer.Font)
   
    If SetPrint Then    'we are ready to print
        Printer.ScaleHeight = PrintSHeight
        Printer.ScaleWidth = PrintSWidth
        CountPages False    'get the pages re-worked ready for the print
        
        'while testing - make sure that the re-counted pages match the requirement
        'If (oldpagesWide <> defPagesWide) Or (oldPagesHigh <> defPagesHigh) Then
        '    Stop
        'End If
    End If
    
    CalcScale = FontFactor
End Function

Private Sub DisplayPageCount()
'© Copyright Adit Limited 1998, 1999, 2000, 2001, 2002
    Dim TotPages As Long
    Dim CaptionString As String
    
    On Error Resume Next
    Label1(1).Caption = CStr(defPagesWide) & " Sheets Wide"
    Label1(2).Caption = CStr(defPagesHigh) & " Sheets High"
    TotPages = defPagesWide * defPagesHigh
    If TotPages = 0 Then
        Command1(0).Enabled = False
    Else
        Command1(0).Enabled = True
    End If
    If TotPages = 1 Then    'no compression is available
        Option1(0).Value = True
        Option1(1).Enabled = False
        Check1(11).Visible = False 'plus no point in print progress
        mShowProgress = False
    Else
        Option1(1).Enabled = True
        If Check1(11).Enabled Then
            Check1(11).Visible = True
        End If
    End If
    If Option1(0).Value Or (Me.Visible = False And Not mSetPages) Or defPagesWide < Text1(0).Text Or defPagesHigh < Text1(1).Text Then
        Text1(0).Text = defPagesWide
        Text1(1).Text = defPagesHigh
        userPagesHigh = defPagesHigh
        userPagesWide = defPagesWide
    End If
    If defPagesWide > 1 Or defPagesHigh > 1 Then
        Option2(2).Enabled = True
    Else
        Option2(2).Enabled = False
    End If
    If Not (mMultiColumnPrint And MC_DENY_MULTI_COL) Then
        If mRequestSectionsWide > 1 And defPagesHigh > 1 Then
            CaptionString = "&Print " & CStr(mRequestSectionsWide) & " up on " & CStr(mRequestSectionsHigh) & " page"
            If mRequestSectionsHigh > 1 Then
                CaptionString = CaptionString & "s"
            End If
            Option1(2).Caption = CaptionString
            If mMultiColumnPrint And MC_SET_AS_DEFAULT Then
                Option1(2).Value = True
            End If
            Option1(2).Visible = True
            Line1(1).Visible = True
        Else
            'things may have changed so make sure the controls are not available or set
            Line1(1).Visible = False
            If Option1(2).Value Then
                Option1(0).Value = True
            End If
            Option1(2).Visible = False
            PrintMultiGrids = False 'should this have been set by Option1()_Click?
        End If
    End If
    StatusBarPages
    
End Sub

Private Sub FindMergedCol(FromRow As Long, ToRow As Long, FromCol As Long, ToCol As Long)
    'MSHFlexGrid cell merging is buggy - even at the last fix in SP6 and we are unlikely to see another applied to this control
    'The PrintGrid routine attempts to support the best of what works
    'This routine looks for the merged cells within vertical columns within the range of rows passed to it
    Dim MsetCount As Long, RowLoop As Long, FixPoint As Long, ColLoop As Long
    'We get passed a group of rows/columns and we return an array indicating any
    'merged cells in the columns. Merges are not allowed to bridge the border
    'between fixed rows and non fixed rows
    ReDim CCells(FromRow To ToRow, FromCol To ToCol)
    MsetCount = 0
    FixPoint = pFlexgrid.FixedRows
    For ColLoop = FromCol To ToCol
        If pFlexgrid.MergeCol(ColLoop) And ColWidths(ColLoop) > 0 Then
            For RowLoop = FromRow + 1 To ToRow
                If pFlexgrid.TextMatrix(RowLoop - 1, ColLoop) = pFlexgrid.TextMatrix(RowLoop, ColLoop) Then
                    If ((RowLoop - 1) < FixPoint And RowLoop < FixPoint) Or ((RowLoop - 1) >= FixPoint) Then
                        If CCells(RowLoop - 1, ColLoop) = 0 Then
                            MsetCount = MsetCount + 1000
                            CCells(RowLoop - 1, ColLoop) = MsetCount + 1
                        End If
                        CCells(RowLoop, ColLoop) = CCells(RowLoop - 1, ColLoop) + 1
                    End If
                End If
            Next RowLoop
        End If
    Next ColLoop
End Sub
Private Sub FindMergedRow(GridRow As Long, FromCol As Long, ToCol As Long)
'© Copyright Adit Limited 1998, 1999, 2000, 2001, 2002
    Dim MsetCount As Long, ColLoop As Long, FixPoint As Long
    Dim MergeVert As Boolean
    
    On Error GoTo FindMergedRowErr
    ReDim MCells(0 To ToCol)
    MergeVert = False
    FixPoint = pFlexgrid.FixedCols
    If pFlexgrid.MergeRow(GridRow) Then
        MsetCount = 0
        For ColLoop = 1 To ToCol
            If ColLoop < FixPoint Or ColLoop > (FromCol + 1) Then
                If PrintMerge.MergeCols Then
                    MergeVert = False
                    If ColLoop > LBound(CCells(), 2) And ColLoop <= UBound(CCells(), 2) Then
                        If CCells(GridRow, ColLoop - 1) > 0 Then
                            'for now we are going to ignore cells that are part of
                            'column groups as well - the MSHFlexGrid just mucks them up anyway
                            MergeVert = True
                        End If
                    End If
                End If
                If pFlexgrid.TextMatrix(GridRow, ColLoop - 1) = pFlexgrid.TextMatrix(GridRow, ColLoop) And Not MergeVert Then
                    If ((ColLoop - 1) < FixPoint And ColLoop < FixPoint) Or ((ColLoop - 1) >= FixPoint) Then
                        If MCells(ColLoop - 1) = 0 Then  'this is a new set
                            MsetCount = MsetCount + 1000
                            MCells(ColLoop - 1) = MsetCount + 1
                        End If
                        MCells(ColLoop) = MCells(ColLoop - 1) + 1
                    End If
                End If
            End If
        Next ColLoop
    End If
    Exit Sub
FindMergedRowErr:
    Select Case Err
        Case vbObjectError + MyE2999
            'Abort selected in a function called from this sub or function
            Exit Sub
        Case Else
            Screen.MousePointer = vbDefault
            errorResponse = MsgBox("Unexpected Error Occured." & vbLf & "Please note details." & vbLf & Error(Err) & vbLf & "in Sub:FindMergedRow", vbExclamation + vbAbortRetryIgnore, "Unexpected Error")
            Select Case errorResponse
                Case vbAbort
                Case vbRetry
                    Resume
                Case vbIgnore
                    Resume Next
            End Select
    End Select
    Screen.MousePointer = vbDefault
    ClassErrorCode = MyE2999
    ClassErrorMessage = Error$
    Err.Raise Number:=ClassErrorCode + vbObjectError, Source:=ClassErrorSource, Description:=ClassErrorMessage
    Exit Sub
End Sub


Private Function FontSizeOK() As Boolean
    'This routine checks the projected compressed grid font size and OKs values of 2 points or more
    Dim FontNeeded As Double
    
    FontNeeded = CalcScale(False)
    If FontNeeded < 2 Then
        FontSizeOK = False
    Else
        FontSizeOK = True
    End If
        
End Function

Public Property Let LeftMargin(ByVal sNewValue As Single)
    Dim workscale As Integer
    
    Select Case mAPIMarginUnit
        Case 1 To 8
            workscale = mAPIMarginUnit
        Case Else
            If LocaleUnits = 1 Then
                workscale = vbInches
            Else
                workscale = vbMillimeters
            End If
    End Select
    sNewValue = Printer.ScaleX(sNewValue, workscale, vbTwips)
    mMarginLeft = sNewValue
End Property
Private Function LocaleUnits() As Long
    Dim RetVal As Long, UserLCID As Long
    Dim ApiBuff As String * 4
    
    RetVal = GetLocaleInfo(USER_DEFAULT, LOCALE_IMEASURE, ApiBuff, 99)
    LocaleUnits = CLng(left$(ApiBuff, InStr(1, ApiBuff, Chr(0)) - 1))
    If LocaleUnits = 1 Then
        MarginAdjust = TwipsTI
    Else
        MarginAdjust = TwipsMM
    End If
End Function
Private Sub MarginDisplay()
    Dim MyTop As Single, MyBottom As Single, Myleft As Single, MyRight As Single
    
    Select Case UserPSize.MSystem
        Case 0  'metric
            MyTop = Printer.ScaleY(setMargin.top, vbTwips, vbMillimeters)
            MyBottom = Printer.ScaleY(setMargin.bottom, vbTwips, vbMillimeters)
            Myleft = Printer.ScaleX(setMargin.left, vbTwips, vbMillimeters)
            MyRight = Printer.ScaleX(setMargin.right, vbTwips, vbMillimeters)
            Text1(4).Text = Format(MyTop, "##0")
            Text1(5).Text = Format(MyBottom, "##0")
            Text1(6).Text = Format(Myleft, "##0")
            Text1(7).Text = Format(MyRight, "##0")
        Case 1  'inches
            MyTop = Printer.ScaleY(setMargin.top, vbTwips, vbInches)
            MyBottom = Printer.ScaleY(setMargin.bottom, vbTwips, vbInches)
            Myleft = Printer.ScaleX(setMargin.left, vbTwips, vbInches)
            MyRight = Printer.ScaleX(setMargin.right, vbTwips, vbInches)
            'Not that X or Y will make a difference above
            Text1(4).Text = Format(MyTop, "#0.00")
            Text1(5).Text = Format(MyBottom, "#0.00")
            Text1(6).Text = Format(Myleft, "#0.00")
            Text1(7).Text = Format(MyRight, "#0.00")
    End Select
    Check1(6).Enabled = False
    Check1(10).Enabled = False
    PageNumbs.underMargin = False
    PageNumbs.overMargin = False
    If setMargin.bottom > defMargin.bottom Then
        'it may be possible to print any page numbering under the user set margin
        If (setMargin.bottom - defMargin.bottom) > Printer.TextHeight("P") Then
            Check1(6).Enabled = True
            PageNumbs.underMargin = (Check1(6).Value = vbChecked)
        End If
    End If
    If setMargin.top > defMargin.top Then
        If (setMargin.top - defMargin.top) > Printer.TextHeight("P") Then
            Check1(1).Enabled = True
            PageNumbs.overMargin = (Check1(1).Value = vbChecked)
        End If
    End If
End Sub

Public Property Let MarginUnits(ByVal eNewValue As ScaleModeConstants)
    mAPIMarginUnit = eNewValue
End Property

Private Function PageCountHasChanged() As Boolean
    Dim OldpageCount As Long
    
    OldpageCount = defPagesWide * defPagesHigh
    CountPages True
    If OldpageCount <> (defPagesWide * defPagesHigh) Then
        PageCountHasChanged = True
    Else
        PageCountHasChanged = False
    End If
End Function

Private Function PaperSize() As PSize
'© Copyright Adit Limited 1998, 1999, 2000, 2001, 2002
    Dim POSize As Variant
    Dim SaveOrientation As Long

    On Error Resume Next
    SaveOrientation = Printer.Orientation
    PaperSize.Name = "Unknown"
    POSize = Printer.PaperSize
    Select Case POSize
        Case 1, 2
            PaperSize.Name = "Letter"
        Case 3
            PaperSize.Name = "Tabloid"
        Case 4
            PaperSize.Name = "Ledger"
        Case 5
            PaperSize.Name = "Legal"
        Case 6
            PaperSize.Name = "Statmnt"
        Case 7
            PaperSize.Name = "Exec"
        Case 8
            PaperSize.Name = "A3"
        Case 9, 10
            PaperSize.Name = "A4"
        Case 11
            PaperSize.Name = "A5"
        Case 12
            PaperSize.Name = "B4"
        Case 13
            PaperSize.Name = "B5"
        Case 14
            PaperSize.Name = "Folio"
        Case 15
            PaperSize.Name = "Quarto"
        Case 16
            PaperSize.Name = "10x14"
        Case 17
            PaperSize.Name = "11x17"
        Case 18
            PaperSize.Name = "Note"
        Case 19, 20, 21, 22, 23, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38
            PaperSize.Name = "Envelope"
        Case 39
            PaperSize.Name = "US Fan"
        Case 40
            PaperSize.Name = "GS Fan"
        Case 41
            PaperSize.Name = "GL Fan"
        Case 256
            PaperSize.Name = "Custom"
        Case Else
            PaperSize.Name = pText  'get any size names known to the print driver
                                    'but not recognised by windows and this routine
    End Select
    If SaveOrientation <> cdlPortrait Then
        Printer.Orientation = cdlPortrait
    End If
    PaperSize.height = Printer.height
    PaperSize.width = Printer.width
    PaperSize.MSystem = LocaleUnits
    Select Case PaperSize.MSystem
        Case 1  'USA
            PaperSize.height = Printer.ScaleY(PaperSize.height, Printer.ScaleMode, vbInches)
            PaperSize.width = Printer.ScaleX(PaperSize.width, Printer.ScaleMode, vbInches)
        Case Else
            PaperSize.height = Printer.ScaleY(PaperSize.height, Printer.ScaleMode, vbMillimeters)
            PaperSize.width = Printer.ScaleX(PaperSize.width, Printer.ScaleMode, vbMillimeters)
    End Select
    If SaveOrientation <> cdlPortrait Then
        Printer.Orientation = SaveOrientation
    End If
End Function
Private Function CalcFontFactor2(ScaleW As Double, ScaleH As Double, FonttoScale As StdFont)
'© Copyright Adit Limited 2005
'new version of sub for version 2.1
    Const RepString = "AbcdefGhijklMnopQrtsUvwxYz1234567890"
    'While it is likely there is a statistically optimal string the long one above should do
    Dim saveScaleWidth As Double, saveScaleHeight As Double
    Dim StringLen As Double, StringHeight As Double
    Dim FontSize As Single
    Dim savefont As New StdFont
    Dim ItFits As Boolean
    
    On Error GoTo FontFactorErr
    If Printer.ScaleMode <> vbTwips Then
        saveScaleWidth = Printer.ScaleWidth
        saveScaleHeight = Printer.ScaleHeight
    End If
    AppFont Printer.Font, savefont
    Printer.ScaleMode = vbTwips
    AppFont FonttoScale, Printer.Font
    StringLen = Printer.TextWidth(RepString)
    StringHeight = Printer.TextHeight(RepString)
    FontSize = Printer.FontSize
    
    ItFits = False
    Printer.ScaleWidth = ScaleW 'Set printer object to desired scale
    Printer.ScaleHeight = ScaleH
    Do Until ItFits Or FontSize <= 0.9
        If Printer.TextWidth(RepString) <= StringLen And Printer.TextHeight(RepString) <= StringHeight Then
            ItFits = True
        Else
            FontSize = FontSize - 0.1   'or whatever step
            FontSize = Round(FontSize, 2)
            Printer.Font.Size = FontSize
        End If
    Loop
    If saveScaleWidth > 0 Then
        Printer.ScaleWidth = saveScaleWidth
        Printer.ScaleHeight = saveScaleHeight
    Else
        Printer.ScaleMode = vbTwips
    End If
    FontSize = Printer.Font.Size
    AppFont savefont, Printer.Font
    CalcFontFactor2 = FontSize
    Exit Function
FontFactorErr:
    Select Case Err
        Case vbObjectError + MyE2999
            'Abort selected in a function called from this sub or function
            Exit Function
        Case Else
            Screen.MousePointer = vbDefault
            errorResponse = MsgBox("Unexpected Error Occured." & vbLf & "Please note details." & vbLf & Error(Err) & vbLf & "in Function:CalcFontFactor2", vbExclamation + vbAbortRetryIgnore, "Unexpected Error")
            Select Case errorResponse
                Case vbAbort
                Case vbRetry
                    Resume
                Case vbIgnore
                    Resume Next
            End Select
    End Select
    Screen.MousePointer = vbDefault
    ClassErrorCode = MyE2999
    ClassErrorMessage = Error$
    Err.Raise Number:=ClassErrorCode + vbObjectError, Source:=ClassErrorSource, Description:=ClassErrorMessage
    Exit Function
End Function

Private Sub LocatePaperImage()
'© Copyright Adit Limited 1998, 1999, 2000, 2001, 2002
    Dim ImageTwipsPmm As Double
    On Error GoTo P001LocPaperErr
    If PrintPageHeight > PrintPageWidth Then
        ImageTwipsPmm = MaxImageSize / PrintPageHeight
    Else
        ImageTwipsPmm = MaxImageSize / PrintPageWidth
    End If
    Image1.width = PrintPageWidth * ImageTwipsPmm
    Image1.height = PrintPageHeight * ImageTwipsPmm
    Image1.top = ((1275 - Image1.height) / 3) * 2
    Image1.left = (Frame1(4).width * 0.75) - (Image1.width / 2)
    'Which is well and good but wide paper (say a pre-printed form) is still portrait
    changingPPS = True
    POOrientation = Printer.Orientation
    Select Case POOrientation
        Case vbPRORPortrait
            Label1(8).Caption = UserPSize.Name & " Portrait"
            optPPS(0).Value = True
        Case vbPRORLandscape
            Label1(8).Caption = UserPSize.Name & " Landscape"
            optPPS(1).Value = True
        Case Else
            Label1(8).Caption = UserPSize.Name & " Portrait" 'well it has to be something
            optPPS(0).Value = True
    End Select
    changingPPS = False
    Label1(8).top = ((1275 - Label1(8).height) / 3) * 2 'STJG CHANGE
    Label1(8).left = (Image1.left + (Image1.width / 2)) - Label1(8).width / 2
    Exit Sub
P001LocPaperErr:
    Select Case Err
        Case vbObjectError + MyE2999
            'Abort selected in a function called from this sub or function
            Exit Sub
        Case Else
            Screen.MousePointer = vbDefault
            errorResponse = MsgBox("Unexpected Error Occured." & vbLf & "Please note details." & vbLf & Error(Err) & vbLf & "in Sub:P001:LocatePaperImage", vbExclamation + vbAbortRetryIgnore, "Unexpected Error")
            Select Case errorResponse
                Case vbAbort
                Case vbRetry
                    Resume
                Case vbIgnore
                    Resume Next
            End Select
    End Select
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub
Public Property Let PrintComplete(ByVal bNewValue As Boolean)
    mPrintComplete = bNewValue
End Property

Private Sub PrintMergedCols(ByVal StartYPos As Single, XOffset As Single)
    'this routine attempts to print any merged columns on a page
    'it is passed the YPosition down the page of the top row in CCells()
    Dim RowLoop As Long, ColLoop As Long, Cloop As Long, Keycell As Long
    Dim ImageRow As Long
    Dim CellContents As CellC
    Dim Ypos As Single, Xpos As Single
    Dim PXPos As Single, PYPos As Single
    Dim WorkText As String
    Dim printcolour As GColr
    Dim CellAlign As CellAlignment
    
    On Error GoTo PMCErr
    Xpos = Printer.ScaleX((setMargin.left - defMargin.left), vbTwips, Printer.ScaleMode)
    Xpos = Xpos + XOffset
    For ColLoop = LBound(CCells, 2) To UBound(CCells, 2)
        CellContents.CellWidth = ColWidths(ColLoop)
        Ypos = StartYPos
        CellContents.CellHeight = 0
        For RowLoop = LBound(CCells, 1) To UBound(CCells, 1)
            If CCells(RowLoop, ColLoop) > 0 Then
                'we have found a merged cell group
                If CellContents.CellHeight = 0 Then    'it is the first one
                    CellContents.CellHeight = RowHeights(RowLoop)
                    PYPos = Ypos
                    ImageRow = RowLoop
                Else
                    If left$(Format$(CCells(RowLoop, ColLoop), "000000"), 3) = left$(Format$(CCells(RowLoop - 1, ColLoop), "000000"), 3) Then
                        CellContents.CellHeight = CellContents.CellHeight + RowHeights(RowLoop)
                        Keycell = RowLoop
                    Else
                        'it's a new group so print old one
                        GoSub PrintCell
                        CellContents.CellHeight = RowHeights(RowLoop)
                        PYPos = Ypos
                    End If
                End If
            Else
                If CellContents.CellHeight > 0 Then
                    GoSub PrintCell
                End If
            End If
            Ypos = Ypos + RowHeights(RowLoop)
        Next RowLoop
        If CellContents.CellHeight > 0 Then
            GoSub PrintCell
        End If
        Xpos = Xpos + ColWidths(ColLoop)
    Next ColLoop
    Exit Sub
PrintCell:
    PXPos = Xpos
    Printer.CurrentX = PXPos
    Printer.CurrentY = PYPos
    CellContents.GridText = pFlexgrid.TextMatrix(Keycell, ColLoop)
    CellContents = TruncateAndWrapC(CellContents)
    'cell content will now have a string or an array in it
    CellAlign = GetCellAlign(Keycell, ColLoop)
    If CellAlign = CellAligndef Then
        'Use column default
        CellAlign = ColAlign(ColLoop)
    End If
    printcolour = GetColours(Keycell, ColLoop)
    'now sort out the grid display and then do the text
    If ShowLines Then
        Printer.FillColor = printcolour.Back
        Printer.FillStyle = vbFSSolid
        Printer.ForeColor = printcolour.Grid
        If InvisibleGrid Then
            Printer.DrawStyle = vbInvisible
        End If
        Printer.Line (PXPos, PYPos)-(PXPos + CellContents.CellWidth, PYPos + CellContents.CellHeight), , B
        If LastLineColour > -1 Then 'we may need to redraw the top of the box to the previous rows grid colour
            If LastLineColour <> printcolour.Grid Then
                Printer.Line (PXPos, PYPos)-(PXPos + CellContents.CellWidth, PYPos), LastLineColour
            End If
        End If
    End If
    If mPrintCellImages Then
        pFlexgrid.Row = ImageRow 'Keycell
        pFlexgrid.Col = ColLoop
        'test to see if there is a cell picture
        If pFlexgrid.CellPicture <> 0 Then
            'There is - so we can have a go at printing it
            Picture1.Picture = pFlexgrid.CellPicture
            CellImagePrint PXPos, PYPos, CellContents.CellWidth, CellContents.CellHeight
        End If
    End If
    If CellAlign And CellAlignMiddle Then
        If (CellContents.CellHeight - CellContents.PrintHeight) >= 0 Then
            PYPos = PYPos + ((CellContents.CellHeight - CellContents.PrintHeight) / 2)
        End If
    End If
    If CellAlign And cellalignbottom Then
        If (CellContents.CellHeight - CellContents.PrintHeight) >= 0 Then
            PYPos = PYPos + (CellContents.CellHeight - CellContents.PrintHeight)
        End If
    End If
    If CellAlign And CellAlignleft Then
        PXPos = PXPos + ColSpace
    End If
    If CellAlign And CellAlignCentre Then
        PXPos = PXPos + (CellContents.CellWidth - CellContents.PrintWidth) / 2
    End If
    If CellAlign And CellAlignRight Then
        PXPos = PXPos + (CellContents.CellWidth - CellContents.PrintWidth) - ColSpace
    End If
    If CellContents.WrapCount > 0 Then
        Printer.CurrentY = PYPos
        For Cloop = 0 To (CellContents.WrapCount - 1)
            Printer.CurrentX = PXPos
            Printer.Print CellContents.WrapLines(Cloop);
            If Cloop < (CellContents.WrapCount - 1) Then
                Printer.Print vbLf;
            End If
        Next Cloop
    Else
        Printer.CurrentX = PXPos
        Printer.CurrentY = PYPos
        Printer.Print CellContents.PrintLine;
    End If
    CellContents.CellHeight = 0
Return
PMCErr:
    Select Case Err
        Case vbObjectError + MyE2999
            'Abort selected in a function called from this sub or function
            Exit Sub
        Case Else
            Screen.MousePointer = vbDefault
            errorResponse = MsgBox("Unexpected Error Occured." & vbLf & "Please note details." & vbLf & Error(Err) & vbLf & "in Sub:PrintMergedCols", vbExclamation + vbAbortRetryIgnore, "Unexpected Error")
            Select Case errorResponse
                Case vbAbort
                Case vbRetry
                    Resume
                Case vbIgnore
                    Resume Next
            End Select
    End Select
    Screen.MousePointer = vbDefault
    'ClassErrorCode = MyE2999
    'ClassErrorMessage = Error$
    'Err.Raise Number:=ClassErrorCode + vbObjectError, Source:=ClassErrorSource, Description:=ClassErrorMessage
    Exit Sub
End Sub

Private Sub ReSetForm()
    'This resets the form appearance for each PrintGrid request
    Dim Combo1Index As Integer
    
    On Error Resume Next 'to make sure that the boolean InitialisingForm is reset at the end
    InitialisingForm = True
    LoadPrinters    'loads the printer list and sets the application default if set by the programmer
    CurrentFrame = 0
    Frame1(CurrentFrame).left = TabStrip1.left
    Frame1(CurrentFrame).top = TabStrip1.height / 2
    Option1(0).Value = True 'this is the default
    If mMultiColumnPrint And MC_SET_AS_DEFAULT Then
        Option1(2).Value = True
    End If
    
    If mSetPages And mAllowDialogue Then    'show the default compression set by the programmer
        Option1(1).Value = True
        Text1(0).Text = CStr(mRequestPagesWide)
        Text1(1).Text = CStr(mRequestPagesHigh)
        userPagesWide = mRequestPagesWide
        userPagesHigh = mRequestPagesHigh
    End If
    Frame1(CurrentFrame).Visible = True
    Line1(1).Visible = False
    Option1(2).Visible = False
    If mSelectionPrintRule = 0 Then 'Another New Feature not set so default to original behaviour
        mSelectionPrintRule = PS_MULTICELL
    End If
    If mPrinterOrientation > 0 Then
        Printer.Orientation = mPrinterOrientation
    End If
    If mPrintCopies < 1 Then
        mPrintCopies = 1    'not set so default value
    End If
    PaperSize.MSystem = LocaleUnits
    If mAPIMarginUnit = 0 Then  'it has not been set so decide a value in case required
        If PaperSize.MSystem = 1 Then
            mAPIMarginUnit = vbInches
        Else
            mAPIMarginUnit = vbMillimeters
        End If
    End If
    If mMarginTop + mMarginBottom + mMarginLeft + mMarginRight > 0 Then
        If mMarginTop > 0 Then
            setMargin.top = mMarginTop
        End If
        If mMarginBottom > 0 Then
            setMargin.bottom = mMarginBottom
        End If
        If mMarginLeft > 0 Then
            setMargin.left = mMarginLeft
        End If
        If mMarginRight > 0 Then
            setMargin.right = mMarginRight
        End If
        'API margins are always supplied assuming a portrait page format as this saves making yet another setting
        setMargin.width = cdlPortrait
        setMargin.hasbeenset = True
    Else
        setMargin.width = 0
        setMargin.hasbeenset = False
    End If
    SetDefmargins   'this routine checks minimum setMargin values as well
    MarginDisplay   'show the setMargin values - defaulted to defMargin
    If mTitlePages = 0 Then
        mTitlePages = TITLE_NORMAL
    End If
    If mTitlePages And USER_MAY_SET Then
        Frame1(11).Enabled = True
        Text1(3).Enabled = True
        Command1(4).Enabled = True
    Else
        Frame1(11).Enabled = False
        Text1(3).Enabled = False
        Command1(4).Enabled = False
    End If
    Option3(0).Value = True 'force the alternate values to generate a click event
    If mTitlePages And FIRST_PAGE_ONLY Then
        Option3(1).Value = True
    End If
    If mTitlePages And TITLE_ALL_PAGES Then
        Option3(2).Value = True
    End If
    If mTitlePages And TITLE_FIRST_ROW_ONLY Then
        Option3(3).Value = True
    End If
    If mTitlePages And TITLE_ALL_ROWS Then
        Option3(4).Value = True
    End If
    If pFlexgrid.FixedRows > 0 Then
        Check1(4).Visible = True
        If mTitlePages And REPEAT_COL_HEADINGS Then
            Check1(4).Value = vbChecked
            Header.IncludeColHeadings = True
        Else
            Check1(4).Value = vbUnchecked
            Header.IncludeColHeadings = False
        End If
    Else
        Check1(4).Visible = False
    End If
    FixedColData.RepeatFixed = False
    If pFlexgrid.FixedCols > 0 Then
        Check1(12).Visible = True
        If mTitlePages And REPEAT_FIXED_COLS Then
            Check1(12).Value = vbChecked
            FixedColData.RepeatFixed = True
        Else
            Check1(12).Value = vbUnchecked
        End If
    Else
        Check1(12).Visible = False
    End If
    Option5(0).Value = True 'force Header.Justify to be set
    If mTitlePages And JUSTIFY_CENTRE Then
        Option5(1).Value = True
    End If
    If mTitlePages And JUSTIFY_RIGHT Then
        Option5(2).Value = True
    End If
    If mTitlePages And FONT_BOLD Then
        'compatability option
        TitleFont.Bold = True
    End If
    If mTitlePages And FONT_UNDERLINE Then
        TitleFont.Underline = True
    End If
    'show the user what the title/font combination looks like
    Text1(3).Text = mReportTitle
    AppFont TitleFont, Text1(3).Font
    
    If mEffects = 0 Then
        mEffects = EFFECTS_NORMAL
    End If
    If mEffects And GRID_LINES Then
        Check1(2).Value = vbChecked
        ShowLines = True
        Check1(3).Enabled = True
    Else
        Check1(3).Enabled = False
        ShowLines = False
        Check1(2).Value = vbUnchecked
    End If
    If mEffects And GRID_COLOURS Then
        Check1(3).Value = vbChecked
        ShowGridColour = True
    Else
        Check1(3).Value = vbUnchecked
        ShowGridColour = False
    End If
    If mEffects And TEXT_COLOURS Then
        Check1(1).Value = vbChecked
        ShowTextColour = True
    Else
        Check1(1).Value = vbUnchecked
        ShowTextColour = False
    End If
    If mEffects And CELL_BACK_COLOURS Then
        Check1(0).Value = vbChecked
        ShowCellColour = True
        Check1(9).Enabled = False
    Else
        Check1(0).Value = vbUnchecked
        ShowCellColour = False
        Check1(9).Enabled = True
    End If
    If mEffects And COLUMN_HEAD_SHADING Then
        Check1(9).Value = vbChecked
        ShadeFixedRows = True
    Else
        Check1(9).Value = vbUnchecked
        ShadeFixedRows = False
    End If
    If mEffects And CELL_FONT_EFFECTS Then
        GridFont = True
    Else
        GridFont = False
    End If
    If mEffects And CELL_GRAPHICS Then
        Check1(13).Value = vbChecked
        mPrintCellImages = True
    Else
        Check1(13).Value = vbUnchecked
        mPrintCellImages = False
    End If
    'decide if the user has active access to these controls
    If mEffects And USER_CAN_CHANGE Then
        Frame1(8).Enabled = True
    Else
        Frame1(8).Enabled = False
    End If
    
    If mPageNumbering = 0 Then
        mPageNumbering = NUMBER_NORMAL
    End If
    Option4(0).Value = True
    If mPageNumbering And AFTER_FIRST Then
        Option4(1).Value = True
    End If
    If mPageNumbering And ALL_PAGES Then
        Option4(2).Value = True
    End If
    If mPageNumbering And INCLUDE_PAGE_COUNT Then
        Check1(5).Value = vbChecked
    Else
        Check1(5).Value = vbUnchecked
    End If
    If mPageNumbering And OVER_TOP_MARGIN Then
        Check1(10).Value = vbChecked
    Else
        Check1(10).Value = vbUnchecked
    End If
    'If there is room for a page number over the top margin then
    'MarginDisplay will already have enabled Check1(10)
    If (Check1(10).Value = vbChecked) And Check1(10).Enabled Then
        PageNumbs.overMargin = True
    Else
        PageNumbs.overMargin = False
    End If
    If mPageNumbering And UNDER_BOTTOM_MARGIN Then
        Check1(6).Value = vbChecked
    Else
        Check1(6).Value = vbUnchecked
    End If
    If (Check1(6).Value = vbChecked) And Check1(6).Enabled Then
        PageNumbs.underMargin = True
    Else
        PageNumbs.underMargin = False
    End If
    If mPageNumbering And TOP_RIGHT Then
        Combo1Index = 0
    End If
    If mPageNumbering And TOP_CENTRE Then
        Combo1Index = 1
    End If
    If mPageNumbering And TOP_LEFT Then
        Combo1Index = 2
    End If
    If mPageNumbering And BOTTOM_LEFT Then
        Combo1Index = 3
    End If
    If mPageNumbering And BOTTOM_CENTRE Then
        Combo1Index = 4
    End If
    If mPageNumbering And BOTTOM_RIGHT Then
        Combo1Index = 5
    End If
    Combo1(1).ListIndex = Combo1Index
    If mPageNumbering And USER_CAN_SET Then
        Frame1(10).Enabled = True
    Else
        Frame1(10).Enabled = False
    End If
    
    Option2(0).Value = True
    mPrintSelected = False
    
    If pFlexgrid.MergeCells > 0 Then
        Select Case pFlexgrid.MergeCells
            Case 1
                PrintMerge.MergeCols = True
                PrintMerge.MergeRows = True
            Case 2
                PrintMerge.MergeRows = True
                PrintMerge.MergeCols = False
            Case 3
                PrintMerge.MergeCols = True
                PrintMerge.MergeRows = False
            Case 4  'mind this does not work properly - but just in case
                PrintMerge.MergeCols = True
                PrintMerge.MergeRows = True
        End Select
        Frame1(14).Visible = True
        If mMergeRule = 0 Then
            mMergeRule = MERGE_NORMAL
        End If
        Option6(0).Value = True
        PrintMerge.MergeRule = 0
        If mMergeRule And MERGE_ACROSS_ROWS Then
            Option6(1).Value = True
        End If
        If mMergeRule And USER_HAS_CONTROL Then
            Frame1(14).Enabled = True
        Else
            Frame1(14).Enabled = False
        End If
        Option2(1).Enabled = False  'supports the documented problem with selection when cells are merged
    Else
        PrintMerge.MergeCols = False
        PrintMerge.MergeRows = False
        Frame1(14).Visible = False
        Option2(1).Enabled = False
        If (mSelectionPrintRule And PS_NOT_AVAILABLE) Or Not mAllowDialogue Then
            'programmer says no to print select or dialogue not shown
            'this provides the programmer with a security option and
            'ensures that selected print is not used without the dialogue
        Else
            If mSelectionPrintRule And PS_MULTICELL Then
                If pFlexgrid.Row <> pFlexgrid.RowSel Or pFlexgrid.Col <> pFlexgrid.ColSel Then
                    Option2(1).Enabled = True
                End If
            End If
            If mSelectionPrintRule And PS_MULTICOL Then
                If pFlexgrid.Col <> pFlexgrid.ColSel Then
                    Option2(1).Enabled = True
                End If
            End If
            If mSelectionPrintRule And PS_MULTIROW Then
                If pFlexgrid.Row <> pFlexgrid.RowSel Then
                    Option2(1).Enabled = True
                End If
            End If
            If (mSelectionPrintRule And PS_DEF_PRINT_SELECT) And Option2(1).Enabled Then
                Option2(1).Value = True
                mPrintSelected = True
            End If
        End If
    End If
    Text1(8).Text = ""
    Text1(8).Enabled = False
    Selectedpages = Empty
    If mShowProgress Then
        Check1(11).Value = vbChecked
    Else
        Check1(11).Value = vbUnchecked
    End If
    InitialisingForm = False
    SetPrintData
    
    Frame1(5).Visible = False
    Frame1(6).Visible = False
    Frame1(8).Visible = False
    Frame1(11).Visible = False
    TabStrip1.Tabs.Item(1).Selected = True
End Sub

Public Property Let RightMargin(ByVal sNewValue As Single)
    Dim workscale As Integer
    
    Select Case mAPIMarginUnit
        Case 1 To 8
            workscale = mAPIMarginUnit
        Case Else
            If LocaleUnits = 1 Then
                workscale = vbInches
            Else
                workscale = vbMillimeters
            End If
    End Select
    sNewValue = Printer.ScaleX(sNewValue, workscale, vbTwips)
    mMarginRight = sNewValue
End Property

Private Sub SetPrintData()
    Dim PScale As Integer
    Dim MFstring As String
    
    SetDefmargins   '**** Should be able to remove this line as values set elsewhere
    'Get Paper Size information
    UserPSize = PaperSize
    
    Select Case UserPSize.MSystem
        Case 1  'USA - well in fact English but MS probably do not know
            Label1(13).Caption = "In"
            PScale = vbInches
            MFstring = "#0.00"
            If UserPSize.width < 100 Then
                Label1(3).Caption = Format$(UserPSize.width, "###.00") & " in Wide"
            Else
                Label1(3).Caption = Format$(UserPSize.width, "##,###.00") & " in"
            End If
            If UserPSize.height < 100 Then
                Label1(4).Caption = Format$(UserPSize.height, "###.00") & " in High"
            Else
                Label1(4).Caption = Format$(UserPSize.height, "##,###.00") & " in"
            End If
        Case Else   'should be metric
            Label1(13).Caption = "mm"
            PScale = vbMillimeters
            MFstring = "##0"
            If UserPSize.width < 1000 Then
                Label1(3).Caption = Format$(UserPSize.width, "###.0") & " mm Wide"
            Else
                Label1(3).Caption = Format$(UserPSize.width, "##,###.0") & " mm"
            End If
            If UserPSize.height < 1000 Then
                Label1(4).Caption = Format$(UserPSize.height, "###.0") & " mm High"
            Else
                Label1(4).Caption = Format$(UserPSize.height, "##,###.0") & " mm"
            End If
    End Select
    Text1(2).Text = CStr(mPrintCopies) ' a convenient place to set this display
    StatusBar1.Panels(3).Text = "Copies: " & Text1(2).Text
    
    'Set the margins in the relevant frame using the correct measurement system
    MarginDisplay
    LocatePaperImage 'give a visual indication of the paper size and orientation

End Sub
Private Sub SetDefmargins()

    On Error Resume Next
    defMargin.left = GetDeviceCaps(Printer.hdc, PHYSICALOFFSETX)
    defMargin.left = Printer.ScaleX(defMargin.left, vbPixels, vbTwips) ' / TwipsMM
    defMargin.top = GetDeviceCaps(Printer.hdc, PHYSICALOFFSETY)
    defMargin.top = Printer.ScaleY(defMargin.top, vbPixels, vbTwips) ' / TwipsMM
    defMargin.width = Printer.ScaleWidth
    defMargin.height = Printer.ScaleHeight
    PrintPageWidth = Printer.width / TwipsMM
    PrintPageHeight = Printer.height / TwipsMM

    defMargin.right = Printer.width - (defMargin.width + defMargin.left)
    defMargin.bottom = Printer.height - (defMargin.height + defMargin.top)
    
    If setMargin.hasbeenset Then
        If setMargin.width <> POOrientation Then    'we will rotate any user settings
            'this does rather assume that all print drivers "rotate" the page in the same way hmnn...
            Select Case setMargin.width
                Case 0  'not set
                Case cdlPortrait
                    setMargin.height = setMargin.left
                    setMargin.left = setMargin.bottom
                    setMargin.bottom = setMargin.right
                    setMargin.right = setMargin.top
                    setMargin.top = setMargin.height
                Case cdlLandscape
                    setMargin.height = setMargin.bottom
                    setMargin.bottom = setMargin.left
                    setMargin.left = setMargin.top
                    setMargin.top = setMargin.right
                    setMargin.right = setMargin.height
            End Select
            setMargin.width = POOrientation     'holds it for future use
        End If
        'we need to make sure than the minimum margin values are reflected
        If setMargin.bottom < defMargin.bottom Then
            setMargin.bottom = defMargin.bottom
        End If
        If setMargin.top < defMargin.top Then
            setMargin.top = defMargin.top
        End If
        If setMargin.left < defMargin.left Then
            setMargin.left = defMargin.left
        End If
        If setMargin.right < defMargin.right Then
            setMargin.right = defMargin.right
        End If
    Else
        setMargin = defMargin
    End If
End Sub

Private Sub StatusBarPages()
    If Option2(2).Value And (UserSelectedPagesCount > 0) Then
        StatusBar1.Panels(2).Text = "Pages: " & CStr(UserSelectedPagesCount) & " of " & CStr(userPagesHigh * userPagesWide)
    Else
        If Option1(2).Value Then    'printing multiple sections on a page
            StatusBar1.Panels(2).Text = "Pages: " & CStr(mRequestSectionsHigh)
        Else    'printing normal or compressed
            StatusBar1.Panels(2).Text = "Pages: " & CStr(userPagesHigh * userPagesWide)
        End If
    End If
End Sub
Public Property Let TopMargin(ByVal sNewValue As Single)
    Dim workscale As Integer
    
    Select Case mAPIMarginUnit
        Case 1 To 8
            workscale = mAPIMarginUnit
        Case Else
            If LocaleUnits = 1 Then
                workscale = vbInches
            Else
                workscale = vbMillimeters
            End If
    End Select
    sNewValue = Printer.ScaleY(sNewValue, workscale, vbTwips)
    mMarginTop = sNewValue
End Property
Private Function TruncateAndWrapC(CellInfo As CellC) As CellC
    'This routine will return the TextString passed in a format that will print within the available cell space
    'Word Wrap is supported if the MSHFlexGrid being printed has this switched on
    Dim HeightString As String, BankIt As String, LineSoFar As String, ThisChar As String
    Dim RetString As Variant, WorkArray() As Variant
    Dim CharPos As Integer, WorkLines As Integer
    Dim CellWidth As Single, CellHeight As Single, HoldWidth As Single
    Dim textstring As String
    
    On Error GoTo TruncAndWrapCErr
    TruncateAndWrapC = CellInfo
    If CellInfo.GridText = "" Then
        TruncateAndWrapC.WrapCount = 0
        TruncateAndWrapC.PrintLine = ""
        TruncateAndWrapC.PrintHeight = 0
        TruncateAndWrapC.PrintWidth = 0
        Exit Function
    End If
    'if AlreadyUsed > 0 then this routine tries to return any remaining text
    'yet to be displayed rather than the best fit of text to the cell space
    If TruncateAndWrapC.AlreadyUsed > 0 Then
        HoldWidth = TruncateAndWrapC.CellWidth
        TruncateAndWrapC.CellWidth = TruncateAndWrapC.AlreadyUsed
    End If
    CellWidth = TruncateAndWrapC.CellWidth - ColSpace
    CellHeight = TruncateAndWrapC.CellHeight - Printer.TwipsPerPixelY * 2 'ColSpace was too much
    textstring = Trim$(TruncateAndWrapC.GridText)
    If Not pFlexgrid.WordWrap Then
        'only truncate - so fit to the width and eliminate any text following a CR/LF
        TruncateAndWrapC.WrapCount = 0
        CharPos = InStr(textstring, vbCr)
        If CharPos > 0 Then
            textstring = left$(textstring, CharPos - 1)
        End If
        CharPos = InStr(textstring, vbLf)
        If CharPos > 0 Then
            textstring = left$(textstring, CharPos - 1)
        End If
        If Printer.TextWidth(textstring) > CellWidth Then
            For CharPos = 1 To Len(textstring)
                ThisChar = Mid$(textstring, CharPos, 1)
                If Printer.TextWidth(RetString & ThisChar) <= CellWidth Then
                    RetString = RetString & ThisChar
                Else
                    Exit For
                End If
            Next CharPos
        Else
            RetString = textstring
        End If
        TruncateAndWrapC.PrintLine = RetString
    Else
        'we wrap on word breaks or programmer inserted vbCR characters
        WorkLines = -1
        For CharPos = 1 To Len(textstring)
            ThisChar = Mid$(textstring, CharPos, 1)
            Select Case ThisChar
                Case " "
                    If LineSoFar > "" Then
                        BankIt = LineSoFar  'we should get at least this mutch on this line
                    End If
                    LineSoFar = LineSoFar & ThisChar
                Case vbCr
                    'The programmer is forcing a line break (undocumented API Feature here guys)
                    WorkLines = WorkLines + 1
                    ReDim Preserve WorkArray(WorkLines)
                    WorkArray(WorkLines) = Trim(LineSoFar)
                    BankIt = ""
                    LineSoFar = ""
                Case vbLf   'skip this character vbCR marks line breaks although vbLF would at least be consistent
                Case Else
                    LineSoFar = LineSoFar & ThisChar
                    If Printer.TextWidth(LineSoFar) > CellWidth Then
                        WorkLines = WorkLines + 1
                        ReDim Preserve WorkArray(WorkLines)
                        WorkArray(WorkLines) = BankIt
                        LineSoFar = Trim$(right$(LineSoFar, (Len(LineSoFar) - Len(BankIt))))
                        BankIt = ""
                    End If
            End Select
        Next CharPos
        If Trim$(LineSoFar) > "" Then
            WorkLines = WorkLines + 1
            ReDim Preserve WorkArray(WorkLines)
            WorkArray(WorkLines) = Trim(LineSoFar)
        End If
        If WorkLines > 0 Then
            'we need to check the height out now
            TruncateAndWrapC.PrintHeight = CellHeight + 1
            Do Until TruncateAndWrapC.PrintHeight <= CellHeight Or WorkLines = 0
                HeightString = ""
                For CharPos = 0 To WorkLines
                    HeightString = HeightString & WorkArray(CharPos)
                    If CharPos < WorkLines Then
                        HeightString = HeightString & vbCr & vbLf
                    End If
                Next CharPos
                TruncateAndWrapC.PrintHeight = Printer.TextHeight(HeightString)
                TruncateAndWrapC.PrintWidth = Printer.TextWidth(HeightString)
                If TruncateAndWrapC.PrintHeight > CellHeight Then
                    If WorkLines > 0 Then
                        WorkLines = WorkLines - 1
                        ReDim Preserve WorkArray(WorkLines)
                    End If
                End If
            Loop
        End If
        Select Case WorkLines
            Case -1
                TruncateAndWrapC.WrapCount = 0
                TruncateAndWrapC.PrintLine = ""
            Case 0
                TruncateAndWrapC.WrapCount = 0
                TruncateAndWrapC.PrintLine = WorkArray(0)
            Case Else
                TruncateAndWrapC.WrapCount = WorkLines + 1
                TruncateAndWrapC.WrapLines = WorkArray
        End Select
    End If
    If TruncateAndWrapC.WrapCount = 0 Then
        TruncateAndWrapC.PrintHeight = Printer.TextHeight(TruncateAndWrapC.PrintLine)
        TruncateAndWrapC.PrintWidth = Printer.TextWidth(TruncateAndWrapC.PrintLine)
    End If
    If TruncateAndWrapC.AlreadyUsed > 0 Then
        TruncateAndWrapC.AlreadyUsed = 0
        TruncateAndWrapC.CellWidth = HoldWidth
        If WorkLines > 0 Then
            textstring = Trim$(WorkArray(WorkLines))
            If right$(TruncateAndWrapC.GridText, Len(textstring)) <> textstring Then
                CharPos = InStr(TruncateAndWrapC.GridText, textstring)
                If CharPos > 0 Then
                    CharPos = CharPos + Len(textstring) - 1
                    TruncateAndWrapC.GridText = right$(TruncateAndWrapC.GridText, Len(TruncateAndWrapC.GridText) - CharPos)
                End If
            Else
                TruncateAndWrapC.GridText = ""  'as it has all been printed
            End If
        Else
            If Len(TruncateAndWrapC.PrintLine) < Len(TruncateAndWrapC.GridText) Then
                TruncateAndWrapC.GridText = right$(TruncateAndWrapC.GridText, Len(TruncateAndWrapC.GridText) - Len(TruncateAndWrapC.PrintLine))
            Else
                TruncateAndWrapC.GridText = ""  'nothing else to show
            End If
        End If
    End If
    Exit Function
TruncAndWrapCErr:
    Select Case Err
        Case vbObjectError + MyE2999
            'Abort selected in a function called from this sub or function
            Exit Function
        Case Else
            Screen.MousePointer = vbDefault
            errorResponse = MsgBox("Unexpected Error Occured." & vbLf & "Please note details." & vbLf & Error(Err) & vbLf & "in Function:TruncateAndWrapC", vbExclamation + vbAbortRetryIgnore, "Unexpected Error")
            Select Case errorResponse
                Case vbAbort
                Case vbRetry
                    Resume
                Case vbIgnore
                    Resume Next
            End Select
    End Select
    Screen.MousePointer = vbDefault
    ClassErrorCode = MyE2999
    ClassErrorMessage = Error$
    Err.Raise Number:=ClassErrorCode + vbObjectError, Source:=ClassErrorSource, Description:=ClassErrorMessage
    Exit Function

End Function


Private Sub Check1_Click(Index As Integer)
'© Copyright Adit Limited 1998, 1999, 2000, 2001, 2002
    Dim CheckOn As Boolean
    If InitialisingForm Then
        Exit Sub
    End If
    CheckOn = (Check1(Index).Value = vbChecked)
    Select Case Index
        Case 0  'Show background colours (if available)
            ShowCellColour = CheckOn
            Check1(9).Enabled = Not ShowCellColour
            ShadeFixedRows = False
            If Check1(0).Enabled And CheckOn Then
                ShadeFixedRows = True
            End If
        Case 1  'Show colours in text (if available)
            ShowTextColour = CheckOn
        Case 2  'Show Grid Lines
            ShowLines = CheckOn
            Check1(3).Enabled = ShowLines
            ShowGridColour = False
            If Check1(3).Enabled And (Check1(3).Value = vbChecked) Then
                ShowGridColour = True
            End If
        Case 3
            ShowGridColour = CheckOn
        Case 4  'Repeat column headings on all pages
            Header.IncludeColHeadings = CheckOn
            CountPages True
        Case 5
            PageNumbs.incPCount = CheckOn
        Case 6
            PageNumbs.underMargin = CheckOn
            Check1(10).Value = vbUnchecked
        Case 7
            Header.Underlined = CheckOn
        Case 8
            Header.Bold = CheckOn
        Case 9
            ShadeFixedRows = CheckOn
        Case 10
            PageNumbs.overMargin = CheckOn
            Check1(6).Value = vbUnchecked
        Case 11
            mShowProgress = CheckOn
        Case 12
            FixedColData.RepeatFixed = CheckOn
            CountPages True
        Case 13
            mPrintCellImages = CheckOn
            'I think that if the users select cell graphics then they should also get
            If mPrintCellImages Then    'proportional compression for the best visual result
                mProportionalCompression = True
            End If
    End Select
End Sub



Private Sub Combo1_Click(Index As Integer)
    Select Case Index
        Case 0
            If Combo1(Index).ListIndex > -1 Then
                Set Printer = Printers(Combo1(Index).ListIndex)
                LoadPaperSize
                StatusBar1.Panels(1).Text = Combo1(0).Text
                CountPages True
                SetPrintData
            End If
        Case 1
            PageNumbs.numPos = Combo1(Index).ListIndex
    End Select
End Sub

Private Sub Command1_Click(Index As Integer)

    On Error GoTo P001Com1Clickerr
    Select Case Index
        Case 0      'OK
            If CheckNeeded Then
                If PageCountHasChanged Then
                    title = "Page Count Change"
                    Msg = "Your last action resulted in a change to the"
                    Msg = Msg & vbLf & "total number of pages to be printed."
                    Msg = Msg & vbLf & "Please check your settings before printing"
                    MsgBox Msg, vbInformation + vbOKOnly, title
                End If
                CheckNeeded = False
            End If
            ScalePrint = False
            If Option1(1).Value Then    'fit to page selected
                If userPagesWide < defPagesWide Then
                    ScalePrint = True
                End If
                If userPagesHigh < defPagesHigh Then
                    ScalePrint = True
                End If
                If ScalePrint Then
                    If Not FontSizeOK Then
                        Msg = "You have elected to scale the print. Unfortunately the"
                        Msg = Msg & vbLf & "number of pages you have chosen is too few to allow"
                        Msg = Msg & vbLf & "the text to be visible."
                        Msg = Msg & vbLf & "Please adjust the number of pages before proceeding."
                        title = "Text Font Too Small To Read"
                        MsgBox Msg, vbInformation + vbOKOnly, title
                        Exit Sub
                    End If
                End If
            End If
            Me.Hide
        Case 1          'cancel print
            CancelSheetPrint = True
            Me.Hide
        Case 2
            'Call Print Dialogue from PrintSetUp
            PrintSetUp
        Case 3  'Default on the margin set panel
            setMargin = defMargin
            MarginDisplay
        Case 4  'set the title font
            LoadTitleFont
    End Select
    Exit Sub
P001Com1Clickerr:
    'Error Recovery - no Database
    Select Case Err
        Case 484
            Resume Next
        Case vbObjectError + MyE2999
            'Abort selected in a function called from this sub or function
            Exit Sub
        Case Else
            Screen.MousePointer = vbDefault
            errorResponse = MsgBox("Unexpected Error Occured." & vbLf & "Please note details." & vbLf & Error(Err) & vbLf & "in Sub:P001:Command1_Click", vbExclamation + vbAbortRetryIgnore, "Unexpected Error")
            Select Case errorResponse
                Case vbAbort
                Case vbRetry
                    Resume
                Case vbIgnore
                    Resume Next
            End Select
    End Select
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub
Private Sub PrintSetUp()
'© Copyright Adit Limited 1998, 1999, 2000, 2001, 2002
    On Error GoTo DiagCError
    Printer.TrackDefault = True
    CommonDialog1.CancelError = True
    CommonDialog1.Copies = Text1(2).Text
    CommonDialog1.Flags = cdlPDPrintSetup
    CommonDialog1.ShowPrinter
    POOrientation = CommonDialog1.Orientation
    Printer.Orientation = POOrientation
    mPrintCopies = CommonDialog1.Copies
    SetDefmargins
    CountPages True 'could be affected by printer, paper or orientation change
    LoadPaperSize 'could have changed printers
    Combo1(0).Text = Printer.DeviceName
    StatusBar1.Panels(1).Text = Combo1(0).Text
    SetPrintData
DiagCError:
    Select Case Err
        Case 484
            Resume Next
        Case Else
            Exit Sub
    End Select
    Exit Sub
End Sub


Private Sub Form_Load()

    Combo1(1).AddItem "Top Right"
    Combo1(1).AddItem "Top Centre"
    Combo1(1).AddItem "Top Left"
    Combo1(1).AddItem "Bottom Left"
    Combo1(1).AddItem "Bottom Centre"
    Combo1(1).AddItem "Bottom Right"
    Combo1(1).ListIndex = 0
    Frame1(5).Visible = False
    Frame1(6).Visible = False
    Frame1(8).Visible = False
    Frame1(11).Visible = False
End Sub
Private Sub LoadPrinters()
'© Copyright Adit Limited 1998, 1999, 2000, 2001, 2002
    Dim pr As Printer
    Dim PrLoop As Long, PrCount As Long
    Dim PortList() As Variant
    
    On Error Resume Next
    Combo1(0).Clear         '*****this is the point the Form Load event occurs if not triggered via the API
    PrCount = Printers.Count - 1
    ReDim PortList(PrCount)
    For Each pr In Printers
        Combo1(0).AddItem pr.DeviceName
        PortList(Combo1(0).NewIndex) = UCase(Trim(pr.Port))
    Next pr
    Select Case VarType(mSetPrinter)
        Case 2, 3
            'the printer number in the printers collection may have been set
            If mSetPrinter <= (Combo1(0).ListCount - 1) And mSetPrinter >= 0 Then
                Set Printer = Printers(mSetPrinter)
            End If
        Case 8
            'perhaps a printer name has been set or the printer port (network designation)
            For PrLoop = 0 To (Combo1(0).ListCount - 1)
                If UCase(Trim(mSetPrinter)) = UCase(Trim(Combo1(0).List(PrLoop))) Then
                    Set Printer = Printers(PrLoop)
                Else
                    If UCase(Trim(mSetPrinter)) = PortList(PrLoop) Then
                        Set Printer = Printers(PrLoop)
                    End If
                End If
            Next PrLoop
    End Select
    Combo1(0).Text = Printer.DeviceName    'of course it is possible there isnt one
    LoadPaperSize
    StatusBar1.Panels(1).Text = Combo1(0).Text
    POOrientation = Printer.Orientation
End Sub
Public Function NumericOnly(KeyAscii As Integer, Optional extrachar As Variant) As Integer
    On Error GoTo NumericOnlyErr
    Select Case Chr$(KeyAscii)
        Case "0" To "9", Chr$(vbKeyBack)
            NumericOnly = KeyAscii
        Case Else
            NumericOnly = 0
            If Not IsMissing(extrachar) Then
                If InStr(extrachar, Chr$(KeyAscii)) Then
                    NumericOnly = KeyAscii
                End If
            End If
    End Select
    Exit Function
NumericOnlyErr:
    NumericOnly = 0
    Exit Function
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set TitleFont = Nothing
End Sub

Private Sub Form_Resize()

    If FormDisplayMode = DISPLAY_PROGRESS Then
        Me.height = 1200
        Me.width = 2500
        Frame1(CurrentFrame).Visible = False
        Command1(0).Visible = False
        Command1(1).Visible = False
        StatusBar1.Visible = False
        TabStrip1.Visible = False
        Me.Caption = "Print Progress"
        Frame1(1).top = (Me.ScaleHeight - Frame1(1).height) / 2
        Frame1(1).left = (Me.ScaleWidth - Frame1(1).width) / 2
        Label2.Caption = ""
        ProgressBar1.Value = 0
        Frame1(1).Visible = True
    Else
        StatusBar1.Visible = True
        TabStrip1.Visible = True
        Command1(0).Visible = True
        Command1(1).Visible = True
        Frame1(1).Visible = False
        Me.Caption = "Print Options"
        TabStrip1.left = Me.ScaleLeft
        TabStrip1.top = Me.ScaleTop
        Me.width = TabStrip1.width
        Me.height = 4860
        Check1(11).top = 3615
        Check1(11).left = 135
        Command1(0).top = 3570
        Command1(1).top = Command1(0).top
        Command1(0).left = 3810
        Command1(1).left = 2850
    End If
End Sub


Private Sub Image1_Click()
    On Error Resume Next
    
    If changingPPS Then
        Exit Sub
    Else
        changingPPS = True
    End If
    If POOrientation = cdlLandscape Then
        POOrientation = cdlPortrait
    Else
        POOrientation = cdlLandscape
    End If
    Printer.Orientation = POOrientation
    DoEvents
    SetDefmargins
    CountPages True 'number of print pages is likely to change
    SetPrintData
    changingPPS = False
End Sub
Private Sub CountPages(ShowPages As Boolean)
'© Copyright Adit Limited 1998 to 2005
    Dim PrinterWidth As Double, Printerheight As Double, Thispage As Double
    Dim FixedRows As Long, lPagesWide As Long, lPagesHigh As Long, GridIn As Long
    Dim SPagesHigh As Long, ColLoop As Long
    Dim ThisSection As Integer
    Dim FromRow As Long, ToRow As Long, FromCol As Long, ToCol As Long
    Dim selGrid() As Variant
    Dim FitsPage As Boolean
    
    On Error GoTo CountPagesErr
    If InitialisingForm Then    'ignore odd click events while things are set up
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    mRequestSectionsWide = 0    'Initialise
    If Header.IncludeColHeadings Then
        FixedRows = pFlexgrid.FixedRows
    End If
    CalcFootHeight  'also sets a minimum header height for page numbering if appropriate
    Header = ReadHeadings(Header, FixedRows)
    CalcFixedWidth
    If defMargin.width = 0 Then
        SetDefmargins
    End If
    PrinterWidth = Printer.ScaleX(defMargin.width - (setMargin.left + setMargin.right), vbTwips, Printer.ScaleMode)
    Printerheight = Printer.ScaleY(defMargin.height - (setMargin.top + setMargin.bottom), vbTwips, Printer.ScaleMode) - PageNumbs.footHeight
    If mPrintSelected Then
        GridHeight = GetSelectedHeight
        GridWidth = GetSelectedWidth
        selGrid() = DefSavePosition
        FromRow = selGrid(0)
        ToRow = selGrid(2)
        FromCol = selGrid(1)
        ToCol = selGrid(3)
    Else
        FromRow = 0
        ToRow = UBound(RowHeights())
        FromCol = 0
        ToCol = UBound(ColWidths())
    End If
    If mPLColEnd > 0 Or mPLRowEnd > 0 Then
        'the programmer has restricted the overall print area via the API
        If FromRow < mPLRowStart Then
            FromRow = mPLRowStart
        End If
        If FromCol < mPLColStart Then
            FromCol = mPLColStart
        End If
        If ToRow > mPLRowEnd Then
            ToRow = mPLRowEnd
        End If
        If ToCol > mPLColEnd Then
            ToCol = mPLColEnd
        End If
    End If
    GridIn = ToCol
    'Bug fix for grids with a terminal sequence of zero width columns on a page border
    'thanks to Ken Hanson
    For ColLoop = GridIn To FromCol Step -1
        If ColWidths(ColLoop) > 0 Then
            Exit For
        Else
            ToCol = ColLoop - 1
        End If
    Next ColLoop
    'quick check for grids that will not break down into pages
    FitsPage = GridFitsPage() ' the user will get error messages if the dialogue is visible
    If Not FitsPage Then
        'we cant print it and the code below will have problems
        If Not ShowPages Then
            CancelSheetPrint = True
        Else
            defPagesHigh = 0
            defPagesWide = 0
            DisplayPageCount
            DoEvents
        End If
        Exit Sub
    End If
    
    lPagesWide = 1
    GridIn = 0
    ReDim PagesW(lPagesWide)
    PagesW(lPagesWide).GridStart = FromCol
    If (GridWidth + FixedColData.FixedPage1) > PrinterWidth Then
        GridIn = FromCol - 1
        Thispage = FixedColData.FixedPage1
        Do Until GridIn = ToCol
            If Thispage + ColWidths(GridIn + 1) > PrinterWidth Then
                PagesW(lPagesWide).GridEnd = GridIn
                PagesW(lPagesWide).Size = Thispage   'used for scaling to fit pages
                lPagesWide = lPagesWide + 1
                ReDim Preserve PagesW(lPagesWide)
                PagesW(lPagesWide).GridStart = GridIn + 1
                Thispage = FixedColData.FixedWidth
            Else
                GridIn = GridIn + 1
                Thispage = Thispage + ColWidths(GridIn)
            End If
        Loop
    Else
        mRequestSectionsWide = Int(PrinterWidth / (GridWidth + SECTION_GAP + FixedColData.FixedPage1))
        If mRequestSectionsWide < 1 Then
            mRequestSectionsWide = 1
        End If
        'this registers the potential for printing multiple sections
        'so now we can count the sections and pages required
        If mRequestSectionsWide > 1 Then
            SPagesHigh = 1
            ThisSection = 1
            ReDim PagesS(mRequestSectionsWide, SPagesHigh)
            PagesS(ThisSection, SPagesHigh).GridStart = FromRow
            GridIn = FromRow - 1
            Thispage = Header.Page1Height
            Do Until GridIn = ToRow
                If Thispage + RowHeights(GridIn + 1) > Printerheight Then
                    'start a new section or page
                    PagesS(ThisSection, SPagesHigh).GridEnd = GridIn
                    ThisSection = ThisSection + 1
                    If ThisSection > mRequestSectionsWide Then
                        'new page this time then
                        SPagesHigh = SPagesHigh + 1
                        ReDim Preserve PagesS(mRequestSectionsWide, SPagesHigh)
                        ThisSection = 1
                    End If
                    PagesS(ThisSection, SPagesHigh).GridStart = GridIn + 1
                    If SPagesHigh > 1 Then
                        Thispage = Header.NextPage
                    Else
                        Thispage = Header.Page1Height
                        If Header.IncludeColHeadings And Not Header.CHinPage1 Then
                            'if column headings need repeating but the heading space has not
                            'allowed for them then we have to add the required height
                            Thispage = Thispage + Header.ColHeadHeight
                        End If
                    End If
                Else
                    GridIn = GridIn + 1
                    Thispage = Thispage + RowHeights(GridIn)
                End If
            Loop
            PagesS(ThisSection, SPagesHigh).GridEnd = ToRow
            If SPagesHigh = 1 Then  'we may not need all of the sections that would fit on a page
                mRequestSectionsWide = ThisSection
            End If
            mRequestSectionsHigh = SPagesHigh
        End If
        'Back to the main count of pages wide
        Thispage = GridWidth
    End If
    PagesW(lPagesWide).GridEnd = ToCol
    PagesW(lPagesWide).Size = Thispage
    lPagesHigh = 1
    ReDim PagesH(lPagesHigh)
    PagesH(lPagesHigh).GridStart = FromRow
    If (GridHeight + Header.Page1Height) > Printerheight Then
        GridIn = FromRow - 1
        Thispage = Header.Page1Height
        Do Until GridIn = ToRow
            If Thispage + RowHeights(GridIn + 1) > Printerheight Then
                PagesH(lPagesHigh).GridEnd = GridIn
                PagesH(lPagesHigh).Size = Thispage
                lPagesHigh = lPagesHigh + 1
                ReDim Preserve PagesH(lPagesHigh)
                PagesH(lPagesHigh).GridStart = GridIn + 1
                Thispage = Header.NextPage
            Else
                GridIn = GridIn + 1
                Thispage = Thispage + RowHeights(GridIn)
            End If
        Loop
    Else
        Thispage = GridHeight + Header.Page1Height
    End If
    PagesH(lPagesHigh).GridEnd = ToRow
    PagesH(lPagesHigh).Size = Thispage
    defPagesHigh = lPagesHigh
    defPagesWide = lPagesWide
    If Me.Visible Or ShowPages Then   'force them out the first time
        DisplayPageCount             'but don't bother later if the form
        DoEvents                        'is unloaded
    End If
    CheckNeeded = False     'the form should now show the page count correctly
    Screen.MousePointer = vbDefault
    Exit Sub
CountPagesErr:
    Select Case Err
        Case vbObjectError + MyE2999
            'Abort selected in a function called from this sub or function
            Exit Sub
        Case Else
            Screen.MousePointer = vbDefault
            errorResponse = MsgBox("Unexpected Error Occured." & vbLf & "Please note details." & vbLf & Error(Err) & vbLf & "in Sub:CountPages", vbExclamation + vbAbortRetryIgnore, "Unexpected Error")
            Select Case errorResponse
                Case vbAbort
                Case vbRetry
                    Resume
                Case vbIgnore
                    Resume Next
            End Select
    End Select
    Screen.MousePointer = vbDefault
    ClassErrorCode = MyE2999
    ClassErrorMessage = Error$
    Err.Raise Number:=ClassErrorCode + vbObjectError, Source:=ClassErrorSource, Description:=ClassErrorMessage
    Exit Sub
End Sub
Private Sub CalcFootHeight()
    'This sub decides if any space is required for the page number
    'at the foot of the page (within the margins set by the user)
    On Error Resume Next
    
    PageNumbs.footHeight = 0
    PageNumbs.minimumTop = 0
    Select Case PageNumbs.numOption
        Case 0
        Case Else
            If PageNumbs.numPos >= 3 Then 'it is at the bottom
                If Not PageNumbs.underMargin Then 'we need some space
                    PageNumbs.footHeight = Printer.TextHeight("P") + TwipsMM
                End If
            Else
                If Not PageNumbs.overMargin Then
                    PageNumbs.minimumTop = Printer.TextHeight("P") + TwipsMM
                End If
            End If
    End Select
End Sub
Private Function ReadHeadings(OldHeading As Heads, RepFixed As Long) As Heads
'© Copyright Adit Limited 1998 to 2005
    Dim MyHead As Heads
    Dim FixLoop As Long
    Dim ScaleH As Double, ScaleW As Double
    Dim savefont As New StdFont
    Dim Scaleset As Boolean, addSubHeight As Boolean
    Dim GridVals() As Variant, workVar() As Variant
    Dim Workstring As String
    
    On Error GoTo ReadHeadingsErr
    MyHead = OldHeading
    If Printer.ScaleMode <> vbTwips Then
        Scaleset = True
        ScaleH = Printer.ScaleHeight
        ScaleW = Printer.ScaleWidth
        Printer.ScaleMode = vbTwips
    End If
    
    If mReportTitle > "" And Header.PrintPage > 0 Then 'ther is a title and it is found on at least the top row of pages
        AppFont Printer.Font, savefont
        AppFont TitleFont, Printer.Font
        
        MyHead.Page1Height = Printer.TextHeight("X") * 1.5
        MyHead.TitleWidth = Printer.TextWidth(mReportTitle)
        AppFont savefont, Printer.Font
    Else
        MyHead.Page1Height = 0
    End If
    
    If Not IsEmpty(mySubHeadings) Then
        'analyse SubTitleUsage ? to make sure we are printing at all
        Select Case mRepeatSubTitle
            Case SubTitleUsage.SUB_AS_MAIN_TITLE
                If Header.PrintPage > 0 Then
                    addSubHeight = True
                Else
                    addSubHeight = False
                End If
            Case Else
                addSubHeight = True
        End Select
        If addSubHeight Then
            workVar() = mySubHeadings
            AppFont mySubHeadFont, Printer.Font
            For FixLoop = 0 To UBound(workVar())
                If FixLoop > 0 Then
                    Workstring = Workstring & vbLf
                End If
                Workstring = Workstring & workVar(FixLoop)
            Next FixLoop
            MyHead.Page1Height = MyHead.Page1Height + Printer.TextHeight(Workstring)
            AppFont savefont, Printer.Font
        End If
    End If
    
    If MyHead.Page1Height < PageNumbs.minimumTop Then
        MyHead.Page1Height = PageNumbs.minimumTop
    End If
    
    'If we are printing a selected area and the user wants column headings we may
    'have to make an allowance for the required space on the first page as well as later ones
    AddHeadingsToP1 = False
    If mPrintSelected And RepFixed > 0 Then
        GridVals() = DefSavePosition
        If GridVals(0) > 0 Or GridVals(2) < RepFixed Then 'the heading rows are not currently in the selected area
            AddHeadingsToP1 = True
        End If
    End If
    If MyHead.PrintPage > 1 Then
        MyHead.NextPage = MyHead.Page1Height
    Else
        MyHead.NextPage = 0
    End If
    MyHead.NextTitle = MyHead.NextPage
    MyHead.FixedCols = RepFixed
    MyHead.ColHeadHeight = 0
    MyHead.CHinPage1 = False
    If RepFixed > 0 Then
        For FixLoop = 0 To RepFixed - 1
            MyHead.ColHeadHeight = MyHead.ColHeadHeight + RowHeights(FixLoop)
        Next FixLoop
        MyHead.NextPage = MyHead.NextPage + MyHead.ColHeadHeight
        If AddHeadingsToP1 Then
            MyHead.Page1Height = MyHead.Page1Height + MyHead.ColHeadHeight
            MyHead.CHinPage1 = True
        End If
    End If
    If Scaleset Then
        'reset the scale to the compressed values
        Printer.ScaleHeight = ScaleH
        Printer.ScaleWidth = ScaleW
    End If
    ReadHeadings = MyHead
    Set savefont = Nothing
    Exit Function
ReadHeadingsErr:
    Select Case Err
        Case 484
            Resume Next
        Case vbObjectError + MyE2999
            'Abort selected in a function called from this sub or function
            Exit Function
        Case Else
            Screen.MousePointer = vbDefault
            errorResponse = MsgBox("Unexpected Error Occured." & vbLf & "Please note details." & vbLf & Error(Err) & vbLf & "in Function:ReadHeadings", vbExclamation + vbAbortRetryIgnore, "Unexpected Error")
            Select Case errorResponse
                Case vbAbort
                Case vbRetry
                    Resume
                Case vbIgnore
                    Resume Next
            End Select
    End Select
    Screen.MousePointer = vbDefault
    ClassErrorCode = MyE2999
    ClassErrorMessage = Error$
    Err.Raise Number:=ClassErrorCode + vbObjectError, Source:=ClassErrorSource, Description:=ClassErrorMessage
    Exit Function
End Function
Public Sub PrintGrid(Flexgrid As MSHFlexGrid, ByVal NumCopies As Integer, ByVal PrintTitle As String, ByVal DefSettings As PrintSettings)
    '© Copyright Adit Limited 1998 to 2005
    'This is the version 1 compatible PrintGrid interface
    Set pFlexgrid = Flexgrid
    PrintStarted = False
    mPrintComplete = False
    SetDefValues DefSettings, PrintTitle, NumCopies
    PrintGridAPI    'join the new track after setting up the values indicated in DefSettings
End Sub
Public Function PrintGridAPI() As Long
    '© Copyright Adit Limited 1998 to 2005
    'This is the main PrintGrid API Method
    Dim FontFactor As Double
    Dim pFont As New StdFont
    Dim PageCount As Long, PageLoop As Long, RowLoop As Long, SectionLoop As Long
    Dim RetVar As Variant
    Dim PrintPage As Boolean
    Dim PageH As Long, PageW As Long
    Dim SectionXOffset As Single, PYPos As Single, saveNewY As Single
    
    On Error GoTo NoPrinterFErr
    PrintStarted = False
    'the next two values should have been pre-set but this is a safty strap
    mPrintComplete = False
    mPrintProgress = 0
    
    Printer.ScaleMode = vbTwips 'Checks there is a printer somewhere - otherwise not much point in all this
    DoEvents
    On Error GoTo PrintGridAPIErr
    Screen.MousePointer = vbHourglass
    'save printer font - just to be polite
    AppFont Printer.Font, pFont
    'Set Printer Font to the Grid Font
    AppFont pFlexgrid.Font, Printer.Font
    Printer.DrawWidth = 1
    SaveGridPosition
    ReSetForm
    GridWidth = GetColWidths
    GridHeight = GetRowHeights
    CancelSheetPrint = False
    If mAllowDialogue Then
        CountPages True 'this may move to the form load equivalent
    Else
        CountPages False 'we need to know if the grid can be printed
        If mSetPages Then
            userPagesHigh = mRequestPagesHigh
            userPagesWide = mRequestPagesWide
            ScalePrint = True
            If Not FontSizeOK Then
                CancelSheetPrint = True
                mSetPages = False
            End If
        Else
            userPagesHigh = defPagesHigh
            userPagesWide = defPagesWide
        End If
    End If
    If mAllowDialogue Then 'we will display options to the user
        Screen.MousePointer = vbDefault
        mSetPages = False
        Me.Show vbModal
    End If
    If CancelSheetPrint Then
        DoEvents
        ReSetGridPosition
        Set pFlexgrid = Nothing
        mPrintComplete = True
        Exit Function
    End If
        
    Screen.MousePointer = vbHourglass
    Printer.Copies = mPrintCopies
    InvisibleGrid = False
    If ShadeFixedRows And Not ShowLines Then
        ShowLines = True
        InvisibleGrid = True
    End If
    ColSpace = COLSPACEMM * TwipsMM
    FontFactor = 1
    If ScalePrint Then
        FontFactor = CalcScale(True)
        'calculates the font factor and sets the scale on the printer
        'plus calls CountPages to recalculate the columns and rows on each
        'page to be printed
    End If
    Printer.FontTransparent = True  'token as this has not worked properly since VB3
    If PrintMultiGrids Then
        PageCount = mRequestSectionsHigh
        If PageCount = 0 Then
            PageCount = userPagesHigh * userPagesWide
            PrintMultiGrids = False
        End If
    Else
        PageCount = userPagesHigh * userPagesWide
    End If
    ColumnAlignment 'sets the default alignments in ColAlign()
    If ScalePrint Then
        Printer.FontSize = FontFactor
        Printer.FontName = Printer.FontName
        Printer.FontSize = FontFactor
    End If
    pFlexgrid.Redraw = False
    If mShowProgress Then
        'Change the window to the progress display layout
        FormDisplayMode = DISPLAY_PROGRESS
        Form_Resize
        ProgressBar1.max = PageCount
        ProgressBarDisplayed = True
        'and then show this window in a non-interactive form
        ShowWindow Me.hwnd, SW_SHOWNOACTIVATE
    End If
    DoEvents
    For PageLoop = 1 To PageCount
        PrintPage = StartNewPage(PageLoop)
        If PrintPage Then
            If mShowProgress Then
                ProgressBar1.Value = PageLoop
                Label2.Caption = "Printing Page " & CStr(PageLoop) & " of " & CStr(PageCount)
                Label2.Refresh
            End If
            'output the relevant cells
            If userPagesWide > 1 Then
                PageH = Int(PageLoop / userPagesWide)
                PageW = PageLoop Mod defPagesWide
                If PageW = 0 Then
                    PageW = userPagesWide
                Else
                    PageH = PageH + 1
                End If
            Else
                PageW = 1
                PageH = PageLoop
            End If
            If PrintMultiGrids Then
                'This is a big block of extra code which rather jumps around the print
                'process. If just reviewing things - skip to the Else and ignore this
                For SectionLoop = 1 To mRequestSectionsWide
                    Printer.CurrentY = CurrentPageTop
                    SectionXOffset = (GridWidth + SECTION_GAP + FixedColData.FixedPage1) * (SectionLoop - 1)
                    
                    If PagesS(SectionLoop, PageH).GridEnd > 0 Then
                        If PageH > 1 Then
                            PYPos = Header.NextPage + Printer.ScaleY((setMargin.top - defMargin.top), vbTwips, Printer.ScaleMode)
                        Else
                            PYPos = Header.Page1Height + Printer.ScaleY((setMargin.top - defMargin.top), vbTwips, Printer.ScaleMode)
                        End If
                        If (SectionLoop > 1 And Header.IncludeColHeadings) Or AddHeadingsToP1 Then
                            RepeatColHeadings 1, PYPos, SectionXOffset
                            If PageH = 1 Then
                                PYPos = PYPos + Header.ColHeadHeight
                            End If
                        End If
                        
                        If PrintMerge.MergeCols Then    'go see if there are any in the fixed rows
                            saveNewY = Printer.CurrentY
                            FindMergedCol PagesS(SectionLoop, PageH).GridStart, PagesS(SectionLoop, PageH).GridEnd, PagesW(1).GridStart, PagesW(1).GridEnd
                            PrintMergedCols PYPos, SectionXOffset
                            Printer.CurrentY = saveNewY
                        End If
                        
                        For RowLoop = PagesS(SectionLoop, PageH).GridStart To PagesS(SectionLoop, PageH).GridEnd
                            If RowHeights(RowLoop) > 0 Then
                                PrintRowOnPage RowLoop, PageW, SectionXOffset
                                Printer.CurrentY = Printer.CurrentY + RowHeights(RowLoop)
                            End If
                        Next RowLoop
                    End If
                Next SectionLoop
            Else
                SectionXOffset = 0
                For RowLoop = PagesH(PageH).GridStart To PagesH(PageH).GridEnd
                    If RowHeights(RowLoop) > 0 Then
                        PrintRowOnPage RowLoop, PageW, SectionXOffset
                        Printer.CurrentY = Printer.CurrentY + RowHeights(RowLoop)
                    End If
                Next RowLoop
            End If
        End If
    Next PageLoop
    Printer.EndDoc  'finish off the document
    ReSetGridPosition
    pFlexgrid.Redraw = True
    Set pFlexgrid = Nothing
    AppFont pFont, Printer.Font
    Set pFont = Nothing
    If InvisibleGrid Then
        Printer.DrawStyle = vbSolid
    End If
    mySubHeadings = nullHeading
    'undo API fixed print area restriction
    mPLColStart = 0
    mPLRowStart = mPLColStart
    mPLColEnd = mPLColStart
    mPLRowEnd = mPLColStart
    If mShowProgress Then
        'hide this window after the print
        'and then re-set to the normal dialogue view
        ShowWindow Me.hwnd, SW_HIDE
        FormDisplayMode = DISPLAY_DIALOGUE
        Form_Resize
    End If
    mPrintComplete = True
    Screen.MousePointer = vbDefault
    
    Exit Function
NoPrinterFErr:
    Screen.MousePointer = vbDefault
    title = "Printer Problems"
    Msg = "This program encountered a problem when checking"
    Msg = Msg & vbLf & "on the default printer for your system."
    Msg = Msg & vbLf & "The Error Code was: " & CStr(Err)
    Msg = Msg & vbLf & Error$
    Msg = Msg & vbLf & vbLf & "This problem must be corrected before printing can continue."
    MsgBox Msg, vbCritical + vbOKOnly, title
    Exit Function
PrintGridAPIErr:
   Select Case Err
        Case 484
            Resume Next 'printer driver does not support what we just tried
        Case 401
            'Cant show the progress bar in modeless form so we will have to skip it
            Resume Next 'This should be fixed now but...
        Case vbObjectError + MyE2999
            'Abort selected in a function called from this sub or function
            'Exits below to kill any started document
        Case Else
            Screen.MousePointer = vbDefault
            errorResponse = MsgBox("Unexpected Error Occured." & vbLf & "Please note details." & vbLf & Error(Err) & vbLf & "in Sub:PrintGrid", vbExclamation + vbAbortRetryIgnore, "Unexpected Error")
            Select Case errorResponse
                Case vbAbort
                Case vbRetry
                    Resume
                Case vbIgnore
                    Resume Next
            End Select
    End Select
    If PrintStarted Then
        Printer.KillDoc 'Try and clean up any mess left
    End If
    Screen.MousePointer = vbDefault
    Exit Function
End Function
Private Sub ColumnAlignment()
'© Copyright Adit Limited 1998, 1999, 2000, 2001, 2002
    Dim ColLoop As Long, RowLoop As Long
    Dim cColAlign As CellAlignment
    Dim CheckRows As Long
    Dim NumberType As Boolean
    Dim Workstring As String
    
    On Error Resume Next
    ReDim ColAlign(pFlexgrid.Cols - 1)
    For ColLoop = 0 To pFlexgrid.Cols - 1
        Select Case pFlexgrid.ColAlignment(ColLoop)
            Case 0
                cColAlign = CellAlignleft + CellAlignTop
            Case 1
                cColAlign = CellAlignleft + CellAlignMiddle
            Case 2
                cColAlign = CellAlignleft + cellalignbottom
            Case 3
                cColAlign = CellAlignCentre + CellAlignTop
            Case 4
                cColAlign = CellAlignCentre + CellAlignMiddle
            Case 5
                cColAlign = CellAlignCentre + cellalignbottom
            Case 6
                cColAlign = CellAlignRight + CellAlignTop
            Case 7
                cColAlign = CellAlignRight + CellAlignMiddle
            Case 8
                cColAlign = CellAlignRight + cellalignbottom
            Case Else      'default so content decides
                CheckRows = pFlexgrid.Rows - 1
                NumberType = True
                For RowLoop = pFlexgrid.FixedRows To CheckRows
                    'sample the first and last 50 rows
                    If RowLoop < 51 Or (CheckRows - RowLoop) < 51 Then
                        Workstring = Trim$(pFlexgrid.TextMatrix(RowLoop, ColLoop))
                        If Not IsDate(Workstring) And Not IsNumeric(Workstring) And Workstring > "" Then
                            NumberType = False
                            Exit For
                        End If
                    End If
                Next RowLoop
                If NumberType Then
                    cColAlign = CellAlignRight + CellAlignTop
                Else
                    cColAlign = CellAlignleft + CellAlignTop
                End If
        End Select
        ColAlign(ColLoop) = cColAlign
    Next ColLoop
End Sub

Private Function StartNewPage(ByVal PageNum As Long) As Boolean
'© Copyright Adit Limited 1998 to 2005
    Dim workVar() As Variant, workVar2() As Variant
    Dim PageLoop As Integer, TitlePortion As Integer
    Dim RetVal As Boolean, PrintNumber As Boolean, PrintTitle As Boolean
    Dim PageW As Long, PageH As Long, RowLoop As Long, ColLoop As Long
    Dim SaveSize As Single, TitleSplit As Single, TitlePage As Single
    Dim PYPos As Single
    Dim PageString As String, TitleText As String
    Dim SavePFont As New StdFont
    
    On Error GoTo StartNewPageErr
    RetVal = True
    If Not IsEmpty(Selectedpages) Then
        workVar() = Selectedpages
        RetVal = False
        For PageLoop = 0 To UBound(workVar())
            If workVar(PageLoop) = PageNum Then
                RetVal = True
                Exit For
            End If
        Next PageLoop
    End If
    If RetVal Then
        LastLineColour = -1
        If PrintStarted Then
            Printer.NewPage 'we are starting a new physical page
        End If
        PrintStarted = True
        mPrintProgress = mPrintProgress + 1
        'Check which sheet position we are working on
        If userPagesWide > 1 Then
            PageH = Int(PageNum / userPagesWide)
            PageW = PageNum Mod userPagesWide
            If PageW = 0 Then
                PageW = userPagesWide
            Else
                PageH = PageH + 1
            End If
        Else
            PageW = 1
            PageH = PageNum
        End If
        'now deal with headings - options now include printing the title just once per row
        PrintTitle = False
        TitlePortion = 0    'print the lot
        TitleText = mReportTitle
        Select Case Header.PrintPage
            Case 0  'No Titles
            Case 1  'First page only
                If PageNum = 1 Then
                    PrintTitle = True
                End If
            Case 2  'Al pages
                PrintTitle = True
            Case 3, 4 'printed for rows
                Select Case Header.Justify
                    Case 0  'left
                        If PageW = 1 Then
                            PrintTitle = True
                        End If
                    Case 1  'Middle
                        TitleSplit = 2
                        TitlePage = Int(userPagesWide / TitleSplit)
                        TitleSplit = (userPagesWide / TitleSplit)
                        If TitleSplit > (TitlePage + 0.4) Then 'so who trusts FP arithmatic?
                            'odd number of pages wide
                            If PageW = TitlePage + 1 Then
                                PrintTitle = True
                            End If
                        Else
                            If PageW = TitlePage Then
                                TitlePortion = 1    'print left half
                                TitleText = left$(TitleText, Int(Len(TitleText) / 2))
                            End If
                            If PageW = TitlePage + 1 Then
                                TitlePortion = 2    'print right half
                                TitleText = right$(TitleText, Len(TitleText) - Int(Len(TitleText) / 2))
                            End If
                        End If
                    Case 2  'right Justify
                        If PageW = userPagesWide Then
                            PrintTitle = True
                        End If
                End Select
                If Header.PrintPage = 3 And PageH > 1 Then
                    PrintTitle = False
                End If
        End Select
        
        If PrintTitle Then
            AppFont Printer.Font, SavePFont
            AppFont TitleFont, Printer.Font
            If ScalePrint Then
                Printer.Font.Size = CalcFontFactor2(Printer.ScaleWidth, Printer.ScaleHeight, Printer.Font)
            End If
            
            Printer.CurrentX = Printer.ScaleX((setMargin.left - defMargin.left), vbTwips, Printer.ScaleMode)
            Select Case Header.Justify
                Case 0  'left and default
                Case 1  'centre
                    Select Case TitlePortion
                        Case 0
                            If Printer.ScaleX(defMargin.width - ((setMargin.left - defMargin.width) + (setMargin.right - defMargin.right)), vbTwips, Printer.ScaleMode) - Header.TitleWidth > 0 Then
                                Printer.CurrentX = (Printer.ScaleX(defMargin.width - ((setMargin.left - defMargin.left) + (setMargin.right - defMargin.right)), vbTwips, Printer.ScaleMode) - Header.TitleWidth) / 2
                            End If
                        Case 1  'printing the left half in the middle of a row so right justify on page
                            Printer.CurrentX = Printer.ScaleX(defMargin.width - (setMargin.right - defMargin.right), vbTwips, Printer.ScaleMode) - (Printer.TextWidth(TitleText & "XX"))
                        Case 2  'printing right half in middle of a row so left justify on page
                    End Select
                Case 2  'right
                    Printer.CurrentX = Printer.ScaleX(defMargin.width - (setMargin.right - defMargin.right), vbTwips, Printer.ScaleMode) - (Header.TitleWidth + Printer.TextWidth("XX"))
            End Select
            Printer.CurrentY = Printer.ScaleY((setMargin.top - defMargin.top), vbTwips, Printer.ScaleMode)
            Printer.Print TitleText
            AppFont SavePFont, Printer.Font
            If Not IsEmpty(mySubHeadings) Then
                workVar2() = mySubHeadings
                AppFont Printer.Font, SavePFont
                AppFont mySubHeadFont, Printer.Font
                If ScalePrint Then
                    Printer.Font.Size = CalcFontFactor2(Printer.ScaleWidth, Printer.ScaleHeight, Printer.Font)
                End If
                If mySubJustify = 0 Then
                    mySubJustify = Header.Justify
                Else
                    Select Case mySubJustify
                        Case 1, 2, 3
                            mySubJustify = mySubJustify - 1
                        Case Else
                            mySubJustify = 0
                    End Select
                End If
                For RowLoop = 0 To UBound(workVar2())
                    TitleText = workVar2(RowLoop)
                    Printer.CurrentX = Printer.ScaleX((setMargin.left - defMargin.left), vbTwips, Printer.ScaleMode)
                    Select Case Header.Justify
                        Case 1  'centre
                            If Printer.ScaleX(defMargin.width - ((setMargin.left - defMargin.width) + (setMargin.right - defMargin.right)), vbTwips, Printer.ScaleMode) - Printer.TextWidth(TitleText) > 0 Then
                                Printer.CurrentX = (Printer.ScaleX(defMargin.width - ((setMargin.left - defMargin.left) + (setMargin.right - defMargin.right)), vbTwips, Printer.ScaleMode) - Printer.TextWidth(TitleText)) \ 2
                            End If
                        Case 2 'right
                            Printer.CurrentX = Printer.ScaleX(defMargin.width - (setMargin.right - defMargin.right), vbTwips, Printer.ScaleMode) - (Printer.TextWidth(TitleText) + Printer.TextWidth("XX"))
                    End Select
                    Printer.Print TitleText
                Next RowLoop
                AppFont SavePFont, Printer.Font
            End If
        End If
        Printer.FontTransparent = True 'this broke back in VB version ? and remains for sentimental reasons
        'SetBkMode Correctly sets the background mix mode to transparent
        iBKMode = SetBkMode(Printer.hdc, TRANSPARENT)
        If PageH > 1 Then
            PYPos = Header.NextPage + Printer.ScaleY((setMargin.top - defMargin.top), vbTwips, Printer.ScaleMode)
        Else
            PYPos = Header.Page1Height + Printer.ScaleY((setMargin.top - defMargin.top), vbTwips, Printer.ScaleMode)
        End If
        
        If (PageH > 1 And Header.IncludeColHeadings) Or AddHeadingsToP1 Then
            'Code moved to Sub (below) at version 2.02
            RepeatColHeadings PageW, PYPos, 0
        End If
        'print any merged columns in the general area of the grid
        If PrintMerge.MergeCols Then    'go see if there are any in the fixed rows
            FindMergedCol PagesH(PageH).GridStart, PagesH(PageH).GridEnd, PagesW(PageW).GridStart, PagesW(PageW).GridEnd
            PrintMergedCols PYPos, 0
        End If
        'Now Print any Page Numbers required
        Select Case PageNumbs.numOption
            Case 0
                PrintNumber = False
            Case 1
                If PageH > 1 Or PageW > 1 Then
                    PrintNumber = True
                Else
                    PrintNumber = False
                End If
            Case Else
                PrintNumber = True
        End Select
        If PrintNumber Then
            PageString = "Page: " & CStr(PageNum)
            If PageNumbs.incPCount Then
                PageString = PageString & " of "
                If PrintMultiGrids Then
                    PageString = PageString & CStr(mRequestSectionsHigh)
                Else
                    PageString = PageString & CStr(userPagesHigh * userPagesWide)
                End If
            End If
            Select Case PageNumbs.numPos
                Case 0, 1, 2   'top
                    If PageNumbs.overMargin Then
                        Printer.CurrentY = 0
                    Else
                        Printer.CurrentY = Printer.ScaleY((setMargin.top - defMargin.top), vbTwips, Printer.ScaleMode)
                    End If
                Case Else   'bottom
                    Printer.CurrentY = Printer.ScaleHeight - Printer.TextHeight("P")
                    If Not PageNumbs.underMargin Then
                        Printer.CurrentY = Printer.CurrentY - Printer.ScaleY(setMargin.bottom - defMargin.bottom, vbTwips, Printer.ScaleMode)
                    End If
            End Select
            Select Case PageNumbs.numPos
                Case 0, 5   'right justify
                    Printer.CurrentX = Printer.ScaleWidth - Printer.ScaleX(setMargin.right - defMargin.right, vbTwips, Printer.ScaleMode) - Printer.TextWidth(PageString & "XX")
                Case 2, 3   'left justify
                    Printer.CurrentX = Printer.ScaleX((setMargin.left - defMargin.left), vbTwips, Printer.ScaleMode)
                Case 1, 4      'centre
                    Printer.CurrentX = (Printer.ScaleX(defMargin.width - ((setMargin.left - defMargin.left) + (setMargin.right - defMargin.right)), vbTwips, Printer.ScaleMode) - Printer.TextWidth(PageString)) / 2
            End Select
            Printer.Print PageString;
        End If
        Printer.CurrentX = Printer.ScaleX((setMargin.left - defMargin.left), vbTwips, Printer.ScaleMode)
        Printer.CurrentY = PYPos
        CurrentPageTop = PYPos
    End If
    StartNewPage = RetVal
    Exit Function
StartNewPageErr:
    Select Case Err
        Case 484
            Resume Next
        Case vbObjectError + MyE2999
            'Abort selected in a function called from this sub or function
            Exit Function
        Case Else
            Screen.MousePointer = vbDefault
            errorResponse = MsgBox("Unexpected Error Occured." & vbLf & "Please note details." & vbLf & Error(Err) & vbLf & "in Function:StartNewPage", vbExclamation + vbAbortRetryIgnore, "Unexpected Error")
            Select Case errorResponse
                Case vbAbort
                Case vbRetry
                    Resume
                Case vbIgnore
                    Resume Next
            End Select
    End Select
    Screen.MousePointer = vbDefault
    ClassErrorCode = MyE2999
    ClassErrorMessage = Error$
    Err.Raise Number:=ClassErrorCode + vbObjectError, Source:=ClassErrorSource, Description:=ClassErrorMessage
    Exit Function
End Function

Private Sub SetDefValues(DefSettings As Long, PrintTitle As String, NumCopies As Integer)
    'This routine takes the parameters passed to the Version 1 API to set up the
    'relevant internal property values before the PrintGridAPI routine
    'propper is called to execute the print
    
    mReportTitle = PrintTitle
    If DefSettings And ALLOW_DIALOGUE Then
        mAllowDialogue = True
    Else
        mAllowDialogue = False
    End If
    If NumCopies > 0 Then
        mPrintCopies = NumCopies
    Else
        mPrintCopies = 1
    End If
    
    'mtitlepages settings
    mTitlePages = 0
    If DefSettings And ALLOW_TITLE Then
        mTitlePages = mTitlePages + USER_MAY_SET
    End If
    If DefSettings And TITLE_FIRSTPAGE Then
        mTitlePages = mTitlePages + FIRST_PAGE_ONLY
    End If
    If DefSettings And TITLE_ALLPAGES Then
        mTitlePages = mTitlePages + TITLE_ALL_PAGES
    End If
    If DefSettings And REPEAT_HEADINGS Then
        mTitlePages = mTitlePages + REPEAT_COL_HEADINGS
    End If
    mTitlePages = mTitlePages + FONT_BOLD
    mTitlePages = mTitlePages + FONT_UNDERLINE
    'emulate old default title font
    TitleFont.Name = pFlexgrid.Font.Name
    TitleFont.Size = pFlexgrid.Font.Size
    TitleFont.Bold = True
    TitleFont.Underline = True
    
    'now set the mEffects values
    mEffects = COLUMN_HEAD_SHADING  'old default for monochrome printing
    If (DefSettings And ALLOW_COLOURS) Or (DefSettings And ALLOW_SETLINES) Then
        mEffects = mEffects + USER_CAN_CHANGE
    End If
    If DefSettings And SHOW_LINES Then
        mEffects = mEffects + GRID_LINES
    End If
    If DefSettings And GRID_FONT Then
        mEffects = mEffects + CELL_FONT_EFFECTS
    End If
    
    'the old interface had no page numbering settings so we will go with these
    mPageNumbering = USER_CAN_SET + AFTER_FIRST + TOP_RIGHT
End Sub
Private Sub ReSetGridPosition()
    Dim GridPos() As Variant
    
    On Error Resume Next
    GridPos() = DefSavePosition
    pFlexgrid.Col = GridPos(1)
    pFlexgrid.Row = GridPos(0)
    pFlexgrid.RowSel = GridPos(2)
    pFlexgrid.ColSel = GridPos(3)

End Sub
Private Function GetRowHeights() As Double
'© Copyright Adit Limited 1998, 1999, 2000, 2001, 2002
    Dim HiddenRow As Double, GridHeight As Double
    Dim RowLoop As Long
    
    On Error GoTo RowHeightErr
    HiddenRow = Screen.TwipsPerPixelY * 2
    ReDim RowHeights(pFlexgrid.Rows - 1)
    RowsToPrint = 0
    For RowLoop = 0 To (pFlexgrid.Rows - 1)
        If pFlexgrid.RowHeight(RowLoop) > HiddenRow Then
            RowHeights(RowLoop) = pFlexgrid.RowHeight(RowLoop)
            RowsToPrint = RowsToPrint + 1
        Else
            RowHeights(RowLoop) = 0
        End If
        GridHeight = GridHeight + RowHeights(RowLoop)
        If RowHeights(RowLoop) > HighRow Then
            HighRow = RowHeights(RowLoop)
        End If
    Next RowLoop
    GetRowHeights = GridHeight
    Exit Function
RowHeightErr:
    Select Case Err
        Case vbObjectError + MyE2999
            'Abort selected in a function called from this sub or function
            Exit Function
        Case Else
            Screen.MousePointer = vbDefault
            errorResponse = MsgBox("Unexpected Error Occured." & vbLf & "Please note details." & vbLf & Error(Err) & vbLf & "in Sub:RowHeightSet", vbExclamation + vbAbortRetryIgnore, "Unexpected Error")
            Select Case errorResponse
                Case vbAbort
                Case vbRetry
                    Resume
                Case vbIgnore
                    Resume Next
            End Select
    End Select
    Screen.MousePointer = vbDefault
    ClassErrorCode = MyE2999
    ClassErrorMessage = Error$
    Err.Raise Number:=ClassErrorCode + vbObjectError, Source:=ClassErrorSource, Description:=ClassErrorMessage
    Exit Function

End Function
Private Function GetSelectedHeight() As Double
    Dim selGrid() As Variant
    Dim RowLoop As Long
    Dim RetVal As Double
    
    selGrid() = DefSavePosition
    HighRow = 0
    RowsToPrint = 0
    For RowLoop = selGrid(0) To selGrid(2)
        If RowHeights(RowLoop) > 0 Then
            RowsToPrint = RowsToPrint + 1
        End If
        RetVal = RetVal + RowHeights(RowLoop)
        If RowHeights(RowLoop) > HighRow Then
            HighRow = RowHeights(RowLoop)
        End If
    Next RowLoop
    GetSelectedHeight = RetVal
End Function
Private Function GetColWidths() As Double
'© Copyright Adit Limited 1998, 1999, 2000, 2001, 2002
    Dim HiddenColumn As Double, GridWidth As Double
    Dim ColLoop As Long
    
    On Error Resume Next
    HiddenColumn = Screen.TwipsPerPixelX * 2
    ReDim ColWidths(pFlexgrid.Cols - 1)
    ColsToPrint = 0
    WideCol = 0
    For ColLoop = 0 To (pFlexgrid.Cols - 1)
        If pFlexgrid.ColWidth(ColLoop) > HiddenColumn Then
            ColWidths(ColLoop) = pFlexgrid.ColWidth(ColLoop) + TwipsMM
            ColsToPrint = ColsToPrint + 1
        Else
            ColWidths(ColLoop) = 0
        End If
        GridWidth = GridWidth + ColWidths(ColLoop)
        If ColWidths(ColLoop) > WideCol Then
            WideCol = ColWidths(ColLoop)
        End If
    Next ColLoop
    GetColWidths = GridWidth

End Function
Private Function GetSelectedWidth() As Double
    Dim selGrid() As Variant
    Dim ColLoop As Long
    Dim RetVal As Double
    
    selGrid() = DefSavePosition
    WideCol = 0
    ColsToPrint = 0
    For ColLoop = selGrid(1) To selGrid(3)
        RetVal = RetVal + ColWidths(ColLoop)
        If ColWidths(ColLoop) > 0 Then
            ColsToPrint = ColsToPrint + 1
        End If
        If ColWidths(ColLoop) > WideCol Then
            WideCol = ColWidths(ColLoop)
        End If
    Next ColLoop
    GetSelectedWidth = RetVal
End Function
Private Sub GetWideCol()
    Dim ColLoop As Long
    For ColLoop = ColWidths(0) To UBound(ColWidths())
        If ColWidths(ColLoop) > WideCol Then
            WideCol = ColWidths(WideCol)
        End If
    Next ColLoop
End Sub
Private Sub GetHighRow()
    Dim RowLoop As Long
    For RowLoop = RowHeights(0) To UBound(RowHeights())
        If RowHeights(RowLoop) > HighRow Then
            HighRow = RowHeights(RowLoop)
        End If
    Next RowLoop
End Sub
Private Sub SaveGridPosition()
    ReDim GridPosition(3) As Variant
    
    On Error Resume Next
    UserSelection = False
    If pFlexgrid.RowSel >= pFlexgrid.Row Then
        GridPosition(0) = pFlexgrid.Row
        GridPosition(2) = pFlexgrid.RowSel
    Else
        GridPosition(2) = pFlexgrid.Row
        GridPosition(0) = pFlexgrid.RowSel
    End If
    If pFlexgrid.ColSel >= pFlexgrid.Col Then
        GridPosition(1) = pFlexgrid.Col
        GridPosition(3) = pFlexgrid.ColSel
    Else
        GridPosition(3) = pFlexgrid.Col
        GridPosition(1) = pFlexgrid.ColSel
    End If
    
    DefSavePosition = GridPosition()
    If pFlexgrid.Row <> pFlexgrid.RowSel Or pFlexgrid.Col <> pFlexgrid.ColSel Then
        UserSelection = True    'there is a selected area > a single cell
    End If
End Sub
Private Function ParsePageList(PageList As Variant) As Variant
'© Copyright Adit Limited 1998, 1999, 2000, 2001, 2002
    Dim CharLoop As Long, PageCount As Long, RangeLoop As Long
    Dim PageNums() As Variant, Thispage As Variant, savepage As Variant
    Dim RangeNow As Boolean
    
    On Error GoTo ParsePageListErr
    PageCount = -1
    UserSelectedPagesCount = 0
    For CharLoop = 1 To Len(PageList)
        Select Case Mid(PageList, CharLoop, 1)
            Case "0" To "9"
                Thispage = Thispage & Mid(PageList, CharLoop, 1)
            Case ",", " "
                If Thispage > "" Then
                    If RangeNow And PageCount > -1 Then
                        GoSub AddARange
                    End If
                    GoSub AddAPage
                
                End If
            Case "-"
                RangeNow = True
                If Thispage > "" Then
                    GoSub AddAPage
                End If
        End Select
    Next CharLoop
    If Thispage > "" Then
        If RangeNow And PageCount > -1 Then
            GoSub AddARange
        End If
        GoSub AddAPage
    End If
    If PageCount > -1 Then
        ParsePageList = PageNums()
    End If
    UserSelectedPagesCount = PageCount + 1
    Exit Function
AddAPage:
        PageCount = PageCount + 1
        ReDim Preserve PageNums(PageCount)
        PageNums(PageCount) = Thispage
        Thispage = ""
    Return
AddARange:
        savepage = Thispage
        For RangeLoop = (PageNums(PageCount) + 1) To (savepage - 1)
            Thispage = RangeLoop
            GoSub AddAPage
        Next RangeLoop
        Thispage = savepage
        RangeNow = False
    Return
ParsePageListErr:
    Select Case Err
        Case vbObjectError + MyE2999
            'Abort selected in a function called from this sub or function
            Exit Function
        Case Else
            Screen.MousePointer = vbDefault
            errorResponse = MsgBox("Unexpected Error Occured." & vbLf & "Please note details." & vbLf & Error(Err) & vbLf & "in Function:ParsePageList", vbExclamation + vbAbortRetryIgnore, "Unexpected Error")
            Select Case errorResponse
                Case vbAbort
                Case vbRetry
                    Resume
                Case vbIgnore
                    Resume Next
            End Select
    End Select
    Screen.MousePointer = vbDefault
    ClassErrorCode = MyE2999
    ClassErrorMessage = Error$
    Err.Raise Number:=ClassErrorCode + vbObjectError, Source:=ClassErrorSource, Description:=ClassErrorMessage
    Exit Function
End Function
Private Sub PrintRowOnPage(GridRow As Long, PageW As Long, SectionXOffset As Single)
'© Copyright Adit Limited 1998 to 2005
    Dim PX As Single, PY As Single, PPX As Single, PPY As Single
    Dim ColLoop As Long, MLoop As Long, FontChoice As Long, Cloop As Long
    Dim UsedWidth As Double
    Dim CellmergedVert As Boolean
    Dim CellAlign As CellAlignment
    Dim CellContents As CellC
    Dim printcolour As GColr
    Dim MergeCells As Boolean, PictureOnly As Boolean
    Dim PageCells() As Long, CellCount As Long, InsPoint As Long, CellLoop As Long

    On Error GoTo PrintRowErr:
    PY = Printer.CurrentY
    PX = Printer.ScaleX((setMargin.left - defMargin.left), vbTwips, Printer.ScaleMode)
    PX = PX + SectionXOffset
    
    'Code block to control repeating fixed columns
    'so we have to be able to manage a non-contiguous sequence of columns
    CellCount = (PagesW(PageW).GridEnd - PagesW(PageW).GridStart) + 1
    ReDim PageCells(CellCount)
    InsPoint = 0
    If (FixedColData.FixedWidth > 0 And PageW > 1) Or (FixedColData.FixedPage1 > 0 And PageW = 1) Then
        CellCount = CellCount + pFlexgrid.FixedCols
        ReDim PageCells(CellCount)
        For ColLoop = 0 To (pFlexgrid.FixedCols - 1)
            PageCells(ColLoop + InsPoint + 1) = ColLoop
        Next ColLoop
        InsPoint = pFlexgrid.FixedCols
    End If
    For ColLoop = PagesW(PageW).GridStart To PagesW(PageW).GridEnd
        InsPoint = InsPoint + 1
        PageCells(InsPoint) = ColLoop
    Next ColLoop
    
    For CellLoop = 1 To CellCount
        ColLoop = PageCells(CellLoop)
        CellmergedVert = False
        If PrintMerge.MergeCols Then
            If CCells(GridRow, ColLoop) > 0 Then
                CellmergedVert = True
            End If
        End If
        If PrintMerge.MergeRows Then
            If PrintMerge.MergeRule = 0 Then
                'restrict the search for merged cells to the current page
                FindMergedRow GridRow, PagesW(PageW).GridStart, PagesW(PageW).GridEnd
            Else
                'look for merged cells across the whole row
                FindMergedRow GridRow, 0, (pFlexgrid.Cols - 1)
            End If
        End If
        If ColWidths(ColLoop) > 0 And Not CellmergedVert Then
            CellAlign = GetCellAlign(GridRow, ColLoop)
            If CellAlign = CellAligndef Then
                'Use column default
                If GridRow <= (pFlexgrid.FixedRows - 1) Then
                    CellAlign = pFlexgrid.FixedAlignment(ColLoop) + 1
                Else
                    CellAlign = ColAlign(ColLoop)
                End If
            End If
            CellContents.GridText = pFlexgrid.TextMatrix(GridRow, ColLoop)
            CellContents.CellWidth = ColWidths(ColLoop)
            CellContents.CellHeight = RowHeights(GridRow)
            If PrintMerge.MergeRows Then
                If MCells(ColLoop) > 0 Then 'this is a merged cell
                    If MCells(ColLoop) Mod 1000 = 1 Or ColLoop = PagesW(PageW).GridStart Then
                        'we have to do something cos it is the first cell or the first one on this page
                        For MLoop = (ColLoop + 1) To UBound(MCells()) '(pFlexgrid.Cols - 1)
                            If left$(Format$(MCells(MLoop), "000000"), 3) = left$(Format$(MCells(ColLoop), "000000"), 3) Then
                                If MLoop > PagesW(PageW).GridEnd Then
                                    CellAlign = CellAlignleft 'overide any other justification
                                    Exit For
                                End If
                                CellContents.CellWidth = CellContents.CellWidth + ColWidths(MLoop)
                            Else
                                Exit For
                            End If
                        Next MLoop
                        If ColLoop > 0 And ColLoop = PagesW(PageW).GridStart And PrintMerge.MergeRule = 1 Then 'some may already have been printed
                            UsedWidth = 0
                            For MLoop = (ColLoop - 1) To LBound(MCells()) Step -1 '0 Step -1
                                If left$(Format$(MCells(MLoop), "000000"), 3) = left$(Format$(MCells(ColLoop), "000000"), 3) Then
                                    UsedWidth = UsedWidth + ColWidths(MLoop)
                                    CellAlign = CellAlignleft
                                Else
                                    Exit For
                                End If
                            Next MLoop
                            If UsedWidth > 0 Then 'we know that some or all of the text has been printed
                                CellContents.AlreadyUsed = UsedWidth
                                CellContents = TruncateAndWrapC(CellContents)
                                'and truncate the string by the amount printed
                            End If
                        End If
                        
                    Else
                        CellContents.GridText = "" 'should have been done
                        CellContents.CellWidth = 0   'The grid will have been drawn as well
                    End If
                Else
                    CellContents.CellWidth = ColWidths(ColLoop)
                End If
            Else
                CellContents.CellWidth = ColWidths(ColLoop)
            End If
            printcolour = GetColours(GridRow, ColLoop)
            If ShowLines And CellContents.CellWidth > 0 Then
                Printer.FillColor = printcolour.Back
                Printer.FillStyle = vbFSSolid
                Printer.ForeColor = printcolour.Grid
                If InvisibleGrid Then
                    Printer.DrawStyle = vbInvisible
                End If
                Printer.Line (PX, PY)-(PX + CellContents.CellWidth, PY + RowHeights(GridRow)), , B
                If LastLineColour > -1 Then 'we may need to redraw the top of the box to the previous rows grid colour
                    If LastLineColour <> printcolour.Grid Then
                        Printer.Line (PX, PY)-(PX + CellContents.CellWidth, PY), LastLineColour
                    End If
                End If
                Printer.CurrentY = PY
            End If
            CellContents = TruncateAndWrapC(CellContents)
            PPX = PX
            If CellContents.PrintLine > "" Or CellContents.WrapCount > 0 Then
                If mPrintCellImages Then
                    'we place cell images behind any printed text
                    'so we will copy any picture to a handy picture control and then move
                    'any content to the printed page
                    
                    pFlexgrid.Row = GridRow
                    pFlexgrid.Col = ColLoop
                    'test to see if there is a cell picture
                    If pFlexgrid.CellPicture <> 0 Then
                        'There is - so we can have a go at printing it
                        Picture1.Picture = pFlexgrid.CellPicture
                        CellImagePrint PX, PY, CellContents.CellWidth, CellContents.CellHeight
                    End If
                End If
                Printer.ForeColor = printcolour.Front
                If CellAlign And CellAlignleft Then
                    PPX = PX + ColSpace
                End If
                If CellAlign And CellAlignCentre Then
                    PPX = PX + ((CellContents.CellWidth - CellContents.PrintWidth) / 2)
                End If
                If CellAlign And CellAlignRight Then
                    PPX = PX + (CellContents.CellWidth - CellContents.PrintWidth) - ColSpace
                End If
                PPY = PY
                If CellAlign And cellalignbottom Then
                    If (CellContents.CellHeight - CellContents.PrintHeight) > 0 Then
                        PPY = PY + (CellContents.CellHeight - CellContents.PrintHeight)
                    End If
                End If
                If CellAlign And CellAlignMiddle Then
                    If (CellContents.CellHeight - CellContents.PrintHeight) > 0 Then
                        PPY = PY + ((CellContents.CellHeight - CellContents.PrintHeight) / 2)
                    End If
                End If
                If GridFont Then
                    FontChoice = GetGridFont(GridRow, ColLoop)
                    If FontChoice And 1 Then
                        Printer.FontBold = True
                    End If
                    If FontChoice And 2 Then
                        Printer.FontItalic = True
                    End If
                    If FontChoice And 4 Then
                        Printer.FontUnderline = True
                    End If
                End If
                
                Printer.CurrentX = PPX
                Printer.CurrentY = PPY
                
                If CellContents.WrapCount > 0 Then
                    For Cloop = 0 To CellContents.WrapCount - 1
                        Printer.CurrentX = PPX
                        Printer.Print CellContents.WrapLines(Cloop);
                        If Cloop < CellContents.WrapCount - 1 Then
                            Printer.Print vbLf;
                        End If
                    Next Cloop
                Else
                    Printer.Print CellContents.PrintLine;
                End If
                If GridFont Then    'undo the things again
                    Printer.FontBold = False
                    Printer.FontItalic = False
                    Printer.FontUnderline = False
                End If
            Else
                'slightly clumsy treatment of non merged cells with a picture but no text
                If mPrintCellImages Then
                    pFlexgrid.Row = GridRow
                    pFlexgrid.Col = ColLoop
                    PictureOnly = False
                    If pFlexgrid.CellPicture <> 0 Then
                        If PrintMerge.MergeRows Then
                            If MCells(ColLoop) = 0 Then
                                PictureOnly = True
                            End If
                        Else
                            PictureOnly = True
                        End If
                        If PictureOnly Then
                            Picture1.Picture = pFlexgrid.CellPicture
                            CellImagePrint PX, PY, CellContents.CellWidth, CellContents.CellHeight
                        End If
                    End If
                End If
            End If
            Printer.CurrentY = PY
        End If
        PX = PX + ColWidths(ColLoop)
    Next CellLoop
    LastLineColour = printcolour.Grid 'not perfect but will almost always work
                        'as this stops the colour of the first non fixed grid line
                        'overwriting the colour of the last fixed grid line - well as far as the printed output is concerned
    Exit Sub
PrintRowErr:
    Select Case Err
        Case 484
            Resume Next
        Case vbObjectError + MyE2999
            'Abort selected in a function called from this sub or function
            Exit Sub
        Case Else
            Screen.MousePointer = vbDefault
            errorResponse = MsgBox("Unexpected Error Occured." & vbLf & "Please note details." & vbLf & Error(Err) & vbLf & "in Sub:PrintRowOnPage", vbExclamation + vbAbortRetryIgnore, "Unexpected Error")
            Select Case errorResponse
                Case vbAbort
                Case vbRetry
                    Resume
                Case vbIgnore
                    Resume Next
            End Select
    End Select
    Screen.MousePointer = vbDefault
    ClassErrorCode = MyE2999
    ClassErrorMessage = Error$
    Err.Raise Number:=ClassErrorCode + vbObjectError, Source:=ClassErrorSource, Description:=ClassErrorMessage
    Exit Sub
End Sub
Private Sub CellImagePrint(CXPos As Single, CYPos As Single, CellWidth As Single, CellHeight As Single)
    'Vehicle to print a cell graphic within a cell (including merged cells)
    'Picture1.Picture already holds the image we are copying
    Dim ImageAlign As Long, PosAdjust As Long
    Dim Xpixel As Single, YPixel As Single, CellDimX As Single, CellDimY As Single
    Dim ImageX As Single, ImageY As Single
    
    'get the picture size in pixels
    ImageX = Picture1.ScaleWidth
    ImageY = Picture1.ScaleHeight
    'estimate the available size in pixels as well
    CellDimX = Int(Me.ScaleX((CellWidth - ColSpace), vbTwips, vbPixels))
    CellDimY = Int(Me.ScaleX(CellHeight, vbTwips, vbPixels)) - 2
    'We will only use part of the picture if it is larger than the available space
    If ImageX > CellDimX Then
        ImageX = CellDimX
    End If
    If ImageY > CellDimY Then
        ImageY = CellDimY
    End If
    'now switch the measurement to the usable image size in twips so we can position it if required
    ' and scale it correctly later
    CellDimX = Me.ScaleX(ImageX, vbPixels, vbTwips)
    CellDimY = Me.ScaleY(ImageY, vbPixels, vbTwips)
    'now sort out alignment within the cell
    'we are using ColSpace to adjust X dimensions and a single pixel to adjust Y dimensions
    'as this seems to improve the presentation - feel free to change this
    ImageAlign = pFlexgrid.CellPictureAlignment
    Select Case ImageAlign
        Case 3, 4, 5    'Central
            PosAdjust = (CellWidth - CellDimX) \ 2
            If PosAdjust > ColSpace Then
                Xpixel = CXPos + PosAdjust
            Else
                Xpixel = CXPos + ColSpace
            End If
        Case 6, 7, 8    'rightish
            PosAdjust = (CellWidth - CellDimX) - Printer.TwipsPerPixelX
            If PosAdjust > ColSpace Then
                Xpixel = CXPos + PosAdjust
            Else
                Xpixel = CXPos + ColSpace
            End If
        Case Else       'leftish
            Xpixel = CXPos + ColSpace
    End Select
    Select Case ImageAlign
        Case 1, 4, 7    'middlish
            PosAdjust = (CellHeight - CellDimY) \ 2
            If PosAdjust > Printer.TwipsPerPixelY Then
                YPixel = CYPos + PosAdjust
            Else
                YPixel = CYPos + Printer.TwipsPerPixelY
            End If
        Case 2, 5, 8    'bottom
            PosAdjust = (CellHeight - CellDimY) - Printer.TwipsPerPixelY
            If PosAdjust > Printer.TwipsPerPixelY Then
                YPixel = CYPos + PosAdjust
            Else
                YPixel = CYPos + Printer.TwipsPerPixelY
            End If
        Case Else       'top
            YPixel = CYPos + Printer.TwipsPerPixelY
    End Select
    'then convert the image size to the printer pixel resolution
    'allowing for any compression set
    CellDimX = Printer.ScaleX(CellDimX, Printer.ScaleMode, vbPixels)
    CellDimY = Printer.ScaleY(CellDimY, Printer.ScaleMode, vbPixels)
    'then define the target image location using the same scale
    Xpixel = Printer.ScaleX(Xpixel, Printer.ScaleMode, vbPixels)
    YPixel = Printer.ScaleY(YPixel, Printer.ScaleMode, vbPixels)
    
    'now use StretchBlt to place the picture on the page
    ImageAlign = StretchBlt(Printer.hdc, Xpixel, YPixel, CellDimX, CellDimY, Picture1.hdc, 0, 0, ImageX, ImageY, SRCCOPY)
    'all done
End Sub
Private Function GetCellAlign(Row As Long, Col As Long) As CellAlignment
'© Copyright Adit Limited 1998, 1999, 2000, 2001, 2002
    Dim RetAlign As CellAlignment
    
    On Error Resume Next
    pFlexgrid.Row = Row
    pFlexgrid.Col = Col
    Select Case pFlexgrid.CellAlignment
        Case 1
            RetAlign = CellAlignleft + CellAlignMiddle
        Case 2
            RetAlign = CellAlignleft + cellalignbottom
        Case 3
            RetAlign = CellAlignCentre + CellAlignTop
        Case 4
            RetAlign = CellAlignCentre + CellAlignMiddle
        Case 5
            RetAlign = CellAlignCentre + cellalignbottom
        Case 6
            RetAlign = CellAlignRight + CellAlignTop
        Case 7
            RetAlign = CellAlignRight + CellAlignMiddle
        Case 8
            RetAlign = CellAlignRight + cellalignbottom
        Case Else 'Default and normal result of this function
            RetAlign = CellAligndef
    End Select
    GetCellAlign = RetAlign
    
End Function
Private Function GetColours(Row As Long, Col As Long) As GColr
'© Copyright Adit Limited 1998, 1999, 2000, 2001, 2002
    Dim ColourLong As Long
    Dim MyReturn As GColr
    
    On Error GoTo GetColoursErr
    If Col < pFlexgrid.FixedCols Or Row < pFlexgrid.FixedRows Then
        If ShowCellColour Then
            MyReturn.Back = pFlexgrid.BackColorFixed
        Else
            If ShadeFixedRows Then
                MyReturn.Back = DefFixB
            Else
                MyReturn.Back = vbWhite
            End If
        End If
        If ShowTextColour Then
            MyReturn.Front = pFlexgrid.ForeColorFixed
        Else
            MyReturn.Front = MYBLACK
        End If
        If ShowGridColour Then
            MyReturn.Grid = pFlexgrid.GridColorFixed
        Else
            MyReturn.Grid = MYBLACK
        End If
    Else
        If ShowCellColour Then
            pFlexgrid.Row = Row
            pFlexgrid.Col = Col
            ColourLong = pFlexgrid.CellBackColor
            If ColourLong = 0 Then
                MyReturn.Back = pFlexgrid.BackColor
            Else
                MyReturn.Back = ColourLong
            End If
        Else
            MyReturn.Back = vbWhite
        End If
        If ShowTextColour Then
            MyReturn.Front = pFlexgrid.CellForeColor
        Else
            MyReturn.Front = MYBLACK
        End If
        If ShowGridColour Then
            MyReturn.Grid = pFlexgrid.GridColor
        Else
            MyReturn.Grid = MYBLACK
        End If
    End If
    GetColours = MyReturn
    Exit Function
GetColoursErr:
    'Error Recovery - no Database
    Select Case Err
        Case vbObjectError + MyE2999
            'Abort selected in a function called from this sub or function
            Exit Function
        Case Else
            Screen.MousePointer = vbDefault
            errorResponse = MsgBox("Unexpected Error Occured." & vbLf & "Please note details." & vbLf & Error(Err) & vbLf & "in Function:GetColours", vbExclamation + vbAbortRetryIgnore, "Unexpected Error")
            Select Case errorResponse
                Case vbAbort
                Case vbRetry
                    Resume
                Case vbIgnore
                    Resume Next
            End Select
    End Select
    Screen.MousePointer = vbDefault
    ClassErrorCode = MyE2999
    ClassErrorMessage = Error$
    Err.Raise Number:=ClassErrorCode + vbObjectError, Source:=ClassErrorSource, Description:=ClassErrorMessage
    Exit Function
End Function

Private Function GetGridFont(ByVal GridRow As Long, ByVal Gridcol As Long) As Long
'© Copyright Adit Limited 1998, 1999, 2000, 2001, 2002
    Dim RetVal As Long
    
    On Error GoTo GetGridFontErr
    pFlexgrid.Col = Gridcol
    pFlexgrid.Row = GridRow
    If pFlexgrid.CellFontBold Then
        RetVal = 1
    End If
    If pFlexgrid.CellFontItalic Then
        RetVal = RetVal + 2
    End If
    If pFlexgrid.CellFontUnderline Then
        RetVal = RetVal + 4
    End If
    
    GetGridFont = RetVal
    Exit Function
GetGridFontErr:
    Select Case Err
        Case vbObjectError + MyE2999
            'Abort selected in a function called from this sub or function
            Exit Function
        Case Else
            Screen.MousePointer = vbDefault
            errorResponse = MsgBox("Unexpected Error Occured." & vbLf & "Please note details." & vbLf & Error(Err) & vbLf & "in Function:GetGridFont", vbExclamation + vbAbortRetryIgnore, "Unexpected Error")
            Select Case errorResponse
                Case vbAbort
                Case vbRetry
                    Resume
                Case vbIgnore
                    Resume Next
            End Select
    End Select
    Screen.MousePointer = vbDefault
    ClassErrorCode = MyE2999
    ClassErrorMessage = Error$
    Err.Raise Number:=ClassErrorCode + vbObjectError, Source:=ClassErrorSource, Description:=ClassErrorMessage
    Exit Function
End Function

Private Sub Label1_Click(Index As Integer)
    Select Case Index
        Case 8
            Image1_Click
    End Select
End Sub


Public Property Let GridToPrint(ByVal vGridName As MSHFlexGrid)
    Set pFlexgrid = vGridName
    AppFont pFlexgrid.Font, TitleFont
End Property

Public Property Let ReportTitle(ByVal vNewValue As Variant)
    mReportTitle = vNewValue
End Property
Private Function GridFitsPage() As Boolean
    'This function checks that it is possible to print the current grid
    'on the selected page within the current margins
    Dim RetVal As Boolean
    Dim PrinterWidth As Double, Printerheight As Double
    Dim QValue As Long
    
    RetVal = True
    PrinterWidth = Printer.ScaleX(defMargin.width - (setMargin.left + setMargin.right), vbTwips, Printer.ScaleMode)
    Printerheight = Printer.ScaleY(defMargin.height - (setMargin.top + setMargin.bottom), vbTwips, Printer.ScaleMode) - PageNumbs.footHeight
    If HighRow > Printerheight Then
        If Me.Visible Then
            Msg = "This report includes one, or more, row higher"
            Msg = Msg & vbLf & "than the available page height on the current"
            Msg = Msg & vbLf & "default printer. Please adjust and try again."
            title = "Invalid Grid Row Height"
            Screen.MousePointer = vbDefault
            MsgBox Msg, vbCritical + vbOKOnly, title
        End If
        RetVal = False
    End If
    If WideCol > PrinterWidth Then
        If Me.Visible Then  'the user should see this message
            Msg = "This report includes one, or more, column wider"
            Msg = Msg & vbLf & "than the available page width on the current"
            Msg = Msg & vbLf & "selected printer. Please change the Paper Size,"
            Msg = Msg & vbLf & "adjust the page margins, or resize the grid column."
            title = "Invalid Grid Column Width"
            Screen.MousePointer = vbDefault
            MsgBox Msg, vbCritical + vbOKOnly, title
        End If
        RetVal = False
    End If
    If Header.Page1Height > Printerheight Then
        If Me.Visible Then  'the user should see this message although it map be an API generated problem
            Msg = "This report includes a title and sub-title longer"
            Msg = Msg & vbLf & "than the available page height on the current"
            Msg = Msg & vbLf & "selected printer. Please change the Paper Size,"
            Msg = Msg & vbLf & "adjust the page margins, or resize the title font."
            title = "Invalid Title Length"
            Screen.MousePointer = vbDefault
            MsgBox Msg, vbCritical + vbOKOnly, title
        End If
        RetVal = False
    End If
    If Header.NextPage > Printerheight And Header.IncludeColHeadings Then
        If Me.Visible Then
            Msg = "The Column headings exceed the available height of the current"
            Msg = Msg & vbLf & "default printer page. Would you like to cancel the"
            Msg = Msg & vbLf & "repeat setting?"
            title = "Unable to Repeat Page Headings"
            Screen.MousePointer = vbDefault
            QValue = MsgBox(Msg, vbInformation + vbYesNo, title)
            If QValue = vbYes Then
                Header.IncludeColHeadings = False
                Header.NextPage = Header.Page1Height
            Else
                RetVal = False
            End If
        Else
            RetVal = False
        End If
    End If
    If FixedColData.FixedWidth > PrinterWidth And FixedColData.RepeatFixed Then
        If Me.Visible Then
            Msg = "The row captions exceed the available width of the current"
            Msg = Msg & vbLf & "default printer page. Would you like to cancel the"
            Msg = Msg & vbLf & "repeat setting?"
            title = "Unable to Repeat Row Captions"
            Screen.MousePointer = vbDefault
            QValue = MsgBox(Msg, vbInformation + vbYesNo, title)
            If QValue = vbYes Then
                FixedColData.RepeatFixed = False
                FixedColData.HighCol = -1
                FixedColData.FixedPage1 = 0
                FixedColData.FixedWidth = 0
            Else
                RetVal = False
            End If
        Else
            RetVal = False
        End If
    End If
    GridFitsPage = RetVal
End Function


Public Property Let AllowDialogue(ByVal bShowForm As Boolean)
    mAllowDialogue = bShowForm
End Property


Public Property Let PrintCopies(ByVal lReportCopies As Long)
    If lReportCopies > 1 Then
        mPrintCopies = lReportCopies
    Else
        mPrintCopies = 1
    End If
End Property

Private Sub Option1_Click(Index As Integer)
    Select Case Index
        Case 0  'Print Full Size
            Text1(0).Text = defPagesWide
            Text1(1).Text = defPagesHigh
            Text1(0).Enabled = False
            Text1(1).Enabled = False
            UpDown1(0).Enabled = False
            UpDown1(1).Enabled = False
            userPagesWide = defPagesWide
            userPagesHigh = defPagesHigh
            PrintMultiGrids = False
        Case 1   'Print to Fit
            Text1(0).Enabled = True
            Text1(1).Enabled = True
            UpDown1(0).Enabled = True
            UpDown1(1).Enabled = True
            userPagesWide = CLng(Text1(0).Text)
            userPagesHigh = CLng(Text1(1).Text)
            PrintMultiGrids = False
        Case 2  'Print multiple sections on each page
            Text1(0).Enabled = False
            Text1(1).Enabled = False
            UpDown1(0).Enabled = False
            UpDown1(1).Enabled = False
            PrintMultiGrids = True
    End Select
    StatusBarPages
End Sub
Private Sub Option2_Click(Index As Integer)

    If InitialisingForm Then
        Exit Sub
    End If
    Text1(8).Enabled = False
    Select Case Index
        Case 0  'Print All
            GridWidth = GetColWidths    'just in case we came from selected
            GridHeight = GetRowHeights
            mPrintSelected = False
            CountPages True
        Case 1  'print selected area of grid
            mPrintSelected = True
            CountPages True
        Case 2  'Print selected pages
            If Option1(2).Value Then
                If mRequestSectionsHigh > 1 Then
                    Text1(8).Text = "1-" & CStr(mRequestSectionsHigh)
                    Text1(8).Enabled = True
                    Text1(8).SetFocus
                Else
                    Text1(8).Text = "1"
                End If
            Else
                If (userPagesWide * userPagesHigh) > 1 Then
                    Text1(8).Text = "1-" & CStr(userPagesWide * userPagesHigh)
                    Text1(8).Enabled = True
                    Text1(8).SetFocus
                Else
                    Text1(8).Text = "1"
                End If
            End If
    End Select
End Sub

Private Sub Option3_Click(Index As Integer)
    Header.PrintPage = Index
    CountPages True
End Sub

Private Sub Option4_Click(Index As Integer)
    PageNumbs.numOption = Index
End Sub

Private Sub Option5_Click(Index As Integer)
    Header.Justify = Index
End Sub

Private Sub Option6_Click(Index As Integer)
    PrintMerge.MergeRule = Index
End Sub

Private Sub TabStrip1_Click()
    Frame1(CurrentFrame).Visible = False
    Select Case TabStrip1.SelectedItem.Index
        Case 1  'page selection
            CurrentFrame = 0
        Case 2  'Title options
            CurrentFrame = 11
        Case 3  'Printer options
            CurrentFrame = 5
        Case 4  'Effects
            CurrentFrame = 8
        Case 5  'margins and page numbering
            CurrentFrame = 6
    End Select
    Frame1(CurrentFrame).left = TabStrip1.left
    Frame1(CurrentFrame).top = TabStrip1.height / 2
    Frame1(CurrentFrame).width = TabStrip1.width
    Frame1(CurrentFrame).Visible = True
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    Text1(Index).SelStart = False
    Text1(Index).SelLength = 32767
    Select Case Index
        Case 3, 4, 5, 6, 7, 9
            CheckNeeded = True
    End Select
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case 0, 1, 2
            KeyAscii = NumericOnly(KeyAscii)
        Case 3
        Case 4, 5, 6, 7
            If UserPSize.MSystem = 1 Then
                KeyAscii = NumericOnly(KeyAscii, ".")
            Else
                KeyAscii = NumericOnly(KeyAscii)
            End If
        Case 8
            KeyAscii = NumericOnly(KeyAscii, ",-")
        Case 9
            KeyAscii = NumericOnly(KeyAscii, ".")
    End Select
End Sub


Private Sub Text1_LostFocus(Index As Integer)
    
    Select Case Index
        Case 3
            mReportTitle = Text1(Index).Text
            CountPages True
        Case 0, 1, 2
            If Text1(Index).Text = "" Then
                Text1(Index).Text = "1"
            ElseIf CInt(Text1(Index).Text) < 1 Then
                Text1(Index).Text = 1
            End If
            Select Case Index
                Case 0
                    If CInt(Text1(Index).Text) > defPagesWide Then
                        Text1(Index).Text = defPagesWide
                    End If
                    userPagesWide = CLng(Text1(Index).Text)
                    StatusBarPages
                Case 1
                    If CInt(Text1(Index).Text) > defPagesHigh Then
                        Text1(Index).Text = defPagesHigh
                    End If
                    userPagesHigh = CLng(Text1(Index).Text)
                    StatusBarPages
                Case 2
                    mPrintCopies = CInt(Text1(Index).Text)
                    StatusBar1.Panels(3).Text = "Copies: " & Text1(Index).Text
            End Select
        Case 4, 5, 6, 7 'margins
            CheckMEntry Index
        Case 8
            SetSelectedPages
            StatusBarPages
        Case 9
            Header.FontScale = CSng(Text1(Index).Text)
            CountPages True
    End Select
End Sub
Private Sub SetSelectedPages()
    Dim workVar As Variant
    
    Selectedpages = Empty
    If Text1(8).Text > "" Then
        workVar = ParsePageList(Text1(8).Text)
        If Not IsEmpty(workVar) Then
            Selectedpages = workVar
        Else
            MsgBox "Selected page range invalid, please try again.", vbOKOnly, "Selected Pages"
        End If
    End If
End Sub
Private Sub CheckMEntry(Index As Integer)
    Dim OldMargin As Double, newMargin As Double
    Select Case Index
        Case 4
            OldMargin = setMargin.top
            If IsNumeric(Text1(Index).Text) Then
                setMargin.top = CDbl(Text1(Index).Text)
                If UserPSize.MSystem = 1 Then
                    setMargin.top = setMargin.top * 1440
                Else
                    setMargin.top = setMargin.top * TwipsMM
                End If
                If setMargin.top < defMargin.top Then
                    setMargin.top = defMargin.top
                    MarginDisplay
                End If
            Else
                setMargin.top = defMargin.top
                MarginDisplay
            End If
            newMargin = setMargin.top
        Case 5
            OldMargin = setMargin.bottom
            If IsNumeric(Text1(Index).Text) Then
                setMargin.bottom = CDbl(Text1(Index).Text)
                If UserPSize.MSystem = 1 Then
                    setMargin.bottom = setMargin.bottom * 1440
                Else
                    setMargin.bottom = setMargin.bottom * TwipsMM
                End If
                If setMargin.bottom < defMargin.bottom Then
                    setMargin.bottom = defMargin.bottom
                    MarginDisplay
                End If
            Else
                setMargin.bottom = defMargin.bottom
                MarginDisplay
            End If
            newMargin = setMargin.bottom
        Case 6  'left margin
            OldMargin = setMargin.left
            If IsNumeric(Text1(Index).Text) Then
                setMargin.left = CDbl(Text1(Index).Text)
                If UserPSize.MSystem = 1 Then
                    setMargin.left = setMargin.left * 1440
                Else
                    setMargin.left = setMargin.left * TwipsMM
                End If
                If setMargin.left < defMargin.left Then
                    setMargin.left = defMargin.left
                    MarginDisplay
                End If
            Else
                setMargin.left = defMargin.left
                MarginDisplay
            End If
            newMargin = setMargin.left
        Case 7  'right margin
            OldMargin = setMargin.right
            If IsNumeric(Text1(Index).Text) Then
                setMargin.right = CDbl(Text1(Index).Text)
                If UserPSize.MSystem = 1 Then
                    setMargin.right = setMargin.right * 1440
                Else
                    setMargin.right = setMargin.right * TwipsMM
                End If
                If setMargin.right < defMargin.right Then
                    setMargin.right = defMargin.right
                    MarginDisplay
                End If
            Else
                setMargin.right = defMargin.right
                MarginDisplay
            End If
            newMargin = setMargin.right
    End Select
    If Abs(OldMargin - newMargin) > 25 Then
        CountPages True
    Else
        CheckNeeded = False
    End If
End Sub


Private Sub UpDown1_DownClick(Index As Integer)
    Select Case Index
        Case 0, 1, 2
            If Text1(Index).Text = "" Then
                Text1(Index).Text = "1"
            End If
            If CInt(Text1(Index).Text) > 1 Then
                Text1(Index).Text = CInt(Text1(Index).Text) - 1
            End If
            Select Case Index
                Case 0
                    userPagesWide = CLng(Text1(0).Text)
                    StatusBarPages
                Case 1
                    userPagesHigh = CLng(Text1(1).Text)
                    StatusBarPages
                Case 2
                    StatusBar1.Panels(3).Text = "Copies: " & Text1(Index).Text
            End Select
        Case 3  'top margin
            If setMargin.top > (defMargin.top + MarginAdjust) Then
                setMargin.top = setMargin.top - MarginAdjust
            Else
                setMargin.top = defMargin.top
            End If
            setMargin.hasbeenset = True
            MarginDisplay
        Case 4  'bottom margin
            If setMargin.bottom > (defMargin.bottom + MarginAdjust) Then
                setMargin.bottom = setMargin.bottom - MarginAdjust
            Else
                setMargin.bottom = defMargin.bottom
            End If
            setMargin.hasbeenset = True
            MarginDisplay
        Case 5  'left margin
            If setMargin.left > (defMargin.left + MarginAdjust) Then
                setMargin.left = setMargin.left - MarginAdjust
            Else
                setMargin.left = defMargin.left
            End If
            setMargin.hasbeenset = True
            MarginDisplay
        Case 6  'righ margin
            If setMargin.right > (defMargin.right + MarginAdjust) Then
                setMargin.right = setMargin.right - MarginAdjust
            Else
                setMargin.right = defMargin.right
            End If
            setMargin.hasbeenset = True
            MarginDisplay
        Case 7  'Font size factor
            If CSng(Text1(9).Text) >= 1.1 Then
                Text1(9).Text = Format(CSng(Text1(9).Text) - 0.1, "#0.0")
            End If
    End Select
End Sub


Private Sub UpDown1_GotFocus(Index As Integer)
    Select Case Index
        Case 3, 4, 5, 6
            CheckNeeded = True
    End Select
End Sub

Private Sub UpDown1_LostFocus(Index As Integer)
    Select Case Index
        Case 3, 4, 5, 6, 9
            CountPages True
    End Select
End Sub


Private Sub UpDown1_UpClick(Index As Integer)
    If Text1(Index).Text = "" Then
        Text1(Index).Text = "0"
    End If
    Select Case Index
        Case 0  'pages Wide
            If CInt(Text1(Index).Text) < defPagesWide Then
                Text1(Index).Text = CInt(Text1(Index).Text) + 1
            End If
            userPagesWide = CLng(Text1(Index).Text)
            StatusBarPages
        Case 1      'pages High
            If CInt(Text1(Index).Text) < defPagesHigh Then
                Text1(Index).Text = CInt(Text1(Index).Text) + 1
            End If
            userPagesHigh = CLng(Text1(Index).Text)
            StatusBarPages
        Case 2      'Number of Copies
            If IsNumeric(Text1(Index).Text) Then
                Text1(Index).Text = CInt(Text1(Index).Text) + 1
            Else
                Text1(Index).Text = "2"
            End If
            StatusBar1.Panels(3).Text = "Copies: " & Text1(Index).Text
        Case 3
            If setMargin.top < (defMargin.top + defMargin.height + MarginAdjust) Then
                setMargin.top = setMargin.top + MarginAdjust
            End If
            setMargin.hasbeenset = True
            MarginDisplay
        Case 4
            If setMargin.bottom < (defMargin.bottom + defMargin.height + MarginAdjust) Then
                setMargin.bottom = setMargin.bottom + MarginAdjust
            End If
            setMargin.hasbeenset = True
            MarginDisplay
        Case 5
            If setMargin.left < (defMargin.left + defMargin.width + MarginAdjust) Then
                setMargin.left = setMargin.left + MarginAdjust
            End If
            setMargin.hasbeenset = True
            MarginDisplay
        Case 6
            If setMargin.right < (defMargin.right + defMargin.width + MarginAdjust) Then
                setMargin.right = setMargin.right + MarginAdjust
            End If
            setMargin.hasbeenset = True
            MarginDisplay
        Case 7
            Text1(9).Text = Format(CSng(Text1(9).Text) + 0.1, "#0.0")
    End Select
End Sub



Public Property Let TitlePages(ByVal eNewValue As TitleOption)
    mTitlePages = eNewValue
End Property

Public Property Let TitleFontSize(ByVal dNewValue As Double)
    'Property now depracated in favour of setting the TotleFont object
    If dNewValue < 1 Then
        dNewValue = 1
    End If
    If dNewValue > 3 Then
        dNewValue = 3
    End If
    If pFlexgrid Is Nothing Then 'not set yet so use default
        TitleFont.Size = 8 * dNewValue
    Else    'use the grid font size
        TitleFont.Size = pFlexgrid.Font.Size * dNewValue
    End If
End Property


Public Property Let SetEffects(ByVal eNewValue As Effects)
    mEffects = eNewValue
End Property


Public Property Let PageNumbering(ByVal eNewValue As PageNumbers)
    mPageNumbering = eNewValue
End Property


Public Property Let MergeRule(ByVal eNewValue As CellMerge)
    mMergeRule = eNewValue
End Property

Public Property Let SetPrinter(ByVal vNewValue As Variant)
    mSetPrinter = vNewValue
End Property

Public Property Get RequestPagesWide() As Long
    RequestPagesWide = mRequestPagesWide
End Property

Public Property Let RequestPagesWide(ByVal lNewValue As Long)
    If lNewValue >= 1 Then
        mRequestPagesWide = lNewValue
    Else
        mRequestPagesWide = 0
    End If
End Property

Public Property Get RequestPagesHigh() As Long
    RequestPagesHigh = mRequestPagesHigh
End Property

Public Property Let RequestPagesHigh(ByVal lNewValue As Long)
    If lNewValue >= 1 Then
        mRequestPagesHigh = lNewValue
    Else
        mRequestPagesHigh = 0
    End If
End Property

Public Property Get PrintComplete() As Boolean
    PrintComplete = mPrintComplete
End Property

Public Property Get SetPages() As Boolean
    SetPages = mSetPages
End Property

Public Property Let SetPages(ByVal bNewValue As Boolean)
    If mRequestPagesHigh > 0 And mRequestPagesWide > 0 Then
        mSetPages = bNewValue
    Else
        mSetPages = False
    End If
End Property

Public Property Let PrinterOrientation(ByVal PGNewValue As PageOrientation)
    mPrinterOrientation = PGNewValue
End Property

Public Property Get PrintProgress() As Long
    PrintProgress = mPrintProgress
End Property
Public Property Let ShowProgress(ByVal bNewValue As Boolean)
    mShowProgress = bNewValue
End Property
Private Sub optPPS_Click(Index As Integer)
    On Error Resume Next
    If changingPPS Then
        Exit Sub
    Else
        changingPPS = True
    End If
    If optPPS(0).Value Then
        POOrientation = cdlPortrait
    Else
        POOrientation = cdlLandscape
    End If
    Printer.Orientation = POOrientation
    DoEvents
    SetDefmargins
    CountPages True 'number of print pages is likely to change
    SetPrintData
    changingPPS = False
End Sub
Private Sub cbPaperSize_Click()
    If cbPaperSize.ItemData(cbPaperSize.ListIndex) = -1 Then
        Exit Sub
    End If
    Printer.PaperSize = cbPaperSize.ItemData(cbPaperSize.ListIndex)
    SetPrintData
End Sub
Private Sub LoadPaperSize()
    Dim TypeLoop As Integer
    
    GetPaperSizes
    cbPaperSize.Clear
    For TypeLoop = 1 To UBound(PaperTypes)
        cbPaperSize.AddItem PaperTypes(TypeLoop).Name
        cbPaperSize.ItemData(cbPaperSize.NewIndex) = PaperTypes(TypeLoop).Index
    Next TypeLoop
    cbPaperSize.Text = pText

End Sub
Private Sub GetPaperSizes()
'Get Current Printer PaperSizes
    Dim TypeLoop As Long, SuportedPapers As Long, NameCount As Long
    Dim PSizeCount As Long, PsArrayC As Long
    Dim Papers() As Integer
    Dim pSizes() As POINTAPI
    Dim pprstr As String, PaperSizes() As String

    SuportedPapers = DeviceCapabilities(Printer.DeviceName, Printer.Port, DC_PAPERS, ByVal vbNullString, 0)
    ReDim PaperTypes(1 To SuportedPapers)
    ReDim pSizes(1 To SuportedPapers)
    ReDim Papers(1 To SuportedPapers)
    'PaperSize Names
    NameCount = DeviceCapabilities(Printer.DeviceName, Printer.Port, DC_PAPERNAMES, ByVal vbNullString, 0)
    If NameCount <> 0 Then
        pprstr = String(NameCount * 64, " ")
        NameCount = DeviceCapabilities(Printer.DeviceName, Printer.Port, DC_PAPERNAMES, ByVal pprstr, 0)
    End If
    'PaperSize index
    PSizeCount = DeviceCapabilities(Printer.DeviceName, Printer.Port, DC_PAPERS, ByVal vbNullString, 0)
    If PSizeCount <> 0 Then
        PSizeCount = DeviceCapabilities(Printer.DeviceName, Printer.Port, DC_PAPERS, Papers(1), 0)
    End If
    'PaperSize Dimensions
    PsArrayC = DeviceCapabilities(Printer.DeviceName, Printer.Port, DC_PAPERSIZE, ByVal vbNullString, 0)
    If PsArrayC <> 0 Then
        PsArrayC = DeviceCapabilities(Printer.DeviceName, Printer.Port, DC_PAPERSIZE, pSizes(1), 0)
    End If
    For TypeLoop = 1 To SuportedPapers
       If PsArrayC > 0 Then
            PaperTypes(TypeLoop).width = pSizes(TypeLoop).x / 10 'to store as mm
            PaperTypes(TypeLoop).height = pSizes(TypeLoop).y / 10
       End If
       If NameCount > 0 Then
            PaperTypes(TypeLoop).Name = Mid$(pprstr, (TypeLoop - 1) * 64 + 1, 64)
            PaperTypes(TypeLoop).Name = Mid$(PaperTypes(TypeLoop).Name, 1, InStr(1, PaperTypes(TypeLoop).Name, Chr$(0)) - 1)
       Else
            PaperTypes(TypeLoop).Name = PaperSizeStr(TypeLoop)
       End If
       If PSizeCount > 0 Then
            PaperTypes(TypeLoop).Index = Papers(TypeLoop)
            If PaperTypes(TypeLoop).Index = 256 And NameCount = 0 Then
                PaperTypes(TypeLoop).Name = "Custom"
            End If
       Else
            PaperTypes(TypeLoop).Index = "-1"
       End If
    Next TypeLoop
End Sub
Private Function PaperSizeStr(PaperIndex As Long) As String
    Dim WidthVal As Single, HeightVal As Single
    
    WidthVal = PaperTypes(PaperIndex).width
    HeightVal = PaperTypes(PaperIndex).height
    Select Case PaperSize.MSystem
        Case 1  'USA measurements
            WidthVal = (WidthVal * TwipsMM) / 1440
            HeightVal = (HeightVal * TwipsMM) / 1440
            PaperSizeStr = Trim(Str(Round(WidthVal, 2))) & "in x " & Trim(Str(Round(HeightVal, 2))) & "in"
        Case Else 'metric
            PaperSizeStr = Trim(Str(Round(WidthVal, 1))) & "mm x " & Trim(Str(Round(HeightVal, 1))) & "mm"
    End Select
End Function

Private Function pText() As String
'Returns paper size name for the current printer as it is known to the relevant printer driver
    Dim SizeIndex As Long, TypeLoop As Long

    SizeIndex = Printer.PaperSize
    pText = "Unknown"
    For TypeLoop = 1 To UBound(PaperTypes)
        If SizeIndex = PaperTypes(TypeLoop).Index Then
            pText = PaperTypes(TypeLoop).Name
            Exit For
        End If
    Next TypeLoop
End Function

Public Property Let SelectionPrintRule(ByVal eNewValue As PrintSelect)
    mSelectionPrintRule = eNewValue
End Property

Public Property Let MultiColumnPrint(ByVal eNewValue As MultiColumn)
    mMultiColumnPrint = eNewValue
End Property

Public Property Get RequestSectionsWide() As Long
    RequestSectionsWide = mRequestSectionsWide
End Property


Public Property Get RequestSectionsHigh() As Variant
    RequestSectionsHigh = mRequestSectionsHigh
End Property

Private Sub RepeatColHeadings(PageW As Long, YOffset As Single, XOffset As Single)
    Dim RowLoop As Long
    
    Printer.CurrentY = Header.NextTitle + Printer.ScaleY((setMargin.top - defMargin.top), vbTwips, Printer.ScaleMode)
    If PrintMerge.MergeCols Then    'go see if there are any in the fixed rows
        FindMergedCol 0, (pFlexgrid.FixedRows - 1), PagesW(PageW).GridStart, PagesW(PageW).GridEnd
        PrintMergedCols YOffset, XOffset
    End If
    'we need to know which columns to print
    For RowLoop = 0 To (pFlexgrid.FixedRows - 1)
        If RowHeights(RowLoop) > 0 Then
            PrintRowOnPage RowLoop, PageW, XOffset
            Printer.CurrentY = Printer.CurrentY + RowHeights(RowLoop)
        End If
    Next RowLoop

End Sub
Private Sub CalcFixedWidth()
    Dim GridPos() As Variant
    Dim ColLoop As Long
    Dim FixWidth As Single
    
    If Not FixedColData.RepeatFixed Then
        FixedColData.FixedPage1 = 0
        FixedColData.FixedWidth = 0
        Exit Sub
    End If

    FixedColData.HighCol = pFlexgrid.FixedCols - 1
    If FixedColData.HighCol >= 0 Then
        For ColLoop = 0 To FixedColData.HighCol
            FixWidth = FixWidth + ColWidths(ColLoop)
        Next ColLoop
    End If
    FixedColData.FixedWidth = FixWidth
    FixedColData.FixedPage1 = 0
    If mPrintSelected Then
        GridPos() = DefSavePosition
        If GridPos(1) > FixedColData.HighCol Then
            FixedColData.FixedPage1 = FixedColData.FixedWidth
        End If
    End If
End Sub

Public Property Let SubHeadings(ByVal vNewValue As Variant)
    Dim SubLoop As Integer, SubCount As Integer
    Dim workVar() As Variant
    If Not IsEmpty(vNewValue) Then
        workVar() = vNewValue
        mySubHeadings = workVar()
        AppFont pFlexgrid.Font, mySubHeadFont 'make sure this is initialised
    End If
End Property

Public Property Let SubHeadFont(ByVal fFontSettings As StdFont)
    AppFont fFontSettings, mySubHeadFont
End Property
Public Property Let SubHeadJustify(ByVal jNewValue As JustifySubTitle)
    mySubJustify = jNewValue
End Property

Private Sub AppFont(ByRef FromFont As StdFont, ByRef ToFont As StdFont)

    On Error Resume Next
    'this sub shows signs that VB6 is getting old - why can't you just equate the objects?
    ToFont.Name = FromFont.Name
    ToFont.Size = FromFont.Size
    If FromFont.Size < 8 Then
        'probably dont need this with the font object
        ToFont.Name = FromFont.Name
        ToFont.Size = FromFont.Size
    End If
    ToFont.Bold = FromFont.Bold
    ToFont.Italic = FromFont.Italic
    ToFont.Strikethrough = FromFont.Strikethrough
    ToFont.Underline = FromFont.Underline
    ToFont.Charset = FromFont.Charset
    ToFont.Weight = FromFont.Weight
End Sub

Public Property Get PrintCellImages() As Boolean
    PrintCellImages = mPrintCellImages
End Property

Public Property Let PrintCellImages(ByVal bNewValue As Boolean)
    mPrintCellImages = bNewValue
End Property
Public Property Let SetTitleFont(ByVal fFontSetting As StdFont)

    AppFont fFontSetting, TitleFont
    
End Property
Public Sub SetPrintLimits(ByVal lColStart As Long, ByVal lColEnd As Long, ByVal lRowStart As Long, ByVal lRowEnd As Long)
    If lColStart >= 0 Then
        mPLColStart = lColStart
    End If
    If lColEnd >= lColStart Then
        mPLColEnd = lColEnd
    End If
    If lRowStart >= 0 Then
        mPLRowStart = lRowStart
    End If
    If lRowEnd >= lRowStart Then
        mPLRowEnd = lRowEnd
    End If
End Sub

Public Property Let SubTitleRepeat(ByVal eNewValue As SubTitleUsage)
    mRepeatSubTitle = eNewValue
End Property
Private Sub LoadTitleFont()
    'Allows user selection of the Title Font
    
    On Error GoTo LoadTitleFontCancel
    CommonDialog1.CancelError = True
    CommonDialog1.DialogTitle = "Select Title Font"
    CommonDialog1.Flags = cdlCFBoth
    'set the dialogue defaults to current settings
    CommonDialog1.FontName = TitleFont.Name
    CommonDialog1.FontBold = TitleFont.Bold
    CommonDialog1.FontItalic = TitleFont.Italic
    CommonDialog1.FontSize = TitleFont.Size
    CommonDialog1.FontStrikethru = TitleFont.Strikethrough
    CommonDialog1.FontUnderline = TitleFont.Underline
    
    CommonDialog1.ShowFont
    
    TitleFont.Name = CommonDialog1.FontName
    TitleFont.Bold = CommonDialog1.FontBold
    TitleFont.Italic = CommonDialog1.FontItalic
    TitleFont.Size = CommonDialog1.FontSize
    TitleFont.Strikethrough = CommonDialog1.FontStrikethru
    TitleFont.Underline = CommonDialog1.FontUnderline
    AppFont TitleFont, Text1(3).Font
    CountPages True
    Exit Sub
LoadTitleFontCancel:
    Exit Sub
End Sub


Public Property Get ProportionalCompression() As Boolean
    ProportionalCompression = mProportionalCompression
End Property

Public Property Let ProportionalCompression(ByVal bNewValue As Boolean)
    mProportionalCompression = bNewValue
End Property
