VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00F3E5CE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "3 Dimensional Shapes!"
   ClientHeight    =   6150
   ClientLeft      =   5640
   ClientTop       =   4845
   ClientWidth     =   8265
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   6150
   ScaleWidth      =   8265
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdcal 
      BackColor       =   &H00F3E5CE&
      Caption         =   "Calculate"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   51
      ToolTipText     =   "Calculate the Volume"
      Top             =   5520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame frm3d 
      BackColor       =   &H00F3E5CE&
      Caption         =   "3D Shape Properties"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   4440
      TabIndex        =   20
      Top             =   120
      Visible         =   0   'False
      Width           =   3735
      Begin VB.ComboBox cmbshapes2 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "shapes3d.frx":0000
         Left            =   1080
         List            =   "shapes3d.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   27
         ToolTipText     =   "Please Select Your 3D Shape"
         Top             =   240
         Width           =   2535
      End
      Begin VB.TextBox txtformula2 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   1080
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   31
         ToolTipText     =   "Formula for the Shape"
         Top             =   2160
         Width           =   2535
      End
      Begin VB.TextBox txtcorner2 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   29
         ToolTipText     =   "Number of Corners"
         Top             =   1200
         Width           =   2535
      End
      Begin VB.TextBox txtside2 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   28
         ToolTipText     =   "Number of Sides"
         Top             =   720
         Width           =   2535
      End
      Begin VB.TextBox txtface 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   30
         ToolTipText     =   "Number of Faces"
         Top             =   1680
         Width           =   2535
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Volume = "
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
         Left            =   120
         TabIndex        =   26
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Corners:"
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
         Left            =   120
         TabIndex        =   25
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Sides:"
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
         Left            =   120
         TabIndex        =   23
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Faces:"
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
         Left            =   120
         TabIndex        =   22
         Top             =   1680
         Width           =   495
      End
      Begin VB.Image img3d 
         Height          =   1575
         Left            =   0
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame frmins 
      BackColor       =   &H00F3E5CE&
      Caption         =   "Instructions"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   4440
      TabIndex        =   60
      Top             =   3120
      Visible         =   0   'False
      Width           =   3735
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "3. Click Back to Select Another Shape"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   63
         Top             =   1800
         Width           =   3255
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "2. Click the Calculate button in order to determine the volume for your shape"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   62
         Top             =   960
         Width           =   3375
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "1. Enter any digit or digits of 0-9 into the text boxes."
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   61
         Top             =   360
         Width           =   3255
      End
      Begin VB.Image imgins 
         Height          =   1815
         Left            =   0
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame frmmain 
      Caption         =   "3D Shapes"
      Height          =   5895
      Left            =   120
      TabIndex        =   32
      Top             =   120
      Width           =   8055
      Begin VB.CommandButton Command3 
         BackColor       =   &H00F3E5CE&
         Caption         =   "Back"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Go Back"
         Top             =   5400
         Width           =   1215
      End
      Begin VB.CommandButton cmdcir 
         BackColor       =   &H00F3E5CE&
         Caption         =   "Sphere"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   1200
         Width           =   2295
      End
      Begin VB.CommandButton cmdsqu 
         BackColor       =   &H00F3E5CE&
         Caption         =   "Cube"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   1200
         Width           =   2295
      End
      Begin VB.CommandButton cmdtri 
         BackColor       =   &H00F3E5CE&
         Caption         =   "Triangle Pyramid"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   2400
         Width           =   2295
      End
      Begin VB.CommandButton cmdrec 
         BackColor       =   &H00F3E5CE&
         Caption         =   "Rectangular Prism"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   2400
         Width           =   2295
      End
      Begin VB.CommandButton cmdcone 
         BackColor       =   &H00F3E5CE&
         Caption         =   "Cone"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   3600
         Width           =   2295
      End
      Begin VB.CommandButton cmdcyl 
         BackColor       =   &H00F3E5CE&
         Caption         =   "Cylinder"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   3600
         Width           =   2295
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Select Your Three Dimensional Shape"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         TabIndex        =   34
         Top             =   360
         Width           =   6615
      End
      Begin VB.Image imgback 
         Height          =   5955
         Left            =   0
         Picture         =   "shapes3d.frx":005F
         Top             =   0
         Width           =   8010
      End
   End
   Begin VB.Frame frmcyl 
      BackColor       =   &H00F3E5CE&
      Caption         =   "Cylinder"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   120
      TabIndex        =   64
      Top             =   120
      Visible         =   0   'False
      Width           =   4095
      Begin VB.CommandButton cmdback6 
         BackColor       =   &H00F3E5CE&
         Caption         =   "Back"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   71
         ToolTipText     =   "Go Back"
         Top             =   5400
         Width           =   1215
      End
      Begin VB.TextBox txtcyl3 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   70
         ToolTipText     =   "Volume"
         Top             =   4680
         Width           =   1695
      End
      Begin VB.TextBox txtcyl2 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         TabIndex        =   68
         ToolTipText     =   "Enter Height"
         Top             =   4200
         Width           =   1695
      End
      Begin VB.TextBox txtcyl1 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         TabIndex        =   67
         ToolTipText     =   "Enter Radius"
         Top             =   3720
         Width           =   1695
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Volume:"
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
         TabIndex        =   69
         Top             =   4680
         Width           =   855
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Height:"
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
         TabIndex        =   66
         Top             =   4200
         Width           =   615
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Radius:"
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
         TabIndex        =   65
         Top             =   3720
         Width           =   615
      End
      Begin VB.Image Image3 
         Height          =   3180
         Left            =   120
         Picture         =   "shapes3d.frx":E81D
         ToolTipText     =   "Cylinder"
         Top             =   240
         Width           =   3825
      End
      Begin VB.Image imgcyl 
         Height          =   1695
         Left            =   0
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame frmcone 
      BackColor       =   &H00F3E5CE&
      Caption         =   "Cone"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   120
      TabIndex        =   52
      Top             =   120
      Visible         =   0   'False
      Width           =   4080
      Begin VB.TextBox txtcone3 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   59
         ToolTipText     =   "Cone"
         Top             =   4920
         Width           =   1695
      End
      Begin VB.TextBox txtcone2 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         TabIndex        =   57
         ToolTipText     =   "Enter Height"
         Top             =   4440
         Width           =   1695
      End
      Begin VB.TextBox txtcone1 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         TabIndex        =   56
         ToolTipText     =   "Enter Radius"
         Top             =   3960
         Width           =   1695
      End
      Begin VB.CommandButton cmdback5 
         BackColor       =   &H00F3E5CE&
         Caption         =   "Back"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Go Back"
         Top             =   5400
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Volume:"
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
         TabIndex        =   58
         Top             =   4920
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Height:"
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
         TabIndex        =   55
         Top             =   4440
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Radius:"
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
         TabIndex        =   54
         Top             =   3960
         Width           =   615
      End
      Begin VB.Image Image1 
         Height          =   3480
         Left            =   120
         Picture         =   "shapes3d.frx":1DAB6
         ToolTipText     =   "Cone"
         Top             =   240
         Width           =   3825
      End
      Begin VB.Image imgcon 
         Height          =   1695
         Left            =   0
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame frmsqu 
      BackColor       =   &H00F3E5CE&
      Caption         =   "Cube"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   4095
      Begin VB.TextBox txtsqu3 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         TabIndex        =   19
         ToolTipText     =   "Enter Side"
         Top             =   3480
         Width           =   1695
      End
      Begin VB.TextBox txtsqu5 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   21
         ToolTipText     =   "Volume"
         Top             =   3960
         Width           =   1695
      End
      Begin VB.CommandButton cmdback2 
         BackColor       =   &H00F3E5CE&
         Caption         =   "Back"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Go Back"
         Top             =   5400
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Volume:"
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
         TabIndex        =   2
         Top             =   3960
         Width           =   735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Side:"
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
         TabIndex        =   1
         Top             =   3480
         Width           =   615
      End
      Begin VB.Image Image2 
         Height          =   3000
         Left            =   120
         Picture         =   "shapes3d.frx":2E3E6
         ToolTipText     =   "Cube"
         Top             =   240
         Width           =   3825
      End
      Begin VB.Image imgcub 
         Height          =   1095
         Left            =   0
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame frmrec 
      BackColor       =   &H00F3E5CE&
      Caption         =   "Rectangular Prism"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   120
      TabIndex        =   36
      Top             =   120
      Visible         =   0   'False
      Width           =   4095
      Begin VB.CommandButton cmdback4 
         BackColor       =   &H00F3E5CE&
         Caption         =   "Back"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Go Back"
         Top             =   5400
         Width           =   1215
      End
      Begin VB.TextBox txtrec7 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   49
         ToolTipText     =   "Volume"
         Top             =   4560
         Width           =   1695
      End
      Begin VB.TextBox txtrec6 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         TabIndex        =   48
         ToolTipText     =   "Enter Height"
         Top             =   4080
         Width           =   1695
      End
      Begin VB.TextBox txtrec5 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         TabIndex        =   47
         ToolTipText     =   "Enter Length"
         Top             =   3600
         Width           =   1695
      End
      Begin VB.TextBox txtrec4 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         TabIndex        =   46
         ToolTipText     =   "Enter Width"
         Top             =   3120
         Width           =   1695
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Volume:"
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
         TabIndex        =   43
         Top             =   4560
         Width           =   735
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Height:"
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
         TabIndex        =   42
         Top             =   4080
         Width           =   615
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "Length:"
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
         TabIndex        =   40
         Top             =   3600
         Width           =   735
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Width:"
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
         TabIndex        =   38
         Top             =   3120
         Width           =   735
      End
      Begin VB.Image Image8 
         Height          =   2460
         Left            =   120
         Picture         =   "shapes3d.frx":3C8F1
         ToolTipText     =   "Rectangular Prism"
         Top             =   360
         Width           =   3840
      End
      Begin VB.Image imgrec 
         Height          =   1215
         Left            =   0
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame frmtri 
      BackColor       =   &H00F3E5CE&
      Caption         =   "Triangular Pyramid"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   4095
      Begin VB.CommandButton cmdback3 
         BackColor       =   &H00F3E5CE&
         Caption         =   "Back"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Go Back"
         Top             =   5400
         Width           =   1215
      End
      Begin VB.TextBox txttri7 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   17
         ToolTipText     =   "Volume"
         Top             =   4680
         Width           =   1695
      End
      Begin VB.TextBox txttri6 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         TabIndex        =   16
         ToolTipText     =   "Enter Height"
         Top             =   4200
         Width           =   1695
      End
      Begin VB.TextBox txttri5 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         TabIndex        =   15
         ToolTipText     =   "Enter Length"
         Top             =   3720
         Width           =   1695
      End
      Begin VB.TextBox txttri4 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         TabIndex        =   14
         ToolTipText     =   "Enter Width"
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Image Image7 
         Height          =   2730
         Left            =   120
         Picture         =   "shapes3d.frx":4860D
         ToolTipText     =   "Triangular Pyramid"
         Top             =   240
         Width           =   3810
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Volume:"
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
         TabIndex        =   13
         Top             =   4680
         Width           =   735
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Height:"
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
         TabIndex        =   12
         Top             =   4200
         Width           =   615
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Length:"
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
         TabIndex        =   11
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Width:"
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
         TabIndex        =   10
         Top             =   3240
         Width           =   975
      End
      Begin VB.Image imgtri 
         Height          =   1695
         Left            =   0
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame frmcir 
      BackColor       =   &H00F3E5CE&
      Caption         =   "Sphere"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   4095
      Begin VB.CommandButton cmdback1 
         BackColor       =   &H00F3E5CE&
         Caption         =   "Back"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Go Back"
         Top             =   5400
         Width           =   1215
      End
      Begin VB.TextBox txtcir4 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   7
         ToolTipText     =   "Volume"
         Top             =   3720
         Width           =   1695
      End
      Begin VB.TextBox txtcir3 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         TabIndex        =   6
         ToolTipText     =   "Enter Radius"
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Volume:"
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
         TabIndex        =   5
         Top             =   3720
         Width           =   735
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Radius:"
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
         TabIndex        =   4
         Top             =   3240
         Width           =   615
      End
      Begin VB.Image Image5 
         Height          =   2715
         Left            =   120
         Picture         =   "shapes3d.frx":555F5
         ToolTipText     =   "Sphere"
         Top             =   240
         Width           =   3840
      End
      Begin VB.Image imgsph 
         Height          =   1335
         Left            =   0
         Top             =   240
         Width           =   2535
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbshapes2_Click()
    Select Case cmbshapes2.ListIndex
        Case Is = "0"
            back2
            frmcir.Visible = True
            txtside2.Text = "1"
            txtface.Text = "1"
            txtcorner2.Text = "0"
            txtformula2.Text = "4 ÷ 3 x Pi x Radius³"
        Case Is = "1"
            back2
            frmsqu.Visible = True
            txtside2.Text = "12"
            txtface.Text = "6"
            txtcorner2.Text = "8"
            txtformula2.Text = "Side³"
        Case Is = "2"
            back2
            frmtri.Visible = True
            txtside2.Text = "6"
            txtface.Text = "4"
            txtcorner2.Text = "4"
            txtformula2.Text = "(Length x Width x Height) ÷ 6"
        Case Is = "3"
            back2
            frmrec.Visible = True
            txtside2.Text = "12"
            txtface.Text = "6"
            txtcorner2.Text = "8"
            txtformula2.Text = "Length x Width x Height"
        Case Is = "4"
            back2
            frmcone.Visible = True
            txtside2.Text = "1"
            txtface.Text = "2"
            txtcorner2.Text = "1"
            txtformula2.Text = "(Circle Area x Height) ÷ 3"
        Case Is = "5"
            back2
            frmcyl.Visible = True
            txtside2.Text = "2"
            txtface.Text = "3"
            txtcorner2.Text = "0"
            txtformula2.Text = "Circle Area x Height"
    End Select
End Sub

Private Sub cmdcal_Click()
    calculate
End Sub

Private Sub cmdcir_Click()
    select_shape
    cmbshapes2.ListIndex = "0"
End Sub

Private Sub cmdsqu_Click()
    select_shape
    cmbshapes2.ListIndex = "1"
End Sub

Private Sub cmdtri_Click()
    select_shape
    cmbshapes2.ListIndex = "2"
End Sub

Private Sub cmdrec_Click()
    select_shape
    cmbshapes2.ListIndex = "3"
End Sub

Private Sub cmdcone_Click()
    select_shape
    cmbshapes2.ListIndex = "4"
End Sub

Private Sub cmdcyl_Click()
    select_shape
    cmbshapes2.ListIndex = "5"
End Sub

Private Sub cmdback1_Click()
    frmcir.Visible = False
    back
End Sub

Private Sub cmdback2_Click()
    frmsqu.Visible = False
    back
End Sub

Private Sub cmdback3_Click()
    frmtri.Visible = False
    back
End Sub

Private Sub cmdback4_Click()
    frmrec.Visible = False
    back
End Sub

Private Sub cmdback5_Click()
    frmcone.Visible = False
    back
End Sub

Private Sub cmdback6_Click()
    frmcyl.Visible = False
    back
End Sub

Private Sub Command3_Click()
    Form3.Hide
    Form1.Show
End Sub

Private Sub Form_QueryUnLoad(cancel As Integer, UnloadMode As Integer)
    If MsgBox("Are you sure you want to EXIT?", vbQuestion + vbYesNo, "Exit") = vbNo Then
        cancel = True
    End If
End Sub
Private Sub form_unload(cancel As Integer)
    End
End Sub

Private Sub txtcir3_keypress(KeyAscii As Integer)
    If abc(KeyAscii) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub


Private Sub txtcone1_keypress(KeyAscii As Integer)
    If abc(KeyAscii) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtcone2_keypress(KeyAscii As Integer)
    If abc(KeyAscii) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtcyl1_keypress(KeyAscii As Integer)
    If abc(KeyAscii) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtcyl2_keypress(KeyAscii As Integer)
    If abc(KeyAscii) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtrec4_keypress(KeyAscii As Integer)
    If abc(KeyAscii) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtrec5_keypress(KeyAscii As Integer)
    If abc(KeyAscii) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtrec6_keypress(KeyAscii As Integer)
    If abc(KeyAscii) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtsqu3_keypress(KeyAscii As Integer)
    If abc(KeyAscii) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txttri4_keypress(KeyAscii As Integer)
    If abc(KeyAscii) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txttri5_keypress(KeyAscii As Integer)
    If abc(KeyAscii) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txttri6_keypress(KeyAscii As Integer)
    If abc(KeyAscii) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub
Private Sub calculate()
    Const Pi = 3.14
    If Not txtcir3.Text = "" Then
        txtcir4 = Format(4 / 3 * Pi * (txtcir3.Text) ^ 3, "standard") & " units³"
    End If
    If Not txtsqu3.Text = "" Then
        txtsqu5 = Format((txtsqu3.Text) ^ 3, "#,###") & " units³"
    End If
    If Not txttri4.Text = "" And Not txttri5.Text = "" And Not txttri6.Text = "" Then
        txttri7 = Format((txttri4.Text) * (txttri5.Text) / 2 * (txttri6.Text) / 3, "standard") & " units³"
    End If
    If Not txtrec4.Text = "" And Not txtrec5.Text = "" And Not txtrec6.Text = "" Then
        txtrec7 = Format((txtrec4.Text) * (txtrec5.Text) * (txtrec6.Text), "#,###") & " units³"
    End If
    If Not txtcone1.Text = "" And Not txtcone2.Text = "" Then
        txtcone3 = Format(Pi * (txtcone1.Text) ^ 2 * 1 / 3 * (txtcone2.Text), "standard") & " units²"
    End If
    If Not txtcyl1.Text = "" And Not txtcyl2.Text = "" Then
        txtcyl3 = Format(Pi * (txtcyl1.Text) ^ 2 * (txtcyl2.Text), "standard") & " units²"
    End If
End Sub
Private Function abc(KeyAscii As Integer) As Boolean
    If KeyAscii > 47 And KeyAscii < 58 Or KeyAscii = 8 Then
       abc = True
    End If
    If KeyAscii = 13 Then
        abc = False
        calculate
    End If
End Function

Private Sub select_shape()
    frmmain.Visible = False
    frm3d.Visible = True
    frmins.Visible = True
    cmdcal.Visible = True
End Sub

Private Sub back()
    frmmain.Visible = True
    frm3d.Visible = False
    frmins.Visible = False
    cmdcal.Visible = False
    cmbshapes2.ListIndex = "-1"
End Sub
Private Sub back2()
    frmtri.Visible = False
    frmrec.Visible = False
    frmsqu.Visible = False
    frmcir.Visible = False
    frmcone.Visible = False
    frmcyl.Visible = False
End Sub

Private Sub Form_Load()
    Form3.Icon = Form1.Icon
    imgtri.Picture = imgback.Picture
    imgrec.Picture = imgback.Picture
    imgcon.Picture = imgback.Picture
    imgcub.Picture = imgback.Picture
    imgsph.Picture = imgback.Picture
    img3d.Picture = imgback.Picture
    imgins.Picture = imgback.Picture
    imgcyl.Picture = imgback.Picture
End Sub
