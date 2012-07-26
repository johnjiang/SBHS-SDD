VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00F3E5CE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "2 Dimensional Shapes!"
   ClientHeight    =   6150
   ClientLeft      =   5640
   ClientTop       =   4845
   ClientWidth     =   8265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   6150
   ScaleWidth      =   8265
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   WhatsThisHelp   =   -1  'True
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
      TabIndex        =   63
      ToolTipText     =   "Calculate the Area"
      Top             =   5520
      Visible         =   0   'False
      Width           =   1215
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
      Height          =   3015
      Left            =   4440
      TabIndex        =   64
      Top             =   3000
      Visible         =   0   'False
      Width           =   3735
      Begin VB.Label Label20 
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
         TabIndex        =   67
         Top             =   1800
         Width           =   3255
      End
      Begin VB.Label Label18 
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
         TabIndex        =   66
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "2. Click the Calculate button in order to determine the area for your shape"
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
         TabIndex        =   65
         Top             =   960
         Width           =   3255
      End
      Begin VB.Image imgins 
         Height          =   1815
         Left            =   0
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame frm2d 
      BackColor       =   &H00F3E5CE&
      Caption         =   "2D Shape Properties"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   4440
      TabIndex        =   36
      Top             =   120
      Visible         =   0   'False
      Width           =   3735
      Begin VB.TextBox txtformula1 
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
         TabIndex        =   29
         ToolTipText     =   "Formula for the shape"
         Top             =   1800
         Width           =   2535
      End
      Begin VB.TextBox txtcorner1 
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
         ToolTipText     =   "Number of Corners"
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox txtside1 
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
         TabIndex        =   27
         ToolTipText     =   "Number of Sides"
         Top             =   840
         Width           =   2535
      End
      Begin VB.ComboBox cmbshapes 
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
         ItemData        =   "shapes2d.frx":0000
         Left            =   1080
         List            =   "shapes2d.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   26
         ToolTipText     =   "Please Select Your 2D Shape"
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label8 
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
         TabIndex        =   39
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label12 
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
         TabIndex        =   38
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Area ="
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
         TabIndex        =   37
         Top             =   1800
         Width           =   735
      End
      Begin VB.Image img2d 
         Height          =   1815
         Left            =   0
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame frmmain 
      Caption         =   "2D Shapes"
      Height          =   5895
      Left            =   120
      TabIndex        =   44
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
         TabIndex        =   51
         ToolTipText     =   "Go Back"
         Top             =   5400
         Width           =   1215
      End
      Begin VB.CommandButton cmdcir 
         BackColor       =   &H00F3E5CE&
         Caption         =   "Circle"
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
         TabIndex        =   2
         Top             =   1200
         Width           =   2295
      End
      Begin VB.CommandButton cmdsqu 
         BackColor       =   &H00F3E5CE&
         Caption         =   "Square"
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
         TabIndex        =   3
         Top             =   1200
         Width           =   2295
      End
      Begin VB.CommandButton cmdtri 
         BackColor       =   &H00F3E5CE&
         Caption         =   "Triangle"
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
         TabIndex        =   4
         Top             =   2400
         Width           =   2295
      End
      Begin VB.CommandButton cmdrec 
         BackColor       =   &H00F3E5CE&
         Caption         =   "Rectangle"
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
         TabIndex        =   5
         Top             =   2400
         Width           =   2295
      End
      Begin VB.CommandButton cmdtra 
         BackColor       =   &H00F3E5CE&
         Caption         =   "Trapezium"
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
         TabIndex        =   6
         Top             =   3600
         Width           =   2295
      End
      Begin VB.CommandButton cmdrho 
         BackColor       =   &H00F3E5CE&
         Caption         =   "Rhombus"
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
         TabIndex        =   7
         Top             =   3600
         Width           =   2295
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Select Your Two Dimensional Shape"
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
         TabIndex        =   45
         Top             =   360
         Width           =   6495
      End
      Begin VB.Image imgback 
         Height          =   5955
         Left            =   0
         Picture         =   "shapes2d.frx":0053
         Top             =   0
         Width           =   8010
      End
   End
   Begin VB.Frame frmsqu 
      BackColor       =   &H00F3E5CE&
      Caption         =   "Square"
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
      Begin VB.TextBox txtsqu2 
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
         TabIndex        =   12
         ToolTipText     =   "Area"
         Top             =   3480
         Width           =   1695
      End
      Begin VB.CommandButton Command7 
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
         TabIndex        =   34
         ToolTipText     =   "Go Back"
         Top             =   5400
         Width           =   1215
      End
      Begin VB.TextBox txtsqu1 
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
         TabIndex        =   10
         ToolTipText     =   "Enter Side"
         Top             =   3120
         Width           =   1695
      End
      Begin VB.Image Image1 
         Height          =   2460
         Left            =   120
         Picture         =   "shapes2d.frx":E811
         ToolTipText     =   "Square"
         Top             =   240
         Width           =   3825
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Area:"
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
         Top             =   3480
         Width           =   495
      End
      Begin VB.Label Label1 
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
         Top             =   3120
         Width           =   615
      End
      Begin VB.Image imgsqu 
         Height          =   1515
         Left            =   0
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.Frame frmrho 
      BackColor       =   &H00F3E5CE&
      Caption         =   "Rhombus"
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
      TabIndex        =   58
      Top             =   120
      Visible         =   0   'False
      Width           =   4095
      Begin VB.CommandButton Command5 
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
         TabIndex        =   62
         ToolTipText     =   "Go Back"
         Top             =   5400
         Width           =   1215
      End
      Begin VB.TextBox txtrho3 
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
         TabIndex        =   18
         ToolTipText     =   "Area"
         Top             =   4560
         Width           =   1695
      End
      Begin VB.TextBox txtrho2 
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
         TabIndex        =   17
         ToolTipText     =   "Enter Height"
         Top             =   4080
         Width           =   1695
      End
      Begin VB.TextBox txtrho1 
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
         ToolTipText     =   "Enter Length"
         Top             =   3600
         Width           =   1695
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Area:"
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
         TabIndex        =   61
         Top             =   4560
         Width           =   495
      End
      Begin VB.Label Label11 
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
         TabIndex        =   60
         Top             =   4080
         Width           =   615
      End
      Begin VB.Label Label10 
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
         TabIndex        =   59
         Top             =   3600
         Width           =   615
      End
      Begin VB.Image Image5 
         Height          =   3045
         Left            =   120
         Picture         =   "shapes2d.frx":1A474
         ToolTipText     =   "Rhombus"
         Top             =   240
         Width           =   3840
      End
      Begin VB.Image imgrho 
         Height          =   4695
         Left            =   0
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.Frame frmtra 
      BackColor       =   &H00F3E5CE&
      Caption         =   "Trapezium"
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
      Width           =   4095
      Begin VB.CommandButton Command4 
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
         TabIndex        =   57
         ToolTipText     =   "Go Back"
         Top             =   5400
         Width           =   1215
      End
      Begin VB.TextBox txttra4 
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
         TabIndex        =   22
         ToolTipText     =   "Area"
         Top             =   5040
         Width           =   1695
      End
      Begin VB.TextBox txttra3 
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
         TabIndex        =   21
         ToolTipText     =   "Enter Height"
         Top             =   4560
         Width           =   1695
      End
      Begin VB.TextBox txttra2 
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
         TabIndex        =   20
         ToolTipText     =   "Enter Side B"
         Top             =   4080
         Width           =   1695
      End
      Begin VB.TextBox txttra1 
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
         ToolTipText     =   "Enter Side A"
         Top             =   3600
         Width           =   1695
      End
      Begin VB.Label Label9 
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
         TabIndex        =   56
         Top             =   4560
         Width           =   615
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Area:"
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
         Top             =   5040
         Width           =   495
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Side B:"
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
         Top             =   4080
         Width           =   735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Side A:"
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
         TabIndex        =   53
         Top             =   3600
         Width           =   735
      End
      Begin VB.Image Image2 
         Height          =   3180
         Left            =   120
         Picture         =   "shapes2d.frx":28DC9
         ToolTipText     =   "Trapezium"
         Top             =   240
         Width           =   3870
      End
      Begin VB.Image imgtra 
         Height          =   4335
         Left            =   0
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame frmrec 
      BackColor       =   &H00F3E5CE&
      Caption         =   "Rectangle"
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
      TabIndex        =   46
      Top             =   120
      Visible         =   0   'False
      Width           =   4095
      Begin VB.TextBox txtrec1 
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
         TabIndex        =   23
         ToolTipText     =   "Enter Width"
         Top             =   3720
         Width           =   1695
      End
      Begin VB.TextBox txtrec2 
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
         TabIndex        =   24
         ToolTipText     =   "Enter Length"
         Top             =   4200
         Width           =   1695
      End
      Begin VB.TextBox txtrec3 
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
         TabIndex        =   25
         ToolTipText     =   "Area"
         Top             =   4680
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
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
         TabIndex        =   49
         ToolTipText     =   "Go Back"
         Top             =   5400
         Width           =   1215
      End
      Begin VB.Image Image6 
         Height          =   3150
         Left            =   120
         Picture         =   "shapes2d.frx":38334
         ToolTipText     =   "Rectangle"
         Top             =   240
         Width           =   3870
      End
      Begin VB.Label Label24 
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
         TabIndex        =   50
         Top             =   3720
         Width           =   495
      End
      Begin VB.Label Label25 
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
         TabIndex        =   48
         Top             =   4200
         Width           =   615
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Area:"
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
         TabIndex        =   47
         Top             =   4680
         Width           =   495
      End
      Begin VB.Image imgrec 
         Height          =   3615
         Left            =   0
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.Frame frmtri 
      BackColor       =   &H00F3E5CE&
      Caption         =   "Triangle"
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
      TabIndex        =   40
      Top             =   120
      Visible         =   0   'False
      Width           =   4095
      Begin VB.CommandButton Command1 
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
         TabIndex        =   30
         ToolTipText     =   "Go Back"
         Top             =   5400
         Width           =   1215
      End
      Begin VB.TextBox txttri3 
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
         TabIndex        =   15
         ToolTipText     =   "Area"
         Top             =   4440
         Width           =   1695
      End
      Begin VB.TextBox txttri2 
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
         ToolTipText     =   "Enter Height"
         Top             =   3960
         Width           =   1695
      End
      Begin VB.TextBox txttri1 
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
         TabIndex        =   13
         ToolTipText     =   "Enter Width"
         Top             =   3480
         Width           =   1695
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Area:"
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
         Top             =   4440
         Width           =   615
      End
      Begin VB.Label Label15 
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
         Top             =   3960
         Width           =   615
      End
      Begin VB.Label Label14 
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
         TabIndex        =   41
         Top             =   3480
         Width           =   855
      End
      Begin VB.Image Image4 
         Height          =   2910
         Left            =   120
         Picture         =   "shapes2d.frx":47655
         ToolTipText     =   "Triangle"
         Top             =   240
         Width           =   3870
      End
      Begin VB.Image imgtri 
         Height          =   3255
         Left            =   0
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame frmcir 
      BackColor       =   &H00F3E5CE&
      Caption         =   "Circle"
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
      TabIndex        =   31
      Top             =   120
      Visible         =   0   'False
      Width           =   4095
      Begin VB.CommandButton Command9 
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
         TabIndex        =   35
         ToolTipText     =   "Go Back"
         Top             =   5400
         Width           =   1215
      End
      Begin VB.TextBox txtcir2 
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
         TabIndex        =   9
         ToolTipText     =   "Area"
         Top             =   3720
         Width           =   1695
      End
      Begin VB.TextBox txtcir1 
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
         TabIndex        =   8
         ToolTipText     =   "Enter Radius"
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Image Image3 
         Height          =   2700
         Left            =   120
         Picture         =   "shapes2d.frx":5572E
         ToolTipText     =   "Circle"
         Top             =   240
         Width           =   3825
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Area:"
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
         TabIndex        =   33
         Top             =   3720
         Width           =   495
      End
      Begin VB.Label Label6 
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
         Index           =   0
         Left            =   240
         TabIndex        =   32
         Top             =   3240
         Width           =   615
      End
      Begin VB.Image imgcir 
         Height          =   3495
         Left            =   0
         Top             =   240
         Width           =   2295
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbshapes_Click()
    Select Case cmbshapes.ListIndex
        Case Is = "0"
            back2
            frmcir.Visible = True
            txtside1.Text = "1"
            txtcorner1.Text = "0"
            txtformula1.Text = "Pi x Radius²"
        Case Is = "1"
            back2
            frmsqu.Visible = True
            txtside1.Text = "4"
            txtcorner1.Text = "4"
            txtformula1.Text = "Side²"
        Case Is = "2"
            back2
            frmrec.Visible = True
            txtside1.Text = "4"
            txtcorner1.Text = "4"
            txtformula1.Text = "Height x Width"
        Case Is = "3"
            back2
            frmtri.Visible = True
            txtside1.Text = "3"
            txtcorner1.Text = "3"
            txtformula1.Text = "(Height x Width) ÷ 2"
        Case Is = "4"
            back2
            frmtra.Visible = True
            txtside1.Text = "4"
            txtcorner1.Text = "4"
            txtformula1.Text = "(SideA + SideB) ÷ 2 x Height"
        Case Is = "5"
            back2
            frmrho.Visible = True
            txtside1.Text = "4"
            txtcorner1.Text = "4"
            txtformula1.Text = "Length x Height"
    End Select
End Sub

Private Sub cmdcal_Click()
    calculate
End Sub

Private Sub cmdcir_Click()
    select_shape
    cmbshapes.ListIndex = "0"
End Sub

Private Sub cmdsqu_Click()
    select_shape
    cmbshapes.ListIndex = "1"
End Sub
Private Sub cmdrec_Click()
    select_shape
    cmbshapes.ListIndex = "2"
End Sub
Private Sub cmdtri_Click()
    select_shape
    cmbshapes.ListIndex = "3"
End Sub
Private Sub cmdtra_Click()
    select_shape
    cmbshapes.ListIndex = "4"
End Sub
Private Sub cmdrho_Click()
    select_shape
    cmbshapes.ListIndex = "5"
End Sub

Private Sub Command1_Click()
    frmtri.Visible = False
    back
End Sub

Private Sub Command2_Click()
    frmrec.Visible = False
    back
End Sub

Private Sub Command4_Click()
    frmtra.Visible = False
    back
End Sub

Private Sub Command5_Click()
    frmrho.Visible = False
    back
End Sub

Private Sub Command7_Click()
    frmsqu.Visible = False
    back
End Sub

Private Sub Command9_Click()
    frmcir.Visible = False
    back
End Sub

Private Sub Command3_Click()
    Form2.Hide
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
Private Sub Form_Load()
    Form2.Icon = Form1.Icon
    imgsqu.Picture = imgback.Picture
    imgrec.Picture = imgback.Picture
    imgtri.Picture = imgback.Picture
    imgrho.Picture = imgback.Picture
    imgtra.Picture = imgback.Picture
    imgcir.Picture = imgback.Picture
    img2d.Picture = imgback.Picture
    imgins.Picture = imgback.Picture
End Sub
Private Sub txtcir1_keypress(KeyAscii As Integer)
    If abc(KeyAscii) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtrec1_keypress(KeyAscii As Integer)
    If abc(KeyAscii) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtrec2_keypress(KeyAscii As Integer)
    If abc(KeyAscii) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtrho1_keypress(KeyAscii As Integer)
    If abc(KeyAscii) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtrho2_keypress(KeyAscii As Integer)
    If abc(KeyAscii) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtsqu1_keypress(KeyAscii As Integer)
    If abc(KeyAscii) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txttra1_keypress(KeyAscii As Integer)
    If abc(KeyAscii) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txttra2_keypress(KeyAscii As Integer)
    If abc(KeyAscii) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txttra3_keypress(KeyAscii As Integer)
    If abc(KeyAscii) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txttri1_keypress(KeyAscii As Integer)
    If abc(KeyAscii) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txttri2_keypress(KeyAscii As Integer)
    If abc(KeyAscii) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub
Private Sub calculate()
    Const Pi = 3.14
    If Not txtcir1.Text = "" Then
        txtcir2 = Format(Pi * (txtcir1.Text) ^ 2, "standard") & " units²"
    End If
    If Not txtsqu1.Text = "" Then
        txtsqu2 = Format((txtsqu1.Text) ^ 2, "#,###") & " units²"
    End If
    If Not txttri1.Text = "" And Not txttri2.Text = "" Then
        txttri3 = Format((txttri1.Text) * (txttri2.Text) / 2, "#,###.0#") & " units²"
    End If
    If Not txtrec1.Text = "" And Not txtrec2.Text = "" Then
        txtrec3 = Format((txtrec1.Text) * (txtrec2.Text), "#,###") & " units²"
    End If
    If Not txttra1.Text = "" And Not txttra2.Text = "" And Not txttra3.Text = "" Then
        txttra4 = Format((((txttra1.Text) + (txttra2.Text)) / 2) * txttra3.Text, "#,###.#") & " units²"
    End If
    If Not txtrho1.Text = "" And Not txtrho2.Text = "" Then
        txtrho3 = Format((txtrho1.Text) * (txtrho2.Text), "#,###") & " units²"
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
    frm2d.Visible = True
    frmins.Visible = True
    cmdcal.Visible = True
End Sub

Private Sub back()
    frmmain.Visible = True
    frm2d.Visible = False
    frmins.Visible = False
    cmdcal.Visible = False
    cmbshapes.ListIndex = "-1"
End Sub

Private Sub back2()
    frmtri.Visible = False
    frmrec.Visible = False
    frmtra.Visible = False
    frmrho.Visible = False
    frmsqu.Visible = False
    frmcir.Visible = False
End Sub
