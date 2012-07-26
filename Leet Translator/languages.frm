VERSION 5.00
Begin VB.Form Languages 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "A parent's primer to computer slang"
   ClientHeight    =   5520
   ClientLeft      =   3180
   ClientTop       =   2610
   ClientWidth     =   9600
   FillColor       =   &H00FFFFFF&
   Icon            =   "languages.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "languages.frx":08CA
   ScaleHeight     =   5520
   ScaleWidth      =   9600
   Begin VB.PictureBox leet 
      BackColor       =   &H80000007&
      Height          =   2175
      Left            =   6000
      Picture         =   "languages.frx":572F
      ScaleHeight     =   2115
      ScaleWidth      =   3465
      TabIndex        =   24
      Top             =   720
      Visible         =   0   'False
      Width           =   3525
   End
   Begin VB.PictureBox fra 
      Height          =   2175
      Left            =   6000
      Picture         =   "languages.frx":698E
      ScaleHeight     =   2115
      ScaleWidth      =   3435
      TabIndex        =   23
      Top             =   720
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.PictureBox ger 
      Height          =   2175
      Left            =   6000
      Picture         =   "languages.frx":50844
      ScaleHeight     =   2115
      ScaleWidth      =   3435
      TabIndex        =   22
      Top             =   720
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.PictureBox eng 
      Height          =   2175
      Left            =   6000
      Picture         =   "languages.frx":5124A
      ScaleHeight     =   2115
      ScaleWidth      =   3435
      TabIndex        =   21
      Top             =   720
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.OptionButton optleet 
      BackColor       =   &H00EAC34F&
      Caption         =   "Leet Speak"
      Height          =   375
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1920
      Width           =   1335
   End
   Begin VB.OptionButton optGer 
      BackColor       =   &H00EAC34F&
      Caption         =   "German"
      Height          =   375
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1440
      Width           =   1335
   End
   Begin VB.OptionButton optFra 
      BackColor       =   &H00E7BF41&
      Caption         =   "French"
      Height          =   375
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
      Width           =   1335
   End
   Begin VB.OptionButton optEng 
      BackColor       =   &H00E7BB38&
      Caption         =   "English"
      Height          =   375
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label lblcoo 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   2400
      TabIndex        =   20
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Cool"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   19
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label lblwar 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   2400
      TabIndex        =   18
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Illegally copied software"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   17
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label lblmad 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   2400
      TabIndex        =   16
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Mad Skills"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   15
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label lblexp 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   2400
      TabIndex        =   14
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Exploits"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   13
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblwoot 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   360
      TabIndex        =   12
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "We Own the Other Team"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   11
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Select Your Language"
      Height          =   255
      Left            =   4080
      TabIndex        =   10
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblhax 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   360
      TabIndex        =   9
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Hacks"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label lbllol 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   360
      TabIndex        =   7
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Laugh Out Loud"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label lblomg 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Oh My God"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Begin VB.Menu mnuexit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "Help"
      Begin VB.Menu mnuabout 
         Caption         =   "About"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "Languages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub leet_DblClick()
    MsgBox "Nice! You found my easter egg!"
    MsgBox "So now, please take the time to visit my blog"
    MsgBox "http://spaces.msn.com/members/wcexo"
    MsgBox "Thanks alot!"
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show
End Sub

Private Sub mnuexit_Click()
    iResponse% = MsgBox("Do you want to exit?", vbQuestion + vbYesNo, "Exit")
    If iResponse% = vbYes Then
        End
    End If
End Sub

Private Sub optEng_Click()
    If optEng.Value = True Then
        lblomg.Caption = "Oh my god"
        lbllol.Caption = "Laugh out loud"
        lblhax.Caption = "Hacks"
        lblwoot.Caption = "We Own The Other Team"
        lblexp.Caption = "Exploits"
        lblmad.Caption = "Mad skills"
        lblwar.Caption = "Illegally copied software"
        lblcoo.Caption = "Cool"
        optFra.Caption = "French"
        optEng.Caption = "English"
        optGer.Caption = "German"
        optleet.Caption = "Leet Speak"
        Languages.Caption = "A parent's primer to computer slang"
        menu
        eng.Visible = True
    End If
End Sub

Private Sub optFra_Click()
    If optFra.Value = True Then
        lblomg.Caption = "l'cOh mon dieu"
        lbllol.Caption = "rire dehors fort"
        lblhax.Caption = "entailles"
        lblwoot.Caption = "Nous possédons l'autre équipe"
        lblexp.Caption = "Exploits"
        lblmad.Caption = "qualifications folles"
        lblwar.Caption = "Logiciel illégalement copié"
        lblcoo.Caption = "frais"
        optFra.Caption = "Français"
        optEng.Caption = "Anglais"
        optGer.Caption = "Allemand"
        optleet.Caption = "Leet parlez"
        Languages.Caption = "L'amorce d'un parent à l'argot d'ordinateur"
        menu
        fra.Visible = True
    End If
End Sub

Private Sub optGer_Click()
    If optGer.Value = True Then
        lblomg.Caption = "OH mein Gott"
        lbllol.Caption = "Lachen heraus loud"
        lblhax.Caption = "Kerben"
        lblwoot.Caption = "Wir besitzen die andere Mannschaft"
        lblexp.Caption = "Großtaten"
        lblmad.Caption = "wütende Fähigkeiten"
        lblwar.Caption = "Illegal kopierte Software"
        lblcoo.Caption = "kühl"
        optFra.Caption = "Französisch"
        optEng.Caption = "Englisch"
        optGer.Caption = "Deutsch"
        optleet.Caption = "Leet sprechen"
        Languages.Caption = "Zündkapsel eines Elternteils zum Computerslang"
        menu
        ger.Visible = True
    End If
End Sub

Private Sub optleet_Click()
    If optleet.Value = True Then
        lblomg.Caption = "OMG"
        lbllol.Caption = "LOL"
        lblhax.Caption = "H4xz0r"
        lblwoot.Caption = "W00t"
        lblexp.Caption = "sploitz"
        lblmad.Caption = "m4d sk1llz"
        lblwar.Caption = "w4r3z or warez"
        lblcoo.Caption = "kewl"
        optFra.Caption = "Fr3nch"
        optEng.Caption = "3ngl15h"
        optGer.Caption = "G3rm4n"
        optleet.Caption = "1337 sp34k"
        Languages.Caption = "4 p4r3nt'5 pr1m3r t0 c0mput3r $14ng"
        menu
        leet.Visible = True
    End If
End Sub

Public Sub menu()
        fra.Visible = False
        eng.Visible = False
        ger.Visible = False
        leet.Visible = False
End Sub
