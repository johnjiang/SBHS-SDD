VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Marks 
   Caption         =   "Mark Book"
   ClientHeight    =   8775
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   12870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "marks.frx":0000
   ScaleHeight     =   8775
   ScaleWidth      =   12870
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd_back 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   8040
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Statistical"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1440
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Editing"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   1
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1440
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.CommandButton cmdDel 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Delete Class"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   10920
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8040
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
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
      Height          =   555
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8040
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Sort By Rank"
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
      Left            =   10920
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Sort By Student Number"
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
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   360
      Width           =   1575
   End
   Begin TabDlg.SSTab ClassTab 
      Height          =   7935
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   13996
      _Version        =   393216
      Tabs            =   10
      Tab             =   1
      TabsPerRow      =   10
      TabHeight       =   520
      BackColor       =   14737632
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "marks.frx":20DB
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Image1(2)"
      Tab(0).Control(1)=   "Book(0)"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "marks.frx":20F7
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Image1(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Book(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "marks.frx":2113
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Image1(11)"
      Tab(2).Control(1)=   "Book(2)"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Tab 3"
      TabPicture(3)   =   "marks.frx":212F
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Book(3)"
      Tab(3).Control(1)=   "Image1(3)"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Tab 4"
      TabPicture(4)   =   "marks.frx":214B
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Image1(4)"
      Tab(4).Control(1)=   "Book(4)"
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "Tab 5"
      TabPicture(5)   =   "marks.frx":2167
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Image1(5)"
      Tab(5).Control(1)=   "Book(5)"
      Tab(5).ControlCount=   2
      TabCaption(6)   =   "Tab 6"
      TabPicture(6)   =   "marks.frx":2183
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Image1(6)"
      Tab(6).Control(1)=   "Book(6)"
      Tab(6).ControlCount=   2
      TabCaption(7)   =   "Tab 7"
      TabPicture(7)   =   "marks.frx":219F
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Image1(7)"
      Tab(7).Control(1)=   "Book(7)"
      Tab(7).ControlCount=   2
      TabCaption(8)   =   "Tab 8"
      TabPicture(8)   =   "marks.frx":21BB
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Image1(8)"
      Tab(8).Control(1)=   "Book(8)"
      Tab(8).ControlCount=   2
      TabCaption(9)   =   "Tab 9"
      TabPicture(9)   =   "marks.frx":21D7
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "Image1(9)"
      Tab(9).Control(1)=   "Book(9)"
      Tab(9).ControlCount=   2
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   6240
         TabIndex        =   5
         Top             =   -360
         Width           =   975
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Book 
         Height          =   6015
         Index           =   0
         Left            =   -74760
         TabIndex        =   1
         Top             =   960
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   10610
         _Version        =   393216
         Rows            =   8
         Cols            =   8
         FixedCols       =   2
         BackColorFixed  =   8421504
         ForeColorFixed  =   16777215
         BackColorBkg    =   12632256
         GridColor       =   8421504
         AllowUserResizing=   1
         RowSizingMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   8
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Book 
         Height          =   6015
         Index           =   1
         Left            =   240
         TabIndex        =   11
         Top             =   960
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   10610
         _Version        =   393216
         Rows            =   8
         Cols            =   8
         FixedCols       =   2
         BackColorFixed  =   8421504
         ForeColorFixed  =   16777215
         BackColorBkg    =   12632256
         GridColor       =   8421504
         AllowUserResizing=   1
         RowSizingMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   8
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Book 
         Height          =   6015
         Index           =   2
         Left            =   -74760
         TabIndex        =   12
         Top             =   960
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   10610
         _Version        =   393216
         Rows            =   8
         Cols            =   8
         FixedCols       =   2
         BackColorFixed  =   8421504
         ForeColorFixed  =   16777215
         BackColorBkg    =   12632256
         GridColor       =   8421504
         AllowUserResizing=   1
         RowSizingMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   8
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Book 
         Height          =   6015
         Index           =   3
         Left            =   -74760
         TabIndex        =   13
         Top             =   960
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   10610
         _Version        =   393216
         Rows            =   8
         Cols            =   8
         FixedCols       =   2
         BackColorFixed  =   8421504
         ForeColorFixed  =   16777215
         BackColorBkg    =   12632256
         GridColor       =   8421504
         AllowUserResizing=   1
         RowSizingMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   8
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Book 
         Height          =   6015
         Index           =   4
         Left            =   -74760
         TabIndex        =   14
         Top             =   960
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   10610
         _Version        =   393216
         Rows            =   8
         Cols            =   8
         FixedCols       =   2
         BackColorFixed  =   8421504
         ForeColorFixed  =   16777215
         BackColorBkg    =   12632256
         GridColor       =   8421504
         AllowUserResizing=   1
         RowSizingMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   8
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Book 
         Height          =   6015
         Index           =   5
         Left            =   -74760
         TabIndex        =   15
         Top             =   960
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   10610
         _Version        =   393216
         Rows            =   8
         Cols            =   8
         FixedCols       =   2
         BackColorFixed  =   8421504
         ForeColorFixed  =   16777215
         BackColorBkg    =   12632256
         GridColor       =   8421504
         AllowUserResizing=   1
         RowSizingMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   8
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Book 
         Height          =   6015
         Index           =   6
         Left            =   -74760
         TabIndex        =   16
         Top             =   960
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   10610
         _Version        =   393216
         Rows            =   8
         Cols            =   8
         FixedCols       =   2
         BackColorFixed  =   8421504
         ForeColorFixed  =   16777215
         BackColorBkg    =   12632256
         GridColor       =   8421504
         AllowUserResizing=   1
         RowSizingMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   8
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Book 
         Height          =   6015
         Index           =   7
         Left            =   -74760
         TabIndex        =   17
         Top             =   960
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   10610
         _Version        =   393216
         Rows            =   8
         Cols            =   8
         FixedCols       =   2
         BackColorFixed  =   8421504
         ForeColorFixed  =   16777215
         BackColorBkg    =   12632256
         GridColor       =   8421504
         AllowUserResizing=   1
         RowSizingMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   8
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Book 
         Height          =   6015
         Index           =   8
         Left            =   -74760
         TabIndex        =   18
         Top             =   960
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   10610
         _Version        =   393216
         Rows            =   8
         Cols            =   8
         FixedCols       =   2
         BackColorFixed  =   8421504
         ForeColorFixed  =   16777215
         BackColorBkg    =   12632256
         GridColor       =   8421504
         AllowUserResizing=   1
         RowSizingMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   8
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Book 
         Height          =   6015
         Index           =   9
         Left            =   -74760
         TabIndex        =   19
         Top             =   960
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   10610
         _Version        =   393216
         Rows            =   8
         Cols            =   8
         FixedCols       =   2
         BackColorFixed  =   8421504
         ForeColorFixed  =   16777215
         BackColorBkg    =   12632256
         GridColor       =   8421504
         AllowUserResizing=   1
         RowSizingMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   8
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   10065
         Index           =   11
         Left            =   -75000
         Picture         =   "marks.frx":21F3
         Stretch         =   -1  'True
         Top             =   360
         Width           =   12720
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   10065
         Index           =   9
         Left            =   -75000
         Picture         =   "marks.frx":42CE
         Stretch         =   -1  'True
         Top             =   360
         Width           =   12720
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   10065
         Index           =   8
         Left            =   -75000
         Picture         =   "marks.frx":63A9
         Stretch         =   -1  'True
         Top             =   360
         Width           =   12720
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   10065
         Index           =   7
         Left            =   -75000
         Picture         =   "marks.frx":8484
         Stretch         =   -1  'True
         Top             =   360
         Width           =   12720
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   10065
         Index           =   6
         Left            =   -75000
         Picture         =   "marks.frx":A55F
         Stretch         =   -1  'True
         Top             =   360
         Width           =   12720
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   10065
         Index           =   5
         Left            =   -75000
         Picture         =   "marks.frx":C63A
         Stretch         =   -1  'True
         Top             =   360
         Width           =   12720
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   10065
         Index           =   4
         Left            =   -75000
         Picture         =   "marks.frx":E715
         Stretch         =   -1  'True
         Top             =   360
         Width           =   12720
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   10065
         Index           =   3
         Left            =   -75000
         Picture         =   "marks.frx":107F0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   12720
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   10065
         Index           =   1
         Left            =   0
         Picture         =   "marks.frx":128CB
         Stretch         =   -1  'True
         Top             =   360
         Width           =   12720
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   10065
         Index           =   0
         Left            =   -75000
         Picture         =   "marks.frx":149A6
         Stretch         =   -1  'True
         Top             =   360
         Width           =   12720
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   10065
         Index           =   2
         Left            =   -75000
         Picture         =   "marks.frx":16A81
         Stretch         =   -1  'True
         Top             =   360
         Width           =   12720
      End
   End
   Begin VB.Image Image3 
      Height          =   615
      Left            =   120
      Picture         =   "marks.frx":18B5C
      Stretch         =   -1  'True
      Top             =   120
      Width           =   690
   End
   Begin VB.Image Image2 
      Height          =   10080
      Left            =   0
      Picture         =   "marks.frx":18FF0
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   12960
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   375
      Left            =   12960
      TabIndex        =   6
      Top             =   480
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   10065
      Index           =   10
      Left            =   0
      Picture         =   "marks.frx":1B0CB
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12960
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   88
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "Marks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim temp(30) As Integer 'long list of temps that can be used for holdding data
    Dim sd As Double 'standard deviation
    Dim lists(30) As Long 'used to sort ranks
    Dim markarray(40) As Single 'used to calculate means,sd,and highest and lowest mark
    Dim ranksort As Boolean 'true if the data is sorted by rank
    Dim i As Integer 'counter
    Dim stuname(10, 30) As String 'array used to hold each students name in each class, faster than database
    Dim Mean(30) As Integer 'array used for mean.
    Dim total As Integer 'the total
    Dim classnum As Integer 'the number of classes a teacher has.
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'declarations for database
    Dim wrkJet As DAO.Workspace
    Dim dbJet As DAO.Database
    Dim rsTeacher As DAO.Recordset
    Dim rsStudent As DAO.Recordset
    Dim rsmarks As DAO.Recordset

Private Sub Book_Click(Index As Integer) 'what happens when a result is clicked
    With Book(Index)
        If .Col <> 1 And .Col <> 0 And .Col <> 2 And .Col <> .Cols - 1 And .Row <> .Rows - 1 Then 'making sure the click is not part of the headers
            If Option1(1).Value = True Then 'if user can edit results
                Text1.height = .CellHeight - 100
                Text1.width = .CellWidth
                Text1.left = .CellLeft + 240
                Text1.top = .CellTop + 960
                Text1.Text = ""
            Else 'open sd box
                
                frmSplash.Visible = False
                frmSplash.Visible = True
                frmSplash.Slider1.Value = Val(.Text)
                temp(5) = .Col
                .Col = 1
                frmSplash.Label1 = .Text
                .Col = temp(5)
                .Row = .Rows - 1
                'making lines fit on screen.
                frmSplash.lnMean.Y1 = 4000 - 30 * (Val(.Text))
                frmSplash.lnMean.Y2 = 4000 - 30 * (Val(.Text))
                .Row = 1
                i = 0
                'making array to calculate the mean and sd.
                Do Until .Row = .Rows - 2
                    markarray(i) = Val(.Text)
                    i = i + 1
                    .Row = i + 1
                Loop
                'moving slider and applying results
                frmSplash.txtTop = biggest_number(i, markarray)
                frmSplash.Slider1.max = Val(frmSplash.txtTop)
                frmSplash.txtLow = smallest_number(i, markarray)
                frmSplash.Slider1.min = Val(frmSplash.txtLow)
                frmSplash.Slider1.Value = frmSplash.Slider1.max - (-frmSplash.Slider1.min + frmSplash.Slider1.Value)
                'standard deviation
                sd = StdDev(i, markarray)
                frmSplash.Shape1.top = frmSplash.lnMean.Y1 - 30 * sd
                frmSplash.Shape1.height = 60 * sd
            
            End If
        Else
            'make textbox inaccessible
            Text1.top = -1000
        End If
    End With
    'assigning temp(10) the books index for later.
    temp(10) = Index
End Sub

Private Sub ClassTab_Click(PreviousTab As Integer)
    If ClassTab.TabCaption(ClassTab.Tab) = "New Class" Then
        Book(ClassTab.Tab).Visible = False
        Form2.Visible = True
        Form2.Label3.Caption = Label1.Caption
    End If
End Sub

Private Sub cmd_back_Click()
    Unload Me
    main.Show
End Sub

Private Sub cmddel_Click()
    classnum = rsTeacher("No of classes")
    If classnum = 0 Then
        MsgBox "You have no classes to delete", vbInformation, "No Classes"
        Exit Sub
    End If
    If MsgBox("Are you sure you would like to delete this class?" & vbCrLf & "Note: All information for this class will be LOST!", vbInformation + vbYesNo, "Delete Class") = vbNo Then
        Exit Sub
    End If
    
    Dim a As Integer
    'edit teacher
    rsTeacher.MoveFirst
    Do Until Label1.Caption = rsTeacher("username")
        rsTeacher.MoveNext
    Loop
    rsTeacher.Edit
    rsTeacher("class" & (ClassTab.Tab + 1)) = ""
    a = ClassTab.Tab + 1
    Do While a < 10
        If rsTeacher("class" & (a + 1)) <> "" Then
            rsTeacher("class" & a) = rsTeacher("class" & (a + 1))
        End If
        a = a + 1
    Loop
    rsTeacher("no of classes") = rsTeacher("no of classes") - 1
    rsTeacher.Update
    'remove roll
    For i = 0 To Book(ClassTab.Tab).Rows - 3
        rsmarks.MoveFirst
        Do Until stuname(ClassTab.Tab, i) = rsmarks("student_id") And ClassTab.TabCaption(ClassTab.Tab) = rsmarks("subject")
            rsmarks.MoveNext
        Loop
    
        rsmarks.Delete
    
    Next
    classnum = classnum - 1
    Label1_Change
    ClassTab.Tab = 9
End Sub

Private Sub Command1_Click()
    'sorting the results in alphebetical order. classtab.tab is the current tab
    tableheaders (ClassTab.Tab)
    fillname (ClassTab.Tab)
    fillmarks (ClassTab.Tab)
    analyzemarks (ClassTab.Tab)
    ranksort = False
End Sub

Private Sub Command2_Click()
    'sorting by rank. using inbuilt search function.
    With Book(ClassTab.Tab)
        .Row = .Rows - 1
        .Col = .Cols - 1
        .Sort = 2
    End With
    ranksort = True
End Sub

Public Sub tableheaders(tablenumber As Integer)
    'fill in the table column headers
    With Book(tablenumber)
        .Cols = 8
        .Row = 0
        .Col = 0
        .Text = "Student Number"
        .Col = 1
        .Text = "Student Name"
        .Col = 2
        .Text = "Rank"
        .Col = .Cols - 1
        .Text = "Total"
        'changing width to fit the data.
        For temp(2) = 0 To .Cols - 2
            .ColWidth(temp(2)) = 1300
        Next
        ' counting the number of assesments being tested.
        For temp(2) = 1 To .Cols - 4
            .Col = temp(2) + 2
            .ColWidth(temp(2) + 2) = 1200
            .Text = "Assessment " & temp(2)
        Next

    End With
End Sub

Public Sub fillname(tablenumber As Integer)
    'filling in the student number and respective name
    'using 2 databases and collaborating data
    'our own function, stuname
    i = 0
    rsmarks.MoveFirst
    Do Until rsmarks.EOF
        If ClassTab.TabCaption(tablenumber) = rsmarks("subject") Then
            temp(4) = temp(4) + 1
            stuname(tablenumber, i) = rsmarks("student_id")
            i = i + 1
        End If
        rsmarks.MoveNext
    Loop
    i = 0
    rsmarks.MoveFirst
    With Book(tablenumber)
        .Rows = temp(4) + 2
        .Col = 1
        For temp(2) = 1 To .Rows - 2
        .Row = temp(2)
        .Text = studentname(Val(stuname(tablenumber, i)))
        i = i + 1
        Next
        .Col = 0
        .Row = .Rows - 1
        'last row says average
        .Text = "Average:"
        .Col = 2
        .Text = "N/A"
        .Col = 0
        .Row = 0
        i = 0
        For temp(2) = 1 To .Rows - 2
        .Row = temp(2)
        .Text = stuname(tablenumber, i)
        i = i + 1
        Next
    End With
    temp(4) = 0
End Sub

Public Sub fillmarks(tablenumber As Integer)
    On Error GoTo errorhandle2 'using error handler to remove divide by 0 error that may occur calculating averages.
    'fill in the table with marks
    i = 0
    With Book(tablenumber)
        Do Until stuname(tablenumber, i) = ""
        rsStudent.MoveFirst
        Do Until stuname(tablenumber, i) = rsmarks("student_id") And ClassTab.TabCaption(tablenumber) = rsmarks("subject")
            rsmarks.MoveNext
        Loop
        .Col = 0
        .Row = 0
        Do Until .Text = stuname(tablenumber, i)
            .Row = .Row + 1
        Loop
        For temp(2) = 3 To .Cols - 2
            .Col = temp(2)
            .Text = Val(rsmarks("assessment " & (.Col - 2)))
        Next
        'increment
        i = i + 1
        Loop
    End With
errorhandle2:
    Exit Sub
End Sub

Public Sub analyzemarks(tablenumber As Integer)
    On Error GoTo errorhandler
    With Book(tablenumber)
        'mean
        'finds the mean of each column of results/each task.
        .Col = 3
        Do Until .Col = .Cols - 1
        .Row = 1
        i = 1
        Do Until .Row = .Rows - 1
            Mean(i) = Mean(i - 1) + Val(.Text)
            i = i + 1
            .Row = i
        Loop
        .Text = FormatNumber((Mean(i - 1) / (i - 1)), 2)
        .Col = .Col + 1
        Loop
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'totals
        'finds the totals of each student - hence used to calculate ranks.
        .Row = 1
        Do Until .Row = .Rows - 1
            .Col = 3
            total = 0
            Do Until .Col = .Cols - 1
                total = total + Val(.Text)
                .Col = .Col + 1
            Loop
            .Col = .Cols - 1
            .Text = total
            .Row = .Row + 1
        Loop
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'rank
        'works out the ranks of each student by their totals so far. It assumes that all weighting is equal.
        .Row = 1
        .Col = 7
        i = 1
        Do Until .Row = .Rows - 1
            lists(i) = Val(.Text)
            i = i + 1
            .Row = i
        Loop
        'uses a selection sort and the array lists()
        Selectionsort lists(), 1, i - 1
        .Row = 1
        Do Until .Row = .Rows - 1
        .Col = 7
        i = 1
        Do Until lists(i) = .Text
            i = i + 1
        Loop
        .Col = 2
        .Text = .Rows - 1 - i
        .Row = .Row + 1
        Loop
    End With
    'this just makes sure that no error occurs.
errorhandler:
    Exit Sub
End Sub

Public Function studentname(studentnumber As Integer) As String
    'get student name from number
    rsStudent.MoveFirst
    Do Until studentnumber = rsStudent("student_id")
        rsStudent.MoveNext
    Loop
    studentname = rsStudent("first_name") & " " & rsStudent("last_name")
End Function

Public Sub Selectionsort(List() As Long, min As Integer, max As Integer) 'selection sort
    Dim i As Integer
    Dim j As Integer
    Dim best_value As Long
    Dim best_j As Integer
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    For i = min To max - 1
        best_value = List(i)
        best_j = i
        For j = i + 1 To max
            If List(j) < best_value Then
                best_value = List(j)
                best_j = j
            End If
        Next j
        List(best_j) = List(i)
        List(i) = best_value
    Next i
End Sub

Private Sub Command3_Click()
Dim listname As String
    listname = ClassTab.Caption
    P001.GridToPrint = Book(ClassTab.Tab)
    P001.ReportTitle = "Class: " + listname
    P001.AllowDialogue = True   'let the user change some of the settings at run time
    P001.TitlePages = TitleOption.USER_MAY_SET + TitleOption.TITLE_ALL_PAGES
    P001.SetEffects = Effects.EFFECTS_NORMAL
    P001.PrinterOrientation = PGLandscape
    P001.TopMargin = 20
    P001.LeftMargin = 20
    'P001.TopMargin = 10
    'P001.optPPS(1) = True
    'plus any more property settings you want to control from the calling program
    'then just
    P001.PrintGridAPI   'call the PrintGrid routines to display the P001 dialogue
End Sub

Private Sub Form_GotFocus()
    frmSplash.Visible = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    main.Show
End Sub

Private Sub Label1_Change()
    On Error GoTo HERE
        'opening db
    Set wrkJet = CreateWorkspace("", "admin", "", dbUseJet)
    Set dbJet = wrkJet.OpenDatabase(App.Path & "\studentdb.mdb")
    Set rsmarks = dbJet.OpenRecordset("mark", dbOpenDynaset)
    rsmarks.Sort = "subject, student_id"
    Set rsTeacher = dbJet.OpenRecordset("teacher")
    Set rsStudent = dbJet.OpenRecordset("student")
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Do Until Label1.Caption = rsTeacher("username")
        rsTeacher.MoveNext
    Loop
    classnum = Val(rsTeacher("No of Classes")) - 1
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'enter data
    For temp(1) = 0 To 9
        If temp(1) > classnum Then
            ClassTab.TabCaption(temp(1)) = ""
            ClassTab.TabEnabled(temp(1)) = False
        Else
            ClassTab.TabCaption(temp(1)) = rsTeacher("class" & (temp(1) + 1))
            tableheaders (temp(1))
            fillname (temp(1))
            fillmarks (temp(1))
            analyzemarks (temp(1))
        End If
    Next
    ClassTab.TabCaption(classnum + 1) = "New Class"
    ClassTab.TabEnabled(classnum + 1) = True
    
HERE:
    If classnum < 0 Then
        ClassTab.TabCaption(0) = "New Class"
        For temp(1) = 1 To 9
            ClassTab.TabCaption(temp(1)) = ""
            ClassTab.TabEnabled(temp(1)) = False
        Next
        Book(ClassTab.Tab).Clear
        Book(ClassTab.Tab).Visible = False
    End If
End Sub

Private Sub Label2_Change()
    Label1_Change
    Label2.Caption = ""
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    'when user enters data.
    If Val(Text1.Text) < 101 And Val(Text1.Text) > -1 Then
    If KeyCode = vbKeyReturn Then
        Book(temp(10)).Text = Val(Text1.Text)
        saveresult (temp(10))
        Text1.top = -500
        analyzemarks (temp(10))
    End If
    End If
End Sub

Sub saveresult(tablenumber As Integer)
    'saves result by finding the student in the database and his respective class and period.
    With Book(tablenumber)
        rsmarks.MoveFirst
        temp(5) = .Col
        .Col = 0
        Do Until .Text = rsmarks("student_id") And ClassTab.TabCaption(tablenumber) = rsmarks("subject")
            rsmarks.MoveNext
        Loop
        .Col = temp(5)
        .Row = 0
        rsmarks.Edit
        rsmarks(.Text) = Val(Text1.Text)
        rsmarks.Update
    End With
End Sub
