VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Class Roll Wizard"
   ClientHeight    =   6825
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   7170
   LinkTopic       =   "Form2"
   ScaleHeight     =   6825
   ScaleWidth      =   7170
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRemove 
      Caption         =   "f"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   11
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "g"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   10
      Top             =   3240
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   960
      Width           =   1095
   End
   Begin VB.ListBox list_class 
      Height          =   3375
      ItemData        =   "newclasswizard.frx":0000
      Left            =   4200
      List            =   "newclasswizard.frx":0002
      OLEDropMode     =   1  'Manual
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   1920
      Width           =   2415
   End
   Begin VB.ListBox list_student 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   480
      OLEDropMode     =   1  'Manual
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   1920
      Width           =   2295
   End
   Begin VB.CommandButton CmdCancel 
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
      Height          =   495
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6120
      Width           =   1455
   End
   Begin VB.CommandButton CmdNext 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Finish"
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
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Image Image2 
      Height          =   1335
      Left            =   0
      Picture         =   "newclasswizard.frx":0004
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   1410
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Class Name:"
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
      Left            =   480
      TabIndex        =   9
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      Height          =   495
      Left            =   5160
      TabIndex        =   8
      Top             =   720
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Class List"
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
      Index           =   1
      Left            =   3960
      TabIndex        =   6
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Student List"
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
      Index           =   0
      Left            =   480
      TabIndex        =   5
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to Class Roll Wizard"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   240
      Width           =   7815
   End
   Begin VB.Image Image1 
      Height          =   6840
      Left            =   0
      Picture         =   "newclasswizard.frx":0498
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8760
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim i As Integer 'counter
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'needed for databse
    Dim wrkJet As DAO.Workspace
    Dim dbJet As DAO.Database
    Dim rsTeacher As DAO.Recordset
    Dim rsStudent As DAO.Recordset
    Dim rsmarks As DAO.Recordset

Private Sub cmdAdd_Click()
    If list_student.List(list_student.ListIndex) <> "" Then
    list_class.AddItem (list_student.List(list_student.ListIndex))
    list_student.RemoveItem (list_student.ListIndex)
    End If
End Sub

Private Sub CmdCancel_Click()
    Unload Me
    Marks.ClassTab.Tab = 1
End Sub

Private Sub CmdNext_Click()
    'saves the resulting new class
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'edit teacher
    rsTeacher.MoveFirst
    Do Until Label3.Caption = rsTeacher("username")
        rsTeacher.MoveNext
    Loop
    rsTeacher.Edit
    rsTeacher("no of classes") = rsTeacher("no of classes") + 1
    rsTeacher("class" & rsTeacher("no of classes")) = Text1.Text
    rsTeacher.Update
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'enter new roll
    rsmarks.MoveLast
    For i = 0 To list_class.ListCount - 1
        
        rsmarks.AddNew
        rsmarks("subject") = Text1.Text
        rsmarks("student_id") = Val(list_class.List(i))
        rsmarks.Update
    Next
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'open the roll
    Marks.Book(Marks.ClassTab.Tab).Visible = True
    Marks.Label2.Caption = Label3.Caption
    Unload Me
End Sub

Private Sub cmdRemove_Click()
    If list_class.List(list_class.ListIndex) <> "" Then
        list_student.AddItem (list_class.List(list_class.ListIndex))
        list_class.RemoveItem (list_class.ListIndex)
    End If
End Sub

Private Sub Form_Load()
    'opening db
    Set wrkJet = CreateWorkspace("", "admin", "", dbUseJet)
    Set dbJet = wrkJet.OpenDatabase(App.Path & "\studentdb.mdb")
    Set rsmarks = dbJet.OpenRecordset("mark", dbOpenDynaset)
    rsmarks.Sort = "subject, student_id"
    Set rsTeacher = dbJet.OpenRecordset("teacher")
    Set rsStudent = dbJet.OpenRecordset("student")
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    rsStudent.MoveFirst
    Do While Not rsStudent.EOF
        list_student.AddItem (Format(rsStudent("student_id"), "000") & " - " & rsStudent("first_name") & " " & rsStudent("last_name"))
        rsStudent.MoveNext
    Loop

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'needed for drag and drop - code from codeguru.net


Private Sub list_student_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    list_student.OLEDrag
End Sub

Private Sub list_student_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    ' Only allow moves
    AllowedEffects = vbDropEffectMove
    ' Assign the ListBox selection to the DataObject
    Data.SetData list_student
End Sub

Private Sub List_class_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim strList As String
    ' Check the format of the DataObject
    If Not Data.GetFormat(vbCFText) Then Exit Sub
    ' Retrieve the text from the DataObject
    strList = Data.GetData(vbCFText)
    ' If the item was not dropped on itself
    If Not strList = list_class.Text Then
        list_class.AddItem strList
        'Remove the item from the ListBox
        list_student.RemoveItem list_student.ListIndex
    End If
End Sub

Private Sub list_class_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    list_class.OLEDrag
End Sub

Private Sub list_class_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    ' Only allow moves
    AllowedEffects = vbDropEffectMove
    ' Assign the ListBox selection to the DataObject
    Data.SetData list_class
End Sub

Private Sub List_student_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim strList As String
    ' Check the format of the DataObject
    If Not Data.GetFormat(vbCFText) Then Exit Sub
    ' Retrieve the text from the DataObject
    strList = Data.GetData(vbCFText)
    ' If the item was not dropped on itself
    If Not strList = list_student.Text Then
        list_student.AddItem strList
        'Remove the item from the ListBox
        list_class.RemoveItem list_class.ListIndex
    End If
End Sub

