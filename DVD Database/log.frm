VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form log 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "log"
   ClientHeight    =   6480
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   9675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   9675
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dvdlog 
      Height          =   6495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   11456
      _Version        =   393216
      Rows            =   3
      Cols            =   4
      FixedRows       =   0
      FixedCols       =   0
      BackColorFixed  =   15984078
      BackColorSel    =   14398577
      BackColorBkg    =   16777215
      BackColorUnpopulated=   16777215
      GridColorFixed  =   0
      GridColorUnpopulated=   0
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      SelectionMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
      _Band(0).ColHeader=   1
      _Band(0)._NumMapCols=   5
      _Band(0)._MapCol(0)._Name=   "Time"
      _Band(0)._MapCol(0)._RSIndex=   0
      _Band(0)._MapCol(1)._Name=   "Date"
      _Band(0)._MapCol(1)._RSIndex=   1
      _Band(0)._MapCol(2)._Name=   "Event"
      _Band(0)._MapCol(2)._RSIndex=   2
      _Band(0)._MapCol(3)._Name=   "User"
      _Band(0)._MapCol(3)._RSIndex=   3
      _Band(0)._MapCol(4)._Name=   "ID"
      _Band(0)._MapCol(4)._RSIndex=   4
      _Band(0)._MapCol(4)._Alignment=   7
      _Band(0)._MapCol(4)._Hidden=   -1  'True
   End
   Begin VB.Menu mnu_context 
      Caption         =   "Context"
      Visible         =   0   'False
      Begin VB.Menu mnu_copy 
         Caption         =   "Copy"
      End
   End
End
Attribute VB_Name = "log"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    
    'binds grid to db
    Set dvdlog.DataSource = login.dblog
    
    'sets the width of individual columns
    dvdlog.ColWidth(0) = dvdlog.Width * 0.1
    dvdlog.ColWidth(1) = dvdlog.Width * 0.1
    dvdlog.ColWidth(2) = dvdlog.Width * 0.51
    dvdlog.ColWidth(3) = dvdlog.Width * 0.25
End Sub

Private Sub Form_Unload(Cancel As Integer)
    view.Enabled = True
End Sub

Private Sub dvdlog_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'displays context menu if right click is pressed
    If Button = vbRightButton Then
        PopupMenu mnu_context, vbPopupMenuRightButton
    End If
End Sub

Private Sub mnu_copy_Click()
    
    'copies the data clicked into clipboard
    Clipboard.Clear
    Clipboard.SetText dvdlog.TextMatrix(dvdlog.MouseRow, 0) & " " & dvdlog.TextMatrix(dvdlog.MouseRow, 1) & " " & dvdlog.TextMatrix(dvdlog.MouseRow, 2) & " by " & dvdlog.TextMatrix(dvdlog.MouseRow, 3)
End Sub
