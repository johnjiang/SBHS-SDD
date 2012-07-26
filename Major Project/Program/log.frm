VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form log 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Action Log"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   7935
   StartUpPosition =   3  'Windows Default
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Actionlog 
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   8281
      _Version        =   393216
      FixedCols       =   0
      BackColorFixed  =   8421504
      ForeColorFixed  =   16777215
      BackColorBkg    =   12632256
      GridColor       =   8421504
      AllowUserResizing=   1
      RowSizingMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Image Image1 
      Height          =   6135
      Left            =   0
      Picture         =   "log.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8055
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
Private Sub actionlog_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'displays context menu if right click is pressed
    If Button = vbRightButton Then
        PopupMenu mnu_context, vbPopupMenuRightButton
    End If
End Sub

Private Sub Form_Load()
    Set Actionlog.DataSource = login.dblog
    
    'sets the width of individual columns
    Actionlog.ColWidth(0) = Actionlog.width * 0.05
    Actionlog.ColWidth(1) = Actionlog.width * 0.12
    Actionlog.ColWidth(2) = Actionlog.width * 0.15
    Actionlog.ColWidth(3) = Actionlog.width * 0.44
    Actionlog.ColWidth(4) = Actionlog.width * 0.19
End Sub

Private Sub Form_Unload(Cancel As Integer)
    main.Show
End Sub

Private Sub mnu_copy_Click()
    
    'copies the data clicked into clipboard
    Clipboard.Clear
    Clipboard.SetText Actionlog.TextMatrix(Actionlog.MouseRow, 1) & "   " & Actionlog.TextMatrix(Actionlog.MouseRow, 2) & "     " & Actionlog.TextMatrix(Actionlog.MouseRow, 3) & " by " & Actionlog.TextMatrix(Actionlog.MouseRow, 4)
End Sub


