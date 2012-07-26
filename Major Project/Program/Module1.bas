Attribute VB_Name = "validate"
Option Explicit
Dim permission As String
'open web browser
Declare Function ShellExecute Lib "shell32.dll" _
Alias "ShellExecuteA" ( _
ByVal hwnd As Long, _
ByVal lpOperation As String, _
ByVal lpFile As String, ByVal lpParameters As String, _
ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Sub check_status()

    stu_exp.stat.Panels(1).Text = stu_exp.list_name.ListCount & " students found"
    stu_exp.stat.Panels(2).Text = "Signed in at " & Time
    stu_exp.stat.Panels(3).Text = login.db_login.Recordset.Fields(0) & " has logged in as " & permission
    
End Sub

Public Sub check_permission()
    
    If login.db_login.Recordset.Fields("Admin") = True Then
            
        permission = "Admin"
        
        stu_exp.Toolbar.Buttons.Item(2).Enabled = True
        stu_exp.Toolbar.Buttons.Item(3).Enabled = True
        stu_exp.Toolbar.Buttons.Item(4).Enabled = True
        
    Else
        permission = "Teacher"
    End If
    
End Sub

Public Sub RunBrowser(strURL As String, iWindowStyle As Integer, fH As Long)
    Dim lSuccess As Long
    '-- Shell to default browser
    lSuccess = ShellExecute(fH, "Open", strURL, 0&, 0&, iWindowStyle)
End Sub



