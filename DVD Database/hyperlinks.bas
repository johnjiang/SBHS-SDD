Attribute VB_Name = "hyperlinks"
Option Explicit

'open web browser
Declare Function ShellExecute Lib "shell32.dll" _
Alias "ShellExecuteA" ( _
ByVal hwnd As Long, _
ByVal lpOperation As String, _
ByVal lpFile As String, ByVal lpParameters As String, _
ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Sub RunBrowser(strURL As String, iWindowStyle As Integer, fH As Long)
    Dim lSuccess As Long
    '-- Shell to default browser
    lSuccess = ShellExecute(fH, "Open", strURL, 0&, 0&, iWindowStyle)
End Sub
