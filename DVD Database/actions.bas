Attribute VB_Name = "actions"
Public Username As String
Option Explicit

Public Sub dvdcount()
    'counts the number of records in the database and displays it in the status bar
    view.dvdstatus.Panels(3).Text = "Total DVDs: " & view.dvddb1.Recordset.RecordCount
End Sub

Public Sub load_imageview()
    If view.txtimage = "" Then
        view.imgcover = LoadPicture(App.Path & "\Images\noimage.gif")
        addentry.imgeditcover = LoadPicture(App.Path & "\Images\noimage.gif")
    Else
        view.imgcover = LoadPicture(App.Path & view.txtimage.Text)
        addentry.imgeditcover = LoadPicture(App.Path & view.txtimage.Text)
    End If
End Sub

Public Sub load_imageedit()
    If view.txtimage = "" Then
        addentry.imgeditcover = LoadPicture(App.Path & "\Images\noimage.gif")
    Else
        addentry.imgeditcover = LoadPicture(App.Path & view.txtimage.Text)
    End If
End Sub
Public Sub check_username()
    'displays username in statusbar
    
    view.dvdstatus.Panels(1).Text = "Logged in as " & Username
End Sub
