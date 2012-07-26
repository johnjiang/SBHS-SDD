Attribute VB_Name = "grids"
Option Explicit

Public Sub set_grid()
    view.dvddb1.RecordSource = "Select * FROM DVDs order by Title"
    Set view.txtimage.DataSource = view.dvddb1
    view.txtimage.DataField = "Cover"
    
    'sets the grid in the viewing window
    
    Set view.dvdgrid1.DataSource = view.dvddb1
    
    ' Sets dvdgrid width and column widths
    view.dvdgrid1.ColWidth(0) = view.dvdgrid1.Width * 2

End Sub

Public Sub Set_datasource()
    ' refresh the data source and rebind it to the flexgrid
    view.dvddb1.Refresh
    
    Set view.dvdgrid1.DataSource = view.dvddb1
    Set addentry.dvdgrid1.DataSource = view.dvddb1
    
End Sub

Public Sub rebind_log()
    
    'refreshes and rebind grid to db
    'login.dblog.Refresh
    Set log.dvdlog.DataSource = login.dblog
    
End Sub


Public Sub rebind_grid()
    
    'refreshes and rebind grid to db
    login.dblogin.Refresh
    Set users.usergrid.DataSource = login.dblogin
    
End Sub
