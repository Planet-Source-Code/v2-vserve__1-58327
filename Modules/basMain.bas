Attribute VB_Name = "basMain"
Option Explicit


Sub Main()

    If App.PrevInstance = True Then End
    App.TaskVisible = False
    ReDim ArrFilterType(1) As String
    ArrFilterType(0) = "Protocol"
    ArrFilterType(1) = "Script"
    
    ScanFilters
    frmMain.TabStrip.Tabs(1).Selected = True
    frmMain.Show
End Sub



