VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    'Improve macro performance
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Worksheets("CATS_FILE").Cells.Clear
    
    DupForm.Show
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub
