Option Explicit

Sub Worksheet_Activate()
    Application.ScreenUpdating = False
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"
    ActiveWindow.zoom = 100
    ActiveWindow.DisplayHorizontalScrollBar = False
    ActiveWindow.DisplayVerticalScrollBar = False
    Application.DisplayFormulaBar = False
    ActiveWindow.DisplayHeadings = False
    Application.DisplayStatusBar = False
    ActiveWindow.DisplayWorkbookTabs = False
    Application.ScreenUpdating = True
End Sub

Sub Worksheet_Deactivate()
    Application.ScreenUpdating = False
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
    Application.WindowState = xlMaximized
    ActiveWindow.DisplayHorizontalScrollBar = True
    ActiveWindow.DisplayVerticalScrollBar = True
    Application.DisplayFormulaBar = True
    ActiveWindow.DisplayHeadings = True
    Application.DisplayStatusBar = True
    ActiveWindow.DisplayWorkbookTabs = True
    Application.ScreenUpdating = True
End Sub
