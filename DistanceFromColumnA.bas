'this script return the distance from Column A and put it on the status bar (the one on the buttom of the screen)
Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    On Error GoTo Clear
    '— Change this line if you want row distance instead:
    Dim dist As Long: dist = Target.Column
    Application.StatusBar = "Distance from Column A: " & dist
    Exit Sub
Clear:
    Application.StatusBar = False
End Sub

