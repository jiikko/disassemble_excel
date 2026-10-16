Private Sub Worksheet_BeforeDoubleClick( _
    ByVal Target As Range, Cancel As Boolean)
    Sheet1.Cells(1, 3) = Sheet1.Cells(1, 1) + Sheet1.Cells(1, 2)
End Sub

Private Sub Worksheet_SelectionChange( _
    ByVal Target As Range)
End Sub
