Sub Group12ColumnsFromActiveCell()
    Dim startCol As Long
    Dim startRow As Long
    Dim rng As Range

    startCol = ActiveCell.Column
    startRow = ActiveCell.row

    Set rng = Range(Cells(startRow, startCol), Cells(startRow, startCol + 11))

    rng.Columns.Group
End Sub

