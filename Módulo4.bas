Attribute VB_Name = "Módulo4"
Sub ApagarTextosENumeros()
    Dim rng As Range
    Set rng = Range("A4", Cells(Rows.Count, Columns.Count))
    rng.SpecialCells(xlCellTypeConstants).ClearContents
End Sub

