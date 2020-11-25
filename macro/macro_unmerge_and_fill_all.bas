Attribute VB_Name = "Modulo1"
Sub UnMergeFill()

Dim cell As Range, joinedCells As Range

For Each cell In ThisWorkbook.ActiveSheet.UsedRange
    If cell.MergeCells Then
        Set joinedCells = cell.MergeArea
        cell.MergeCells = False
        joinedCells.Value = cell.Value
    End If
Next

End Sub
