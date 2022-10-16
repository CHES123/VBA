Attribute VB_Name = "Module1"

Sub delet_first_symbol()
For Each Cell In Selection
str1 = Cell
' Cell.Formula = "AAA"
Cell.Formula = Right(str1, Len(str1) - 1)
Next Cell
MsgBox "OK"
End Sub


Attribute VB_Name = "Module1"
Sub DeleteEmptyRows_NotID()
    SelectedRange = Selection.Rows.Count
    ActiveCell.Offset(0, 0).Select
    For i = 1 To SelectedRange
        If IsNumeric(ActiveCell.Value) And ActiveCell.Value > "" And ActiveCell.Value <> 0 Then
            ActiveCell.Offset(1, 0).Select
        
        Else
            Selection.EntireRow.Delete
        End If
                       
    Next i
End Sub
