
Public Sub separate_line_break()
    target_col = "B"     'Define the column you want to break
    ColLastRow = Range(target_col & Rows.Count).End(xlUp).Row
    Application.ScreenUpdating = False
    For Each Rng In Range(target_col & "1" & ":" & target_col & ColLastRow)
        If InStr(Rng.Value, vbLf) Then
            Rng.EntireRow.Copy
            Rng.EntireRow.Insert
            Rng.Offset(-1, 0) = Mid(Rng.Value, 1, InStr(Rng.Value, vbLf) - 1)
            Rng.Value = Mid(Rng.Value, Len(Rng.Offset(-1, 0).Value) + 2, Len(Rng.Value))
        End If
    Next
    
    ColLastRow2 = Range(target_col & Rows.Count).End(xlUp).Row
    For Each Rng2 In Range(target_col & "1" & ":" & target_col & ColLastRow2)
        If Len(Rng2) = 0 Then
            Rng2.EntireRow.Delete
        End If
    Next
    Application.ScreenUpdating = True
End Sub
