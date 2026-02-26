Attribute VB_Name = "DeleteAndShiftLeft"
' Deletes cells F8:F14 and I7 on every sheet whose name starts with "D",
' shifting the cells to the right of each deleted range to the left.
Sub DeleteAndShiftLeft()
    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        If Left(ws.Name, 1) = "D" Then
            ws.Range("F8:F14").Delete Shift:=xlShiftToLeft
            ws.Range("I7").Delete Shift:=xlShiftToLeft
        End If
    Next ws
End Sub
