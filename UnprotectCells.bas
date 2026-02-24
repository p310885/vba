Attribute VB_Name = "UnprotectCells"
' Unprotects cells A23:F1500 permanently on all sheets except Metadata, Parameters and Übersicht.
' "Permanently unprotect" means setting the Locked property to False for those cells,
' so they remain editable even when the sheet protection is active.
Sub UnprotectCellsOnAllSheets()
    Dim ws As Worksheet
    Dim excludedSheets As Variant
    Dim sheetName As String
    Dim isExcluded As Boolean
    Dim i As Integer
    Const PASSWORD As String = "JeKaMaKla"

    excludedSheets = Array("Metadata", "Parameters", "Übersicht")

    For Each ws In ThisWorkbook.Worksheets
        sheetName = ws.Name
        isExcluded = False

        For i = LBound(excludedSheets) To UBound(excludedSheets)
            If sheetName = excludedSheets(i) Then
                isExcluded = True
                Exit For
            End If
        Next i

        If Not isExcluded Then
            On Error Resume Next
            ws.Unprotect Password:=PASSWORD
            On Error GoTo 0
            ws.Range("A23:F1500").Locked = False
            ws.Protect Password:=PASSWORD
        End If
    Next ws

    MsgBox "Done: Cells A23:F1500 are permanently unprotected on all relevant sheets.", vbInformation
End Sub
