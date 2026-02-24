Attribute VB_Name = "UnprotectCells"
' Unprotects cells A23:F1500 permanently on all sheets except
' "Metadata", "Parameters", and "Übersicht".
' The sheet protection password is defined in SHEET_PASSWORD below.
Private Const SHEET_PASSWORD As String = "JeKaMaKla"

Sub UnprotectCellsOnAllSheets()
    Dim ws As Worksheet
    Dim excludedSheets As Variant
    Dim sheetName As String
    Dim isExcluded As Boolean
    Dim i As Long

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
            ws.Unprotect Password:=SHEET_PASSWORD
            ws.Range("A23:F1500").Locked = False
            ws.Protect Password:=SHEET_PASSWORD
        End If
    Next ws

    MsgBox "Done: cells A23:F1500 are permanently unlocked on all applicable sheets.", vbInformation
End Sub
