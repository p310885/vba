Attribute VB_Name = "DeleteAndShiftLeft"
' Deletes cells F8:F14 and I7 on every sheet whose name starts with "D",
' shifting the cells to the right of each deleted range to the left.
' SHEET_PASSWORD is used to unprotect / re-protect sheets before and after
' the operation when running against external workbooks.
Private Const SHEET_PASSWORD As String = "JeKaMaKla"

' ---------------------------------------------------------------------------
' Helper: perform the delete-and-shift-left operation on a single worksheet.
' The caller is responsible for unprotecting the sheet beforehand and
' re-protecting it afterwards when sheet protection is active.
' ---------------------------------------------------------------------------
Private Sub ProcessSheet(ws As Worksheet)
    ws.Range("F8:F14").Delete Shift:=xlShiftToLeft
    ws.Range("I7").Delete Shift:=xlShiftToLeft
End Sub

' ---------------------------------------------------------------------------
' Processes all "D"-prefixed sheets in the workbook that contains this macro.
' ---------------------------------------------------------------------------
Sub DeleteAndShiftLeft()
    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        If Left(ws.Name, 1) = "D" Then
            ProcessSheet ws
        End If
    Next ws
End Sub

' ---------------------------------------------------------------------------
' Opens every Excel file (*.xls*) in the same directory as this workbook,
' unprotects all "D"-prefixed sheets with SHEET_PASSWORD, runs the
' delete-and-shift-left operation, re-protects those sheets, then saves
' and closes each workbook.
' Files that cannot be opened and sheets that cannot be unprotected with
' SHEET_PASSWORD are silently skipped so that processing continues.
' ---------------------------------------------------------------------------
Sub DeleteAndShiftLeftAllFiles()
    Dim sDir      As String
    Dim sFile     As String
    Dim sFullPath As String
    Dim wb        As Workbook
    Dim ws        As Worksheet

    sDir = ThisWorkbook.Path & Application.PathSeparator

    ' Iterate over all Excel files in the directory.
    sFile = Dir(sDir & "*.xls*")
    Do While sFile <> ""
        sFullPath = sDir & sFile

        ' Skip the workbook that is running this macro.
        If LCase(sFullPath) <> LCase(ThisWorkbook.FullName) Then

            ' Attempt to open the workbook; skip it if it cannot be opened.
            On Error Resume Next
            Set wb = Workbooks.Open(Filename:=sFullPath, UpdateLinks:=False)
            On Error GoTo 0

            If Not wb Is Nothing Then
                For Each ws In wb.Worksheets
                    If Left(ws.Name, 1) = "D" Then
                        ' Attempt to unprotect; skip the sheet if the password
                        ' does not match or the sheet is not protected.
                        On Error Resume Next
                        ws.Unprotect Password:=SHEET_PASSWORD
                        On Error GoTo 0

                        ProcessSheet ws

                        ws.Protect Password:=SHEET_PASSWORD
                    End If
                Next ws

                wb.Close SaveChanges:=True
                Set wb = Nothing
            End If

        End If

        sFile = Dir()
    Loop

    MsgBox "Done: all Excel files in """ & sDir & """ have been processed.", vbInformation
End Sub
