Attribute VB_Name = "saveToCsv"
'A sample macro used to save a range to csv file format
Sub SAVE_DATA_TO_CSV()

    Dim filename As String
    Dim curWb As Workbook
    Dim tempWb As Workbook
    Dim rngToSave As Range

    Application.DisplayAlerts = False
    On Error GoTo err

    Set curWb = ThisWorkbook
    filename = curWb.Path & "\" & "Sample-" & VBA.Format(VBA.Now, "dd-MM-yyyy") & ".csv"
    RowCount = Worksheets("Grand Final").Range("C" & Rows.Count).End(xlUp).Row
    Set rngToSave = Worksheets("Grand Final").Range("B1:F" & RowCount)
    rngToSave.Copy

    Set tempWb = Application.Workbooks.Add(1)
    With tempWb
        .Sheets(1).Range("A1").PasteSpecial xlPasteValues
        .SaveAs filename:=filename, FileFormat:=xlCSV, CreateBackup:=False
        .Close
    End With
err:
    Application.DisplayAlerts = True
End Sub
