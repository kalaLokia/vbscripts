Private srcSheet, dstSheet, bomSheet As Worksheet
Private dst_row, dst_col As Integer
Private source, destination, bom As String
Private bomModel, bomComponent, bomValue As Range


Sub PLAN()
' Main Sub for program execution. 

Dim col_count, starter_cols, row_count, first_col, first_row, quantity  As Integer
Dim order, model As String

source = "BOOKING SHEET-GPU,GCUN,GDS,RCUN"
destination = "ALLOCATION"
bom = "BOM SHEET"

Set srcSheet = ThisWorkbook.Worksheets(source)
Set dstSheet = ThisWorkbook.Worksheets(destination)
Set bomSheet = ThisWorkbook.Worksheets(bom)

Set bomModel = bomSheet.Range("D1:SY1")
Set bomComponent = bomSheet.Range("B4:B1000")
Set bomValue = bomSheet.Range("D4:SY1000")

dst_row = 1
dst_col = 4

first_col = 6   'First row where process starts
first_row = 3   'First column where process starts

'Starting header columns count
starter_col = 5

'MO Number last column number. finds from the 2nd row of source sheet
col_count = srcSheet.Cells(2, srcSheet.Columns.Count).End(xlToLeft).Column - starter_cols

'Clear Destination range
dstSheet.Range(Cells(1, dst_col), Cells(3, dstSheet.Columns.Count - dst_col)).ClearContents

For col_no = first_col To col_count
    'Find last row with a value entered
    row_count = srcSheet.Cells(Rows.Count, col_no).End(xlUp).Row
    total_values = ValidCellsCount(srcSheet.Range(srcSheet.Cells(3, col_no), srcSheet.Cells(row_count, col_no)))
    row_no = 3

    For i = 0 To total_values
        If row_no > row_count Then
            Exit For
        End If

        If Not (IsEmpty(srcSheet.Cells(row_no, col_no))) Then
            order = srcSheet.Cells(2, col_no).Value
            model = srcSheet.Cells(row_no, 2).Value
            quantity = srcSheet.Cells(row_no, col_no).Value
            WriteToAllocation dst_col, order, model, quantity
            dst_col = dst_col + 1
        End If
        row_no = srcSheet.Cells(row_no, col_no).End(xlDown).Row
    Next i
Next col_no
GetBom      'Sub GetBom: To get the bom of each item

Msgbox("Success")

End Sub


Private Sub WriteToAllocation(col As Integer, order As Variant, model As Variant, qty As Integer)
    ' Write data to Destination Cell

    dstSheet.Cells(1, col).Value = order
    dstSheet.Cells(2, col).Value = model
    dstSheet.Cells(3, col).Value = qty
End Sub


Private Sub GetBom()
    ' Getting BOM for the items
 
    Dim col_num, row_num As Integer
    Dim temp_value As Integer

 
 
    dst_row = dstSheet.Cells(Rows.Count, 2).End(xlUp).Row
 
    For j = 4 To dst_row
        temp_value = dstSheet.Cells(j, 3)

        For i = 6 To (dst_col - 1)
            If Not (dstSheet.Cells(j, i).HasFormula) Then
                If Application.WorksheetFunction.CountIf(bomModel, Cells(2, i).Value) = 0 Then
                    dstSheet.Cells(j, i).Value = 0
                ElseIf Application.WorksheetFunction.CountIf(bomComponent, Cells(j, 2).Value) = 0 Then
                    dstSheet.Cells(j, i).Value = 0
                Else:
                    col_num = WorksheetFunction.Match(dstSheet.Cells(2, i), bomModel, 0)
                    row_num = WorksheetFunction.Match(dstSheet.Cells(j, 2), bomComponent, 0)
                    dstSheet.Cells(j, i).Value = WorksheetFunction.Index(bomValue, row_num, col_num).Value * dstSheet.Cells(3, i).Value
                End If
            End If

            If dstSheet.Cells(j, i).Value = 0 Then
                dstSheet.Cells(j, i).ClearContents
            ElseIf i = FirstColumn(dstSheet, j, i) Then
                dstSheet.Cells(j, i).FormatConditions.Delete
                dstSheet.Cells(j, i).FormatConditions.Add Type:=xlCellValue, Operator:=xlLessEqual, _Formula1:=temp_value
                dstSheet.Cells(j, i).FormatConditions(1).Interior.Color = vbRed
            Else:
                dstSheet.Cells(j, i).FormatConditions.Delete
                sum_value = temp_value - Application.WorksheetFunction.Sum(dstSheet.Range(Cells(j, 4), Cells(j, i)))
                dstSheet.Cells(j, i).FormatConditions.Add Type:=xlCellValue, Operator:=xlLessEqual, _Formula1:=sum_value
                dstSheet.Cells(j, i).FormatConditions(1).Interior.Color = vbBlue
            End If
 
        Next i
    Next j
 
End Sub


Private Function ValidCellsCount(rng As Range)
    ' Counts total cells with either a constant value or a formula.

    Count = 0
    On Error GoTo zero
    Count = rng.Cells.SpecialCells(xlCellTypeConstants).Count

    On Error GoTo zero
    Count = Count + rng.Cells.SpecialCells(xlCellTypeFormulas).Count

    zero:
        Count = Count + 0
        Resume Next

    ValidCellsCount = Count

End Function

 
Function FirstColumn(sht As Variant, r As Integer, c As Integer) As Integer
'   Finds the first column without any empty cell

    For x = 1 To 100
        If IsEmpty(sht.Cells(r, c)) Then
            c = c + 1
        Else:
            Exit For
        End If
    Next x
    FirstRow = c

End Function


' Material Planning Macro: https://m.box.com/view/800713326893  # private
' Made by kalaLokia
