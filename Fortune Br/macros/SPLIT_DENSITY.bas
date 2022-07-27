Attribute VB_Name = "SPLIT_DENSITY"
Sub SPLIT_DENSITY()

    Dim rowCount As Integer
    Dim arrayOfDoubleDensity As Variant
    Dim rS, rD As Integer
    rS = 2
    rD = 2

    Worksheets("MPU").Range("E:J").cellS.Clear

    rowCount = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row
    ActiveSheet.Range("E1").value = "FB/MM001"
    ActiveSheet.Range("I1").value = "FB/MM002"
    For i = 2 To rowCount
       
        If InStr(ActiveSheet.Range("A" & i), "4-MPU-D") = 1 Or InStr(ActiveSheet.Range("A" & i), "4-FU-D") = 1 Then
            ActiveSheet.Range("I" & rD).value = ActiveSheet.Range("A" & i)
            ActiveSheet.Range("J" & rD).value = ActiveSheet.Range("B" & i)
            rD = rD + 1
        Else:
            ActiveSheet.Range("E" & rS).value = ActiveSheet.Range("A" & i)
            ActiveSheet.Range("F" & rS).value = ActiveSheet.Range("B" & i)
            rS = rS + 1
        End If
    Next i
       
    'ADJUST COLUMN WIDTH
    Worksheets("MPU").Columns("E").ColumnWidth = 22
    Worksheets("MPU").Columns("I").ColumnWidth = 22

End Sub




