Attribute VB_Name = "BOM_DB_HEIRARCHY"
Sub LOOPPP()
    
    Dim r As Integer
    
    For i = 2 To 8452
        If Not Worksheets("Sheet2").Range("A" & i - 1) = Worksheets("Sheet2").Range("A" & i) Then
            r = 1
            Worksheets("Sheet2").Range("G" & i).Value = Worksheets("Sheet2").Range("B" & i) & r
        Else:
            r = r + 1
            Worksheets("Sheet2").Range("G" & i).Value = Worksheets("Sheet2").Range("B" & i) & r
        End If
    Next i
End Sub

Sub LOOPPPPP()
    
    Dim r As Integer
    Dim c As Integer
    r = 1
    c = 12
    For i = 1 To 1568
        If Not Worksheets("Sheet3").Range("K" & r) = Worksheets("Sheet3").Range("A" & i) Then
            r = r + 1
            c = 12
            Worksheets("Sheet3").Cells(r, c).Value = Worksheets("Sheet3").Range("B" & i)
        Else:
            c = c + 1
            Worksheets("Sheet3").Cells(r, c).Value = Worksheets("Sheet3").Range("B" & i)
        End If
    Next i

End Sub
