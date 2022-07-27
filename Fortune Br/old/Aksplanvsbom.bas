Attribute VB_Name = "Module1"
Function myeqn(val As Range, week As Integer)
Dim result As Variant
'MsgBox (val.Value)

'result = WorksheetFunction.VLookup(Worksheets("PLAN").Range("A" & i), Worksheets("PLAN").Range("A:B"), 2, 0) * WorksheetFunction.VLookup(Worksheets("PLAN").Range("A3"), Worksheets("PLAN").Range("A:Z"), week, 0) / WorksheetFunction.HLookup(Worksheets("PLAN").Range("A3"), Worksheets("BOM").Range("D:PV"), WorksheetFunction.Match(val, Worksheets("BOM").Range("A:A"), 0), 0)

For i = 3 To 500
    If WorksheetFunction.HLookup(Worksheets("PLAN").Range("A" & i), Worksheets("BOM").Range("D:PV"), WorksheetFunction.Match(val, Worksheets("BOM").Range("A:A"), 0), 0) <> 0 Then
        result = result + WorksheetFunction.VLookup(Worksheets("PLAN").Range("A" & i), Worksheets("PLAN").Range("A:B"), 2, 0) * WorksheetFunction.VLookup(Worksheets("PLAN").Range("A" & i), Worksheets("PLAN").Range("A:Z"), week, 0) / WorksheetFunction.HLookup(Worksheets("PLAN").Range("A" & i), Worksheets("BOM").Range("D:PV"), WorksheetFunction.Match(val, Worksheets("BOM").Range("A:A"), 0), 0)
    End If
Next i

myeqn = result
End Function
