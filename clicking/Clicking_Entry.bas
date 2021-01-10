Sub CLICK_ENTRY()
    Dim ccu_i As Integer
    Dim cci_i As Integer
    Dim ccu_n As Integer
    Dim cci_n As Integer
    Dim ccProcess As String
    Dim n As Integer

    Dim oClick As New clsClicking
    Worksheets("datas").Cells.Clear    'Clearing all data on the "datas" sheet before writing out
    
    With Worksheets("CLICKING")
    'Checking Insole, Upper entries
        On Error Resume Next
            cci_i = Application.WorksheetFunction.Match("INSOLE", .Range("B:B"), 0)
        On Error GoTo 0
        
        If cci_i <> 0 Then
            cci_n = Worksheets("CLICKING").Range("B" & cci_i).MergeArea.Rows.Count
            'MsgBox cci_n
        Else
            MsgBox "INSOLE not found"
        End If
        
        On Error Resume Next
            ccu_i = Application.WorksheetFunction.Match("UPPER", .Range("B:B"), 0)
        On Error GoTo 0
        
        If ccu_i <> 0 Then
            'MsgBox "Found at " & .Cells(i, 2).Address(0, 0)
            ccu_n = Worksheets("CLICKING").Range("B" & ccu_i).MergeArea.Rows.Count
        Else
            MsgBox "UPPER not found"
        End If
           'PULLED NO.OF CELLS AND INITIAL CELL REF FOR CCP AND CCS FROM CLICKING SHEET
    End With

    oClick.WriteHeaders
    n = 3

    For i = 1 To cci_n
        If WorksheetFunction.SUM(Range("G" & i & ":S" & i)) > 0 Then
            oClick.Process = "CCP"
            oClick.Artno = Worksheets("CLICKING").Range("D" & cci_i).Value
            oClick.Color = Worksheets("CLICKING").Range("E" & cci_i).Value
            oClick.Category = Worksheets("CLICKING").Range("F" & cci_i).Value
            
            oClick.Jobno = Worksheets("CLICKING").Range("C" & cci_i).Value
            oClick.Plan = Worksheets("CLICKING").Range("T" & cci_i).Value
            
            For j = 1 To 13
                If IsEmpty(Worksheets("CLICKING").Cells(cci_i, j + 6).Value) = False Or Worksheets("CLICKING").Cells(cci_i, j + 6).Value <> 0 Then
                    oClick.WriteToSheet rowNo:=n, colNo:=cci_i, sSize:=j
                    n = n + 1
                End If
            Next j
            cci_i = cci_i + 1
        End If
    Next i

    ' Page Break
     n = n + 2

    Dim pCCP1, pCCP1_common, pCCF, pNoCCS As Variant

    'Common CCP1
    pxCCP1 = Split("3290,3780,3059", ",")
    'Common CCS
    pxCCS = Split("3391B", ",")
    'CCP1
    pCCP1 = Split("3791,D4003,8180", ",")
    'CCF
    pCCF = Split("8170", ",")
    'CCS aswell
    pNoCCS = Split("3290,3780,3791,D4003,8180", ",")

    For i = 1 To ccu_n
    
        If WorksheetFunction.SUM(Range("G" & i & ":S" & i)) > 0 Then
            oClick.Artno = Worksheets("CLICKING").Range("D" & ccu_i).Value
            oClick.Color = Worksheets("CLICKING").Range("E" & ccu_i).Value
            oClick.Category = Worksheets("CLICKING").Range("F" & ccu_i).Value
            
            oClick.Jobno = Worksheets("CLICKING").Range("C" & ccu_i).Value
            oClick.Plan = Worksheets("CLICKING").Range("T" & ccu_i).Value
            
            ' Common sized CCP1
            If UBound(Filter(pxCCP1, oClick.ArticleCategory)) = 0 Then
                oClick.Process = "CCP1"
                oClick.WriteToSheet rowNo:=n, colNo:=ccu_i
                n = n + 1
            End If

            ' Common sized CCS
            If UBound(Filter(pxCCS, oClick.ArticleCategory)) = 0 Then
                oClick.Process = "CCS"
                oClick.WriteToSheet rowNo:=n, colNo:=ccu_i
                n = n + 1
            End If

            ' CCP1 or CCF process
            If UBound(Filter(pCCP1, oClick.ArticleCategory)) = 0 Then
                oClick.Process = "CCP1"
                For j = 1 To 13
                    If IsEmpty(Worksheets("CLICKING").Cells(ccu_i, j + 6).Value) = False Or Worksheets("CLICKING").Cells(ccu_i, j + 6).Value <> 0 Then
                        oClick.WriteToSheet rowNo:=n, colNo:=ccu_i, sSize:=j
                        n = n + 1
                    End If
                Next j
            ElseIf UBound(Filter(pCCF, oClick.ArticleCategory)) = 0 Then
                oClick.Process = "CCF"
                For j = 1 To 13
                    If IsEmpty(Worksheets("CLICKING").Cells(ccu_i, j + 6).Value) = False Or Worksheets("CLICKING").Cells(ccu_i, j + 6).Value <> 0 Then
                        oClick.WriteToSheet rowNo:=n, colNo:=ccu_i, sSize:=j
                        n = n + 1
                    End If
                Next j
            End If

            

            ' Articles having CCS
            If UBound(Filter(pNoCCS, oClick.ArticleCategory)) <> 0 Then
                oClick.Process = "CCS"
                For j = 1 To 13
                    If IsEmpty(Worksheets("CLICKING").Cells(ccu_i, j + 6).Value) = False Or Worksheets("CLICKING").Cells(ccu_i, j + 6).Value <> 0 Then
                        oClick.WriteToSheet rowNo:=n, colNo:=ccu_i, sSize:=j
                        n = n + 1
                    End If
                Next j
            End If

            ccu_i = ccu_i + 1
        End If
    Next i
  
End Sub
