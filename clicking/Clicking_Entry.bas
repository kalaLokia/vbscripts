
Sub CLICK_ENTRY_NEW()
    Dim ccu_i As Integer
    Dim cci_i As Integer
    Dim ccu_n As Integer
    Dim cci_n As Integer
    Dim ccProcess As String
    Dim n As Integer

    Dim oClick As New clsClicking
    Worksheets("datas").cellS.Clear    'Clearing all data on the "datas" sheet before writing out
    
    With Worksheets("CLICKING")
    'Checking Insole, Upper entries
        On Error Resume Next
            cci_i = Application.WorksheetFunction.MATCH("INSOLE", .Range("B:B"), 0)
        On Error GoTo 0
        
        If cci_i <> 0 Then
            cci_n = Worksheets("CLICKING").Range("B" & cci_i).MergeArea.Rows.Count
            'MsgBox cci_n
        Else
            MsgBox "INSOLE not found"
        End If
        
        On Error Resume Next
            ccu_i = Application.WorksheetFunction.MATCH("UPPER", .Range("B:B"), 0)
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
        If Worksheets("CLICKING").Range("U" & ccu_i).Value > 0 Then
            oClick.Process = "CCP"
            oClick.Artno = Worksheets("CLICKING").Range("D" & cci_i).Value
            oClick.Color = Worksheets("CLICKING").Range("E" & cci_i).Value
            oClick.Category = Worksheets("CLICKING").Range("F" & cci_i).Value
            
            oClick.Jobno = Worksheets("CLICKING").Range("C" & cci_i).Value
            oClick.Plan = Worksheets("CLICKING").Range("T" & cci_i).Value
            
            For j = 1 To 13
                If IsEmpty(Worksheets("CLICKING").cellS(cci_i, j + 6).Value) = False Or Worksheets("CLICKING").cellS(cci_i, j + 6).Value <> 0 Then
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
    pxCCS = Split("3391B,3391", ",")
    'CCP1
    pCCP1 = Split("3791,D4003,8180", ",")
    'CCF
    pCCF = Split("8170", ",")
    'CCS aswell
    pNoCCS = Split("3290,3780,3791,D4003,8180,3391", ",")

    For i = 1 To ccu_n
    
        If Worksheets("CLICKING").Range("U" & ccu_i).Value > 0 Then
            oClick.Artno = Worksheets("CLICKING").Range("D" & ccu_i).Value
            oClick.Color = Worksheets("CLICKING").Range("E" & ccu_i).Value
            oClick.Category = Worksheets("CLICKING").Range("F" & ccu_i).Value
            
            oClick.Jobno = Worksheets("CLICKING").Range("C" & ccu_i).Value
            oClick.Plan = Worksheets("CLICKING").Range("T" & ccu_i).Value
            
            ' Common sized CCP1
            If MatchUp(pxCCP1, oClick.ArticleCategory) Then
                oClick.Process = "CCP1"
                oClick.WriteToSheet rowNo:=n, colNo:=ccu_i
                n = n + 1
            End If

            ' Common sized CCS
            If MatchUp(pxCCS, oClick.ArticleCategory) Then
                oClick.Process = "CCS"
                oClick.WriteToSheet rowNo:=n, colNo:=ccu_i
                n = n + 1
            End If

            ' CCP1 or CCF process
            If MatchUp(pCCP1, oClick.ArticleCategory) Then
                oClick.Process = "CCP1"
                For j = 1 To 13
                    If IsEmpty(Worksheets("CLICKING").cellS(ccu_i, j + 6).Value) = False Or Worksheets("CLICKING").cellS(ccu_i, j + 6).Value <> 0 Then
                        oClick.WriteToSheet rowNo:=n, colNo:=ccu_i, sSize:=j
                        n = n + 1
                    End If
                Next j
            ElseIf MatchUp(pCCF, oClick.ArticleCategory) Then
                oClick.Process = "CCF"
                For j = 1 To 13
                    If IsEmpty(Worksheets("CLICKING").cellS(ccu_i, j + 6).Value) = False Or Worksheets("CLICKING").cellS(ccu_i, j + 6).Value <> 0 Then
                        oClick.WriteToSheet rowNo:=n, colNo:=ccu_i, sSize:=j
                        n = n + 1
                    End If
                Next j
            End If

            

            ' Articles having CCS
            If MatchUp(pNoCCS, oClick.ArticleCategory) = False Then
                oClick.Process = "CCS"
                For j = 1 To 13
                    If IsEmpty(Worksheets("CLICKING").cellS(ccu_i, j + 6).Value) = False Or Worksheets("CLICKING").cellS(ccu_i, j + 6).Value <> 0 Then
                        oClick.WriteToSheet rowNo:=n, colNo:=ccu_i, sSize:=j
                        n = n + 1
                    End If
                Next j
            End If

            ccu_i = ccu_i + 1
        End If
    Next i
End Sub

' Looks up a string in an array
Function MatchUp(arr As Variant, lookUpValue As String) As Boolean

For Each element In arr
    MatchUp = False
    If StrComp(lookUpValue, element) = 0 Then
        MatchUp = True
        Exit For
    End If
Next
    
End Function

' Created by kalaLokia
