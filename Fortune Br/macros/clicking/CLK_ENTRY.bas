Attribute VB_Name = "CLK_ENTRY"

Sub CLICKING_ENTRY()
    Dim ccu_i As Integer
    Dim cci_i As Integer
    Dim ccu_n As Integer
    Dim cci_n As Integer
    Dim ccProcess As String
    Dim n As Integer


    Dim oClick As New clsClick
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
    
    '## SECTION INSOLE
    For i = 1 To cci_n
        If Worksheets("CLICKING").Range("U" & cci_i).Value > 0 Then
            
            oClick.Process = "CCP"
            oClick.Artno = Worksheets("CLICKING").Range("D" & cci_i).Value
            oClick.Color = Worksheets("CLICKING").Range("E" & cci_i).Value
            oClick.Category = Worksheets("CLICKING").Range("F" & cci_i).Value

            oClick.Jobno = Worksheets("CLICKING").Range("C" & cci_i).Value
            oClick.Plan = Worksheets("CLICKING").Range("T" & cci_i).Value
            
            '########################################################################################################################
            ' If mentioned "P" in column A and Article is in INS_CCP1_SEP table, only CCP1 will be generated not CCP
            If UCase(Worksheets("CLICKING").Range("A" & cci_i).Value) = "P" And oClick.MatchUp(oClick.Arr_IP1_SEP, oClick.ArticleCategory) Then
                oClick.Process = "CCP1"
                '##  Common sized CCP1
                If oClick.MatchUp(oClick.Arr_IP1_COM, oClick.ArticleCategory) Then
                    oClick.WriteToSheet rowNo:=n, colNo:=cci_i
                    n = n + 1
                End If
                
                '## All size CCP1 only
                If oClick.MatchUp(oClick.Arr_IP1, oClick.ArticleCategory) Then
                    For j = 1 To 13
                        If IsEmpty(Worksheets("CLICKING").cellS(cci_i, j + 6).Value) = False Or Worksheets("CLICKING").cellS(cci_i, j + 6).Value <> 0 Then
                            oClick.WriteToSheet rowNo:=n, colNo:=cci_i, sSize:=j
                            n = n + 1
                        End If
                    Next j
                End If

            Else
                ' Normal INSOLE Entry
                For j = 1 To 13
                    If IsEmpty(Worksheets("CLICKING").cellS(cci_i, j + 6).Value) = False Or Worksheets("CLICKING").cellS(cci_i, j + 6).Value <> 0 Then
                        oClick.WriteToSheet rowNo:=n, colNo:=cci_i, sSize:=j
                        n = n + 1
                    End If
                Next j
                
                '##  Common sized CCP1 #comes with INSOLE entry  ##
                If oClick.MatchUp(oClick.Arr_IP1_COM, oClick.ArticleCategory) Then
                    oClick.Process = "CCP1"
                    oClick.WriteToSheet rowNo:=n, colNo:=cci_i
                    n = n + 1
                End If
    
                '## CCP1 with insole
                If oClick.MatchUp(oClick.Arr_IP1, oClick.ArticleCategory) Then
                    oClick.Process = "CCP1"
                    For j = 1 To 13
                        If IsEmpty(Worksheets("CLICKING").cellS(cci_i, j + 6).Value) = False Or Worksheets("CLICKING").cellS(cci_i, j + 6).Value <> 0 Then
                            oClick.WriteToSheet rowNo:=n, colNo:=cci_i, sSize:=j
                            n = n + 1
                        End If
                    Next j
                End If
            
            End If

            

            cci_i = cci_i + 1
        End If
    Next i
    
    
    '## PAGE BREAK
    n = n + 1
    
    
    '## SECTION UPPER
    For i = 1 To ccu_n

        If Worksheets("CLICKING").Range("U" & ccu_i).Value > 0 Then
            oClick.Artno = Worksheets("CLICKING").Range("D" & ccu_i).Value
            oClick.Color = Worksheets("CLICKING").Range("E" & ccu_i).Value
            oClick.Category = Worksheets("CLICKING").Range("F" & ccu_i).Value

            oClick.Jobno = Worksheets("CLICKING").Range("C" & ccu_i).Value
            oClick.Plan = Worksheets("CLICKING").Range("T" & ccu_i).Value

            ' Common sized CCP1
            If oClick.MatchUp(oClick.Arr_UP1_COM, oClick.ArticleCategory) Then
                oClick.Process = "CCP1"
                oClick.WriteToSheet rowNo:=n, colNo:=ccu_i
                n = n + 1
            End If

            ' Common sized CCS
            If oClick.MatchUp(oClick.Arr_US_COM, oClick.ArticleCategory) Then
                oClick.Process = "CCS"
                oClick.WriteToSheet rowNo:=n, colNo:=ccu_i
                n = n + 1
            End If

            ' CCP1 or CCF process
            If oClick.MatchUp(oClick.Arr_UP1, oClick.ArticleCategory) Then
                oClick.Process = "CCP1"
                For j = 1 To 13
                    If IsEmpty(Worksheets("CLICKING").cellS(ccu_i, j + 6).Value) = False Or Worksheets("CLICKING").cellS(ccu_i, j + 6).Value <> 0 Then
                        oClick.WriteToSheet rowNo:=n, colNo:=ccu_i, sSize:=j
                        n = n + 1
                    End If
                Next j
                ElseIf oClick.MatchUp(oClick.Arr_UF, oClick.ArticleCategory) Then
                oClick.Process = "CCF"
                For j = 1 To 13
                    If IsEmpty(Worksheets("CLICKING").cellS(ccu_i, j + 6).Value) = False Or Worksheets("CLICKING").cellS(ccu_i, j + 6).Value <> 0 Then
                        oClick.WriteToSheet rowNo:=n, colNo:=ccu_i, sSize:=j
                        n = n + 1
                    End If
                Next j
            End If



            ' Articles having CCS
            If oClick.MatchUp(oClick.Arr_UNOS, oClick.ArticleCategory) = False Then
                oClick.Process = "CCS"
                For j = 1 To 13
                    If IsEmpty(Worksheets("CLICKING").cellS(ccu_i, j + 6).Value) = False Or Worksheets("CLICKING").cellS(ccu_i, j + 6).Value <> 0 Then
                        oClick.WriteToSheet rowNo:=n, colNo:=ccu_i, sSize:=j
                        n = n + 1
                    End If
                Next j
            End If


        End If
        ccu_i = ccu_i + 1
    Next i

    Worksheets("datas").Activate
End Sub


' Created by kalaLokia

