Attribute VB_Name = "CLICK_ENTRY"

' VERSION 0.4
' CLICK_ENTRY Macro, specific to Fortune.Br
' CLICKING DATA FOR SAP ENTRY, requires a sheet with name "datas" to work without any error.   Created By, Sabareesh A P ;-)
'
' Keyboard Shortcut: Ctrl+Shift+E
'
Sub CLICK_ENTRY()
    Dim ccs_s As Long
    Dim ccp_s As Long
    Dim ccs_n As Long
    Dim ccp_n As Long
    Dim ccs_t As String
    Dim n As Integer

    Worksheets("datas").cellS.Clear    'Clearing all data on the "datas" sheet before writing out
    
    With Worksheets("CLICKING")
    'Checking Insole, Upper entries
        On Error Resume Next
            ccp_s = Application.WorksheetFunction.MATCH("INSOLE", .Range("B:B"), 0)
        On Error GoTo 0
        
        If ccp_s <> 0 Then
            ccp_n = Worksheets("CLICKING").Range("B" & ccp_s).MergeArea.Rows.Count
            'MsgBox ccp_n
        Else
            MsgBox "INSOLE not found"
        End If
        
        On Error Resume Next
            ccs_s = Application.WorksheetFunction.MATCH("UPPER", .Range("B:B"), 0)
        On Error GoTo 0
        
        If ccs_s <> 0 Then
            'MsgBox "Found at " & .Cells(i, 2).Address(0, 0)
            ccs_n = Worksheets("CLICKING").Range("B" & ccs_s).MergeArea.Rows.Count
        Else
            MsgBox "UPPER not found"
        End If
           'PULLED NO.OF CELLS AND INITIAL CELL REF FOR CCP AND CCS FROM CLICKING SHEET
    End With

    Worksheets("datas").Range("B1").Value = "JOB NO."
    Worksheets("datas").Range("C1").Value = "SAP ITEM CODE"
    Worksheets("datas").Range("D1").Value = "QTY"
    Worksheets("datas").Range("E1").Value = "H. WHR"
    Worksheets("datas").Range("F1").Value = "C. WHR"
    Worksheets("datas").Range("I1").Value = "qty"
    Worksheets("datas").Range("J1").Value = "plan"
    n = 3
'##################################################################################################################################################
   ' CCP ENTRY
'##################################################################################################################################################
    For i = 1 To ccp_n
        ARTICLENO = Worksheets("CLICKING").Range("D" & ccp_s).Value
        ARTCOL = Color(Worksheets("CLICKING").Range("E" & ccp_s).Value)
        ARTICLEMODEL = ARTICLENO & "-" & ARTCOL & "-" & Worksheets("CLICKING").Range("F" & ccp_s).Value
        JOBNUM = Worksheets("CLICKING").Range("C" & ccp_s).Value
        'If IsEmpty(ARTICLENO) = False Then 'CHANCE TO FORGET TYPING ARTICLE NUMBER, SO EXCLUDED THIS IF STATEMENT
        
            For j = 1 To 13
                If IsEmpty(Worksheets("CLICKING").cellS(ccp_s, j + 6).Value) = False Or Worksheets("CLICKING").cellS(ccp_s, j + 6).Value <> 0 Then
                    Worksheets("datas").Range("A" & n).Value = j
                    Worksheets("datas").Range("B" & n).Value = JOBNUM
                    Worksheets("datas").Range("C" & n).Value = "4-CCP-" & ARTICLEMODEL & WorksheetFunction.Text(j, "00")
                    Worksheets("datas").Range("I" & n).Value = Worksheets("CLICKING").cellS(ccp_s, j + 6)
                    Worksheets("datas").Range("J" & n).Value = "=CLICKING!$T$" & ccp_s
                    Worksheets("datas").Range("D" & n).Value = "=I" & n & "*J" & n
                    Worksheets("datas").Range("E" & n).Value = "FB/CF001"
                    Worksheets("datas").Range("F" & n).Value = "FB/CF001"
                    n = n + 1
                End If
            Next j
        ccp_s = ccp_s + 1
        'End If
    Next i
'##################################################################################################################################################
   ' CCS ENTRY
'##################################################################################################################################################
    n = n + 2

    Dim arr, common_size, both_ccp1_ccs As Variant
    'ARTICLES WITH CCP1
    arr = Split("3290,3791,D4003,3780,8180,3059,1234", ",")
    'ARTICLE HAVING COMMON SIZE IN
    common_size = Split("3290,3780,3059", ",")
    'ARTICLES HAVING BOTH CCP1 and CCS
    both_ccp1_ccs = Split("3059,8170", ",")
    'ARTICLE HAVING CCF
    ccf_article = Split("8170", ",")

    For i = 1 To ccs_n
        ARTICLENO = Worksheets("CLICKING").Range("D" & ccs_s).Value
        ARTCOL = Color(Worksheets("CLICKING").Range("E" & ccs_s).Value)
        ARTICLEMODEL = ARTICLENO & "-" & ARTCOL & "-" & Worksheets("CLICKING").Range("F" & ccs_s).Value
        JOBNUM = Worksheets("CLICKING").Range("C" & ccs_s).Value
        
        'If IsEmpty(ARTICLENO) = False Then 'CHANCE TO FORGET TYPING ARTICLE NUMBER, SO EXCLUDED THIS IF STATEMENT
            If UBound(Filter(arr, UCase(ARTICLENO))) = 0 Then
                ccs_t = "4-CCP1-"
            ElseIf UBound(Filter(ccf_article, UCase(ARTICLENO))) = 0 Then
                ccs_t = "4-CCF-"
            Else:
                ccs_t = "4-CCS-"
            End If
            
            If UBound(Filter(common_size, UCase(ARTICLENO))) = 0 Then
                Worksheets("datas").Range("A" & n).Value = j
                Worksheets("datas").Range("B" & n).Value = JOBNUM
                Worksheets("datas").Range("C" & n).Value = ccs_t & ARTICLEMODEL
                Worksheets("datas").Range("D" & n).Value = Worksheets("CLICKING").Range("U" & ccs_s).Value
                Worksheets("datas").Range("E" & n).Value = "FB/CF001"
                Worksheets("datas").Range("F" & n).Value = "FB/CF001"
                Worksheets("datas").Range("I" & n).Value = Worksheets("CLICKING").Range("U" & ccs_s).Value
                Worksheets("datas").Range("J" & n).Value = "=CLICKING!$T$" & ccs_s
                n = n + 1
                
            Else:
                For j = 1 To 13
                    If IsEmpty(Worksheets("CLICKING").cellS(ccs_s, j + 6).Value) = False Or Worksheets("CLICKING").cellS(ccs_s, j + 6).Value <> 0 Then
                
                        Worksheets("datas").Range("A" & n).Value = j
                        Worksheets("datas").Range("B" & n).Value = JOBNUM
                        Worksheets("datas").Range("C" & n).Value = ccs_t & ARTICLEMODEL & WorksheetFunction.Text(j, "00")
                        Worksheets("datas").Range("D" & n).Value = "=I" & n & "*J" & n
                        Worksheets("datas").Range("E" & n).Value = "FB/CF001"
                        Worksheets("datas").Range("F" & n).Value = "FB/CF001"
                        Worksheets("datas").Range("I" & n).Value = Worksheets("CLICKING").cellS(ccs_s, j + 6)
                        Worksheets("datas").Range("J" & n).Value = "=CLICKING!$T$" & ccs_s
                        
                        n = n + 1
                    End If
                Next j
            End If
            
'#############    IF ARTICLE HAS BOTH CCP1 & CCS
            If UBound(Filter(both_ccp1_ccs, UCase(ARTICLENO))) = 0 Then
                ccs_t = "4-CCS-"
                For j = 1 To 13
                    If IsEmpty(Worksheets("CLICKING").cellS(ccs_s, j + 6).Value) = False Or Worksheets("CLICKING").cellS(ccs_s, j + 6).Value <> 0 Then
                
                        Worksheets("datas").Range("A" & n).Value = j
                        Worksheets("datas").Range("B" & n).Value = JOBNUM
                        Worksheets("datas").Range("C" & n).Value = ccs_t & ARTICLEMODEL & WorksheetFunction.Text(j, "00")
                        Worksheets("datas").Range("D" & n).Value = "=I" & n & "*J" & n
                        Worksheets("datas").Range("E" & n).Value = "FB/CF001"
                        Worksheets("datas").Range("F" & n).Value = "FB/CF001"
                        Worksheets("datas").Range("I" & n).Value = Worksheets("CLICKING").cellS(ccs_s, j + 6)
                        Worksheets("datas").Range("J" & n).Value = "=CLICKING!$T$" & ccs_s
                        
                        n = n + 1
                    End If
                Next j
            End If
'#######################################
            
        ccs_s = ccs_s + 1
        'End If
    Next i
  
End Sub
'#################################################################################################################################################################
'#################################################################################################################################################################
Function Color(colour As String)
    Dim col As String
    colour = UCase(colour)
    Select Case colour
        Case "BLACK"
        col = "BK"
        Case "BROWN"
        col = "BR"
        Case "BLUE"
        col = "BL"
        Case "RED"
        col = "RD"
        Case "PINK"
        col = "PK"
        Case "TAN"
        col = "TA"
        Case "PINK BLUE"
        col = "PE"
        Case "BLUE RED"
        col = "LR"
        Case "GREY"
        col = "GY"
        Case "GOLD"
        col = "GD"
        Case "COPPER"
        col = "CO"
        Case "WHITE"
        col = "WT"
        Case "GREEN"
        col = "GR"
        Case "ORANGE"
        col = "OR"
        Case "N.BLUE"
        col = "NB"
        Case "D.GREEN"
        col = "DN"
        Case "MEHANDI"
        col = "MH"
        Case "PEACH"
        col = "PH"
        Case "WK"
        col = "WK"
        Case "OV"
        col = "OV"
        Case "OLIVE"
        col = "OV"
        Case "SK BLACK"
        col = "SK"
        Case "TAN BLACK"
        col = "TB"
        Case "MAROON"
        col = "MR"
        Case "MR"
        col = "MR"
        Case "TB"
        col = "TB"
        Case "ST"
        col = "ST"
        Case "SA"
        col = "SA"
        Case "NR"
        col = "NR"
        Case "NG"
        col = "NG"
        Case "SK"
        col = "SK"
        Case "BK"
        col = "BK"
        Case "BR"
        col = "BR"
        Case "BL"
        col = "BL"
        Case "WT"
        col = "WT"
        Case "RD"
        col = "RD"
        Case "PK"
        col = "PK"
        Case "TA"
        col = "TA"
        Case "PE"
        col = "PE"
        Case "LR"
        col = "LR"
        Case "GY"
        col = "GY"
        Case "KG"
        col = "KG"
        Case "GD"
        col = "GD"
        Case "CO"
        col = "CO"
        Case "GR"
        col = "GR"
        Case "OR"
        col = "OR"
        Case "NB"
        col = "NB"
        Case "DN"
        col = "DN"
        Case "PH"
        col = "PH"
        Case "TR"
        col = "TR"
        Case "MH"
        col = "MH"
        Case "SE"
        col = "SE"
        Case "NY"
        col = "NY"
        Case "LY"
        col = "LY"
        Case Else
        col = "NOT-FOUND"
    End Select
    
    Color = col
End Function


'Created by, Sabareesh A P ;-)
