Attribute VB_Name = "CLICK_ENTRY"
Sub CLICK_ENTRY()
Attribute CLICK_ENTRY.VB_Description = "CLICKING DATA FOR SAP ENTRY, requires a sheet with name ""datas"" to work without any error. \n\nCreated By, Sabareesh A P ;-)"
Attribute CLICK_ENTRY.VB_ProcData.VB_Invoke_Func = "E\n14"
'
' CLICK_ENTRY Macro
' CLICKING DATA FOR SAP ENTRY, requires a sheet with name "datas" to work without any error.   Created By, Sabareesh A P ;-)
'
' Keyboard Shortcut: Ctrl+Shift+E
'

Dim ccs_s As Long
Dim ccp_s As Long
Dim ccs_n As Long
Dim ccp_n As Long
Dim ccs_t As String
Dim n As Integer


    With Worksheets("CLICKING")
    On Error Resume Next
    ccp_s = Application.WorksheetFunction.Match("INSOLE", .Range("B:B"), 0)
    On Error GoTo 0
    
    If ccp_s <> 0 Then
    
        ccp_n = Worksheets("CLICKING").Range("B" & ccp_s).MergeArea.Rows.Count
        'MsgBox ccp_n
    Else
        MsgBox "INSOLE not found"
        
    End If
    
     On Error Resume Next
    ccs_s = Application.WorksheetFunction.Match("UPPER", .Range("B:B"), 0)
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
' CCP ENTRY LOOP
For i = 1 To ccp_n
For j = 1 To 13
'If IsEmpty(ThisWorkbook.Worksheets("CLICKING").Range("D" & ccp_s).Value) = False Then
    If IsEmpty(Worksheets("CLICKING").Cells(ccp_s, j + 6).Value) = False Or Worksheets("CLICKING").Cells(ccp_s, j + 6).Value <> 0 Then
    
Worksheets("datas").Range("A" & n).Value = j
Worksheets("datas").Range("B" & n).Value = "=CLICKING!$C$" & ccp_s
Worksheets("datas").Range("C" & n).Value = Chr(61) & Chr(34) & "4-CCP-" & Chr(34) & Chr(38) & "CLICKING!$D$" & ccp_s & Chr(38) & Chr(34) & Chr(45) & color(Worksheets("CLICKING").Range("E" & ccp_s)) & Chr(45) & Chr(34) & Chr(38) & "CLICKING!$F$" & ccp_s & Chr(38) & "TEXT" & Chr(40) & j & Chr(44) & Chr(34) & "00" & Chr(34) & Chr(41)
Worksheets("datas").Range("I" & n).Value = Worksheets("CLICKING").Cells(ccp_s, j + 6)
Worksheets("datas").Range("J" & n).Value = "=CLICKING!$T$" & ccp_s
Worksheets("datas").Range("D" & n).Value = "=I" & n & "*J" & n
Worksheets("datas").Range("E" & n).Value = "FB/CF001"
Worksheets("datas").Range("F" & n).Value = "FB/CF001"
'color (Worksheets("CLICKING").Range("E" & ccp_s).Value)
n = n + 1
End If
'Else: ccp_n = ccp_n + 1
'End If
Next j
'n = n + 1
ccp_s = ccp_s + 1
Next i
   
   ' CCS ENTRY
n = n + 2
For i = 1 To ccs_n
For j = 1 To 13
'If IsEmpty(ThisWorkbook.Worksheets("CLICKING").Range("D" & ccp_s).Value) = False Then
    If IsEmpty(Worksheets("CLICKING").Cells(ccs_s, j + 6).Value) = False Or Worksheets("CLICKING").Cells(ccs_s, j + 6).Value <> 0 Then
If Worksheets("CLICKING").Range("D" & ccs_s) = 3290 Then
ccs_t = "4-CCP1-"
Else: ccs_t = "4-CCS-"
End If

Worksheets("datas").Range("A" & n).Value = j
Worksheets("datas").Range("B" & n).Value = "=CLICKING!$C$" & ccs_s
Worksheets("datas").Range("C" & n).Value = Chr(61) & Chr(34) & ccs_t & Chr(34) & Chr(38) & "CLICKING!$D$" & ccs_s & Chr(38) & Chr(34) & Chr(45) & color(Worksheets("CLICKING").Range("E" & ccs_s)) & Chr(45) & Chr(34) & Chr(38) & "CLICKING!$F$" & ccs_s & Chr(38) & "TEXT" & Chr(40) & j & Chr(44) & Chr(34) & "00" & Chr(34) & Chr(41)
Worksheets("datas").Range("I" & n).Value = Worksheets("CLICKING").Cells(ccs_s, j + 6)
Worksheets("datas").Range("J" & n).Value = "=CLICKING!$T$" & ccs_s
Worksheets("datas").Range("D" & n).Value = "=I" & n & "*J" & n
Worksheets("datas").Range("E" & n).Value = "FB/CF001"
Worksheets("datas").Range("F" & n).Value = "FB/CF001"
n = n + 1
End If
'Else: ccp_n = ccp_n + 1
'End If
Next j
'n = n + 1
ccs_s = ccs_s + 1
Next i
  
End Sub

Function color(colour As String)
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
Case "GREEN"
col = "GR"
Case "ORANGE"
col = "OR"
Case "N.BLUE"
col = "NB"
Case "D.GREEN"
col = "DN"
Case "PEACH"
col = "PH"
Case "BK"
col = "BK"
Case "BR"
col = "BR"
Case "BL"
col = "BL"
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
Case Else
col = "NOT-FOUND"
End Select

color = col
End Function


'Created by, Sabareesh A P ;-)
