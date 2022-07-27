Attribute VB_Name = "INWARD"
Sub INWARD()

    Dim rowCount As Integer
    Dim tmpItem, tmpSemiItem, tmpMpuItem, art_number, art_color, art_category, art_size As String

    rowCount = Worksheets("Sheet1").Range("B" & Rows.Count).End(xlUp).Row
    
    'Create sheets if not exists else clear the cells.
    If WorksheetExists("MPU") Then
        Worksheets("SMALL CARTON").cellS.Clear
    Else:
        Sheets.Add(After:=Sheets(Sheets.Count)).name = "SMALL CARTON"
    End If
    
    If WorksheetExists("MPU") Then
        Worksheets("MPU").cellS.Clear
    Else:
        Sheets.Add(After:=Sheets(Sheets.Count)).name = "MPU"
    End If
    
    
    For i = 2 To rowCount
        tmpItem = Worksheets("Sheet1").Range("B" & i).value
        art_number = ART_NO(tmpItem)
        art_color = COLOR(tmpItem)
        art_category = CATG(tmpItem)
        art_size = SIZEE(tmpItem)
        tmpMpuItem = art_number & "-" & art_color & "-" & art_category & art_size
        tmpSemiItem = tmpMpuItem
        If art_number = "L2152" And art_color = "OL" Then
            tmpSemiItem = art_number & "-" & "OV" & "-" & art_category & art_size
        End If

        Worksheets("Sheet1").Range("D" & i).value = "3-FB-" & tmpSemiItem
        
        Worksheets("SMALL CARTON").Range("A" & i).value = "3-FB-" & tmpSemiItem
        Worksheets("SMALL CARTON").Range("B" & i).value = Worksheets("Sheet1").Range("C" & i)
        
        Worksheets("MPU").Range("A" & i).value = "4-MPU-" & tmpMpuItem
        Worksheets("MPU").Range("B" & i).value = Worksheets("Sheet1").Range("C" & i)
    Next i
    
    'REMOVE Z, S (3074S) from MPU and <space> from both MPU and SMALL CARTON
    Worksheets("MPU").Columns("A").Replace _
        What:="Z", Replacement:="", _
        SearchOrder:=xlByColumns, MatchCase:=True
    Worksheets("MPU").Columns("A").Replace _
        What:="3074S", Replacement:="3074", _
        SearchOrder:=xlByColumns, MatchCase:=True
    Worksheets("MPU").Columns("A").Replace _
        What:=" ", Replacement:="", _
        SearchOrder:=xlByColumns, MatchCase:=True
    Worksheets("SMALL CARTON").Columns("A").Replace _
        What:=" ", Replacement:="", _
        SearchOrder:=xlByColumns, MatchCase:=True
        
    'ADJUST COLUMN WIDTH
    Worksheets("MPU").Columns("A").ColumnWidth = 22
    Worksheets("SMALL CARTON").Columns("A").ColumnWidth = 22

End Sub

Function WorksheetExists(shtname As String) As Boolean
    Dim ws As Worksheet
    Dim wb As Workbook
    
    Set wb = ActiveWorkbook
    
    For Each ws In wb.Sheets
        If ws.name = shtname Then
            WorksheetExists = True
            Exit Function
        End If
    Next
    WorksheetExists = False
End Function


'###################################################
'############  SIZE FUNCTION      ##############

'DO NOT TOUCH IN THIS CODE
Function SIZEE(ByVal value As String)
value = UCase(value)
Dim RE As Object
Dim MATCH As Object
Set RE = CreateObject("VBScript.regexp")
RE.Pattern = "(\d{1,}$)"
RE.Global = True
RE.ignoreCase = False
Set MATCH = RE.Execute(value)
If MATCH.Count <> 0 Then
result = MATCH.Item(0).submatches.Item(0)
End If
SIZEE = WorksheetFunction.Text(result, "00")
End Function
'###################################################
'############  ARTICLE_NO FUNCTION      ##############
'DO NOT TOUCH IN THIS CODE
Function ART_NO(ByVal value As String)
value = UCase(value)
Dim RE As Object
'Dim RR As Object
Dim MATCH As Object
Set RE = CreateObject("VBScript.regexp")
Set RR = CreateObject("VBScript.regexp")
RE.Pattern = "(\b((D|DG|OG|DL|OL|DX|SP|K|L|GP|LP)(| )?\d{4,5}|\d{4}(S|s)?)(\s{0,3}(Z|ZSP))?\b)"
RE.Global = True
RE.ignoreCase = False
Set MATCH = RE.Execute(value)
If MATCH.Count <> 0 Then
result = MATCH.Item(0).submatches.Item(0)
End If

'RR.Pattern = "(\s*ZSP)|(\s+Z)"
'RR.Global = False

ART_NO = result

End Function

'###################################################
'############  CATEGORY FUNCTION      ##############

Function CATG(ByVal value As String)
value = UCase(value)
Dim RE As Object
Dim MATCH As Object
Set RE = CreateObject("VBScript.regexp")
RE.Pattern = "(?!ZSP)([A-Z]{3,})"
RE.Global = True
RE.ignoreCase = False
Set MATCH = RE.Execute(value)
If MATCH.Count <> 0 Then
result = MATCH.Item(0).submatches.Item(0)
End If


Dim cat As String
result = UCase(result)
Select Case result
Case "GENTS"
cat = "G"
Case "LADIES"
cat = "L"
Case "KIDS"
cat = "K"
Case "CHILDREN"
cat = "C"
Case "BOYS"
cat = "B"
Case "GIRLS"
cat = "R"
Case "INFANT"
cat = "I"
Case "GAINTS"
cat = "X"
Case "GIANTS"
cat = "X"
Case Else
cat = "NOT-FOUND"
End Select

CATG = cat
End Function


'COLOR CODES IN HERE, ADD OR EDIT IS ALLOWED
Function COLOR(ByVal colour As String)

Dim RE As Object
Dim MATCH As Object
Set RE = CreateObject("VBScript.regexp")
RE.Pattern = "((?!ZSP)([A-Z]{3,})\s)(.*)(\sFB)"
RE.Global = True
RE.ignoreCase = True
Set MATCH = RE.Execute(UCase(colour))
If MATCH.Count <> 0 Then
    reslt = MATCH.Item(0).submatches.Item(2)
End If

Dim col As String

reslt = UCase(reslt)
Select Case reslt

'#################################################
'Case "enter stallion color"
'col = "enter sap color code"
'#################################################
Case "BLK-GREY"
col = "KG"
Case "SPECIAL TAN"
col = "ST"
Case "SPECIAL BLACK"
col = "SA"
Case "COFFEE BROWN"
col = "CB"
Case "BROWN SP"
col = "SR"
Case "CAMEL"
col = "CM"
Case "CREAM"
col = "CR"
Case "BLACK(SP)"
col = "SA"
Case "D.BROWN"
col = "DR"
Case "BLACK"
col = "BK"
Case "BROWN"
col = "BR"
Case "BLACK-RED"
col = "RK"
Case "BLUE"
col = "BL"
Case "YELLOW"
col = "YL"
Case "RED"
col = "RD"
Case "BLU-PINK"
col = "LP"
Case "PINK"
col = "PK"
Case "TAN"
col = "TA"
Case "OLIVE"
col = "OL"
Case "WHITE"
col = "WT"
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
Case "N.BLU-TAN"
col = "NT"
Case "D.GREEN"
col = "DN"
Case "MEHANDI"
col = "MH"
Case "PEACH"
col = "PH"
Case "SK BLACK"
col = "SK"
Case "TAN BLACK"
col = "TB"
Case "PNK"
col = "PE"
Case "PNK-BLU"
col = "PE"
Case "MAROON"
col = "MR"
Case "NAVYBLUE"
col = "NB"
Case "BLU-BLK"
col = "BB"
Case "RED-BLACK"
col = "RK"
Case "BLURED"
col = "LR"
Case "NAVY"
col = "NY"
Case "MR"
col = "MR"
Case "TB"
col = "TB"
Case "SK"
col = "SK"
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
Case "BLACK-GRY"
col = "KG"
Case "N-BLUE-RED"
col = "NR"
Case "N.BLU-RED"
col = "NR"
Case "N.BL-RED"
col = "NR"
Case "N.BLU-GREY"
col = "NG"
Case "N.BL-GREY"
col = "NG"
Case "TA"
col = "TA"
Case "PE"
col = "PE"
Case "SBLACK"
col = "SK"
Case "TR"
col = "SAND"
Case "SD"
col = "SK"
Case "LR"
col = "LR"
Case "TAN-BLACK"
col = "TB"
Case "TAN-BRN"
col = "TR"
Case "BLK WHITE"
col = "WK"
Case "GY"
col = "GY"
Case "BEIGE"
col = "BG"
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
Case "BLUE-TAN"
col = "LT"
Case "PH"
col = "PH"
Case "TR"
col = "TR"
Case "MH"
col = "MH"
Case "SP-TAN"
col = "ST"
Case Else
col = "NOT-FOUND"
End Select

COLOR = col
End Function



