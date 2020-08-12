Attribute VB_Name = "StallionToSap"
Sub StallionToSAP()

    Dim rowCount As Integer
    rowCount = Worksheets("Sheet1").Range("B" & Rows.Count).End(xlUp).Row
    
    For i = 2 To rowCount
        Worksheets("Sheet1").Range("D" & i).Value = ART_NO(Worksheets("Sheet1").Range("B" & i)) & "-" & COLOR(Worksheets("Sheet1").Range("B" & i)) & "-" & CATG(Worksheets("Sheet1").Range("B" & i)) & SIZEE(Worksheets("Sheet1").Range("B" & i))
    Next i


End Sub



'DO NOT TOUCH IN THIS CODE
Function SIZEE(artic As String)
Dim RE As Object
Dim MATCH As Object
Set RE = CreateObject("VBScript.regexp")
RE.Pattern = "(\d{1,}$)"
RE.Global = True
RE.ignoreCase = False
Set MATCH = RE.Execute(artic)
If MATCH.Count <> 0 Then
result = MATCH.Item(0).submatches.Item(0)
End If
If result > 39 Then
    result = result - 34
End If
SIZEE = WorksheetFunction.Text(result, "00")
End Function

'DO NOT TOUCH IN THIS CODE
Function ART_NO(artic As String)
Dim RE As Object
Dim MATCH As Object
Set RE = CreateObject("VBScript.regexp")
RE.Pattern = "(\b(D|S|L|K)?\d{4})"
RE.Global = True
RE.ignoreCase = False
Set MATCH = RE.Execute(artic)
If MATCH.Count <> 0 Then
result = MATCH.Item(0).submatches.Item(0)
End If
ART_NO = result
End Function


Function CATG(artic As String)

Dim RE As Object
Dim MATCH As Object
Set RE = CreateObject("VBScript.regexp")
RE.Pattern = "(?!ZSP)([A-Z]{3,})"
RE.Global = True
RE.ignoreCase = False
Set MATCH = RE.Execute(artic)
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
Case Else
cat = "NOT-FOUND"
End Select

CATG = cat
End Function


'COLOR CODES IN HERE, ADD OR EDIT IS ALLOWED
Function COLOR(colour As String)

Dim RE As Object
Dim MATCH As Object
Set RE = CreateObject("VBScript.regexp")
RE.Pattern = "(?!ZSP)([A-Z,a-z,-]{3,})"
RE.Global = True
RE.ignoreCase = True
Set MATCH = RE.Execute(colour)
If MATCH.Count > 1 Then
    reslt = MATCH.Item(1).submatches.Item(0)
End If

Dim col As String

reslt = UCase(reslt)
Select Case reslt
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
Case "TAN-BROWN"
col = "TR"
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
Case "BLURED"
col = "LR"
Case "NR"
col = "NR"
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
Case "N-BLUE-RED"
col = "NR"
Case "TAN-BRN"
col = "TR"

Case "RD"
col = "RD"
Case "PK"
col = "PK"
Case "TA"
col = "TA"
Case "PE"
col = "PE"
Case "SBLACK"
col = "TAN BRN"
Case "TR"
col = "SAND"
Case "SD"
col = "SK"
Case "LR"
col = "LR"
Case "TAN-BLACK"
col = "TB"
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
Case "MH"
col = "MH"
Case Else
col = "NOT-FOUND"
End Select

COLOR = col
End Function

