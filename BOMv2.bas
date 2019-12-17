Attribute VB_Name = "BOMv2"
Public artSize() As Long
Public siz, row_count, itemCount, scCount() As Integer
Public article, brandSize As String
Public cellX() As Long


Sub BOM()
'TEST SAMPLE, requires "BOM", "LINE" & "TREE" sheets to excecute


Dim artNo, artColor, artCat As String
Dim c, r, l As Integer

Dim mc, sc, fu, mpu, pcs, ccp, ccp1, ccs, fcmx, mcs, art As String
Dim scf(2) As String
Dim scs(2), fcs(2), fcm(2) As String

Dim rng As Range




mc = "2-FB-"
sc = "3-FB-"
mpu = "4-MPU-"
fu = "4-FU-"
ccp = "4-CCP-"
ccp1 = "4-CCP1-"
pcs = "4-PCS-"
pcs1 = "4-PCS1-"
ccs = "4-CCS-"
mcs = "4-MCS-"
fcm(0) = "4-FCM-"
fcm(1) = "4-FCM1-"
fcs(0) = "4-FCS-"
fcs(1) = "4-FCS1-"
fcs(2) = "4-FCS2-"
scs(0) = "4-SCS-"
scs(1) = "4-SCS1-"
scs(2) = "4-SCS2-"
scf(0) = "4-SCF-"
scf(1) = "4-SCF1-"
scf(2) = "4-SCF2-"
Dim n As Integer
r = 3
row_count = 3
c = n = 0

artNo = Worksheets("BOM").Range("D3")
artColor = Worksheets("BOM").Range("D4")
artCat = Worksheets("BOM").Range("D5")
artSize = SIZE_DECODE(Worksheets("BOM").Range("D7"))
siz = artSize(1) - artSize(0)
article = artNo & "-" & artColor & "-" & artCat
If InStr(1, artNo, "Z", vbTextCompare) > 0 Then
        brandSize = UCase(Worksheets("BOM").Range("D7") & "Z")
Else
        brandSize = UCase(Worksheets("BOM").Range("D7"))
End If
scCount = MC_ITEMS(brandSize) 'small carton counts in mc


'Master Carton - MC
    If C_LOOK("MC", "B") > 0 Then
        MASTER_CARTON "2-fb-" & article
    End If
    
     If C_LOOK("SC", "B") > 0 Then
        SMALL_CARTON "3-fb-" & article
    End If
    
     If C_LOOK("MPU", "B") > 0 Then
        MOULDED_PU "4-mpu-" & article
    End If
    
     If C_LOOK("FU", "B") > 0 Then
        FINSHED_UPPER "4-fu-" & article
    End If




End Sub
'Master Carton - MC
Sub MASTER_CARTON(ite As String)
    itemCount = 0
    cellX = CELL_X("MC")
    For i = 0 To siz
        LINE_CELLS ite & scCount(6), itemCount, "3-fb-" & article & WorksheetFunction.Text(artSize(0) + i, "00"), scCount(i), "4"
    Next i
    For j = 0 To cellX(1) - 1
        If IsEmpty(Worksheets("BOM").Range("D" & cellX(0) + j).Value) = False Then
            LINE_CELLS ite & scCount(6), itemCount, Worksheets("BOM").Range("D" & cellX(0) + j), Worksheets("BOM").cellS(cellX(0) + j, 6), 4
        End If
    Next j
    LINE_CELLS ite & scCount(6), itemCount, "FGMC-OH", scCount(i), 290
    row_count = row_count + itemCount
End Sub

'Small Carton - SC
Sub SMALL_CARTON(ite As String)
    cellX = CELL_X("SC")
    For i = 0 To siz
         itemCount = 0
        LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, "4-MPU-" & article & WorksheetFunction.Text(artSize(0) + i, "00"), 1, 4
        For j = 0 To cellX(1) - 1
            If IsEmpty(Worksheets("BOM").Range("D" & cellX(0) + j).Value) = False Then
                LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, Worksheets("BOM").Range("D" & cellX(0) + j), Worksheets("BOM").cellS(cellX(0) + j, 6), 4
            End If
        Next j
        LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, "FGSC-OH", 1, 290
        row_count = row_count + itemCount
    Next i
End Sub

Sub MOULDED_PU(ite As String)
    cellX = CELL_X("MPU")
    For i = 0 To siz
        itemCount = 0
        LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, "4-fu-" & article & WorksheetFunction.Text(artSize(0) + i, "00"), 1, 4
        LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, "4-PUX-0004", Worksheets("BOM").cellS(cellX(0), i + 6), 4
        
        If C_LOOK("SOFT", "C") > 0 Then
            LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, "5-PO01-0018", Worksheets("BOM").cellS(C_LOOK("SOFT", "C"), i + 6) * 34 / 134, 4
            LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, "5-PO01-0004", Worksheets("BOM").cellS(C_LOOK("SOFT", "C"), i + 6) - (Worksheets("BOM").cellS(C_LOOK("SOFT", "C"), i + 6) * 34 / 134), 4
        End If
         LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, "6-ADH-0029", 0.0003, 4
         LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, "6-CHM-0126", 0.0008, 4
        LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, "MPU-OH", 1, 290
       row_count = row_count + itemCount
    Next i
End Sub

Sub FINISHED_UPPER(ite As String)
   cellX = CELL_X("FU")
    Let x = r
    For i = 0 To siz
        n = 0
        If C_LOOK("CCP", "B") > 0 Then
            Worksheets("LINE").Range("B" & r + n).Value = n
            Worksheets("LINE").Range("C" & r + n).Value = pcs & art & WorksheetFunction.Text(s1 + i, "00")
            Worksheets("LINE").Range("D" & r + n).Value = 1
            Worksheets("LINE").Range("H" & r + n).Value = 4
            n = n + 1
        End If
        If C_LOOK("CCP1", "B") > 0 Then
            Worksheets("LINE").Range("B" & r + n).Value = n
            Worksheets("LINE").Range("C" & r + n).Value = pcs1 & art & WorksheetFunction.Text(s1 + i, "00")
            Worksheets("LINE").Range("D" & r + n).Value = 1
            Worksheets("LINE").Range("H" & r + n).Value = 4
            n = n + 1
        End If
        If C_LOOK("CCS", "B") > 0 Then
            Worksheets("LINE").Range("B" & r + n).Value = n
            Worksheets("LINE").Range("C" & r + n).Value = ccs & art & WorksheetFunction.Text(s1 + i, "00")
            Worksheets("LINE").Range("D" & r + n).Value = 1
            Worksheets("LINE").Range("H" & r + n).Value = 4
            n = n + 1
        End If
        If C_LOOK("MARK", "B") > 0 Then
            Worksheets("LINE").Range("B" & r + n).Value = n
            Worksheets("LINE").Range("C" & r + n).Value = mcs & art & WorksheetFunction.Text(s1 + i, "00")
            Worksheets("LINE").Range("D" & r + n).Value = 1
            Worksheets("LINE").Range("H" & r + n).Value = 4
            n = n + 1
        End If

        If C_LOOK("FOLD", "B") > 0 Then
            If C_LOOK("FCS", "C") > 0 Then
                Worksheets("LINE").Range("B" & r + n).Value = n
                Worksheets("LINE").Range("C" & r + n).Value = fcs(0) & art
                Worksheets("LINE").Range("D" & r + n).Value = Worksheets("BOM").cellS(C_LOOK("FCS", "C"), i + 6)
                Worksheets("LINE").Range("H" & r + n).Value = 4
                n = n + 1
            End If
            If C_LOOK("FCS1", "C") > 0 Then
                Worksheets("LINE").Range("B" & r + n).Value = n
                Worksheets("LINE").Range("C" & r + n).Value = fcs(1) & art
                Worksheets("LINE").Range("D" & r + n).Value = Worksheets("BOM").cellS(C_LOOK("FCS1", "C"), i + 6)
                Worksheets("LINE").Range("H" & r + n).Value = 4
                n = n + 1
            End If
            If C_LOOK("FCS2", "C") > 0 Then
                Worksheets("LINE").Range("B" & r + n).Value = n
                Worksheets("LINE").Range("C" & r + n).Value = fcs(2) & art
                Worksheets("LINE").Range("D" & r + n).Value = Worksheets("BOM").cellS(C_LOOK("FCS2", "C"), i + 6)
                Worksheets("LINE").Range("H" & r + n).Value = 4
                n = n + 1
            End If
        End If
  
        If C_LOOK("SLIT", "B") > 0 Then
            If C_LOOK("SCS", "C") > 0 Then
                Worksheets("LINE").Range("B" & r + n).Value = n
                Worksheets("LINE").Range("C" & r + n).Value = scs(0) & art
                Worksheets("LINE").Range("D" & r + n).Value = Worksheets("BOM").cellS(C_LOOK("SCS", "C"), i + 6)
                Worksheets("LINE").Range("H" & r + n).Value = 4
                n = n + 1
            End If
            If C_LOOK("SCS1", "C") > 0 Then
                Worksheets("LINE").Range("B" & r + n).Value = n
                Worksheets("LINE").Range("C" & r + n).Value = scs(1) & art
                Worksheets("LINE").Range("D" & r + n).Value = Worksheets("BOM").cellS(C_LOOK("SCS1", "C"), i + 6)
                Worksheets("LINE").Range("H" & r + n).Value = 4
                n = n + 1
            End If
            If C_LOOK("SCS2", "C") > 0 Then
                Worksheets("LINE").Range("B" & r + n).Value = n
                Worksheets("LINE").Range("C" & r + n).Value = scs(2) & art
                Worksheets("LINE").Range("D" & r + n).Value = Worksheets("BOM").cellS(C_LOOK("SCS2", "C"), i + 6)
                Worksheets("LINE").Range("H" & r + n).Value = 4
                n = n + 1
            End If
        End If
    
        For j = 0 To cellX(1) - 1
            If IsEmpty(Worksheets("BOM").Range("D" & cellX(0) + j).Value) = False Then
                Worksheets("LINE").Range("B" & r + n).Value = n
                Worksheets("LINE").Range("C" & r + n).Value = Worksheets("BOM").Range("D" & cellX(0) + j)
                Worksheets("LINE").Range("D" & r + n).Value = Worksheets("BOM").cellS(cellX(0) + j, i + 6)
                Worksheets("LINE").Range("H" & r + n).Value = 4
                n = n + 1
            End If
        Next j
        r = r + n
        Worksheets("LINE").Range("B" & r).Value = n
        Worksheets("LINE").Range("C" & r).Value = "STITCHING-CHARGES"
        Worksheets("LINE").Range("B" & r + 1).Value = n + 1
        Worksheets("LINE").Range("C" & r + 1).Value = "STITCH-OH"
        Worksheets("LINE").Range("D" & r).Value = 1
        Worksheets("LINE").Range("D" & r + 1).Value = 1
        Worksheets("LINE").Range("H" & r).Value = 290
        Worksheets("LINE").Range("H" & r + 1).Value = 290
        r = r + 2
        Set rng = Worksheets("LINE").Range("A" & x & ":A" & r - 1)
        rng.Value = fu & art & WorksheetFunction.Text(s1 + i, "00")
        x = r
    Next i

End Sub

Sub LINE_CELLS(valueA, valueB, valueC, valueD, valueH As String)
    Worksheets("LINE").Range("A" & row_count + itemCount).Value = UCase(valueA)
    Worksheets("LINE").Range("B" & row_count + itemCount).Value = UCase(valueB)
    Worksheets("LINE").Range("C" & row_count + itemCount).Value = UCase(valueC)
    Worksheets("LINE").Range("D" & row_count + itemCount).Value = UCase(valueD)
    Worksheets("LINE").Range("H" & row_count + itemCount).Value = UCase(valueH)
    itemCount = itemCount + 1
End Sub



Function SIZE_DECODE(siz As String) As Long()
    Dim result(2) As Long
    Dim RE As Object
    Dim MATCH As Object
    Set RE = CreateObject("VBScript.regexp")
    RE.Pattern = "(\d{1,2})"
    RE.Global = True
    RE.ignoreCase = False
    Set MATCH = RE.Execute(siz)
    If MATCH.Count <> 0 Then
        result(0) = MATCH.Item(0).submatches.Item(0)
        If MATCH.Count > 1 Then
            result(1) = MATCH.Item(1).submatches.Item(0)
        Else:
            result(1) = MATCH.Item(0).submatches.Item(0)
        End If
    End If
    SIZE_DECODE = result
End Function


Function CELL_X(cell_name As String) As Long()
    Dim cellS(2) As Long
    
    With Worksheets("BOM")
        On Error Resume Next
        cellS(0) = Application.WorksheetFunction.MATCH(cell_name, .Range("B:B"), 0)
        
        On Error GoTo 0
        
        If cellS(0) <> 0 Then
        
            cellS(1) = Worksheets("BOM").Range("B" & cellS(0)).MergeArea.Rows.Count
        Else
            'NOT-FOUND
        End If
    End With
    
    CELL_X = cellS
End Function

Function C_LOOK(lookUpValue As String, colmnName As String)
    Dim cellS As Long
    With Worksheets("BOM")
        On Error Resume Next
        cellS = Application.WorksheetFunction.MATCH(lookUpValue, .Range(colmnName & ":" & colmnName), 0)
        On Error GoTo 0
    End With
    C_LOOK = cellS
End Function

Function MC_ITEMS(sizee As String) As Integer()
    Dim sc_count(6) As Integer
        
        Select Case sizee
            Case "6X10"
                sc_count(0) = 3
                sc_count(4) = 3
                sc_count(1) = 6
                sc_count(2) = 6
                sc_count(3) = 6
                sc_count(5) = 24
                sc_count(6) = 1
                    
            Case "6X10Z"
                sc_count(0) = 2
                sc_count(1) = 2
                sc_count(2) = 3
                sc_count(3) = 3
                sc_count(4) = 2
                sc_count(5) = 12
                sc_count(6) = 1
                
            Case "5X9"
                sc_count(0) = 7
                sc_count(1) = 7
                sc_count(2) = 7
                sc_count(3) = 7
                sc_count(4) = 2
                sc_count(5) = 30
                sc_count(6) = 1
            Case "5X8"
                sc_count(0) = 8
                sc_count(1) = 8
                sc_count(2) = 7
                sc_count(3) = 7
                sc_count(4) = 30
                sc_count(5) = 0
                sc_count(6) = 2
            Case "1X3"
                sc_count(0) = 10
                sc_count(1) = 10
                sc_count(2) = 10
                sc_count(3) = 30
                sc_count(4) = 0
                sc_count(5) = 0
                sc_count(6) = 2
            Case "1X5"
                sc_count(0) = 6
                sc_count(1) = 6
                sc_count(2) = 6
                sc_count(3) = 6
                sc_count(4) = 6
                sc_count(5) = 30
                sc_count(6) = 1
            Case "11X13"
                sc_count(0) = 12
                sc_count(1) = 12
                sc_count(2) = 12
                sc_count(3) = 36
                sc_count(4) = 0
                sc_count(5) = 0
                sc_count(6) = 1
            Case "8X10"
                sc_count(0) = 12
                sc_count(1) = 12
                sc_count(2) = 12
                sc_count(3) = 36
                sc_count(4) = 0
                sc_count(5) = 0
                sc_count(6) = 4
            Case "11X12"
                sc_count(0) = 6
                sc_count(1) = 6
                sc_count(2) = 12
                sc_count(3) = 0
                sc_count(4) = 0
                sc_count(5) = 0
                sc_count(6) = 1
            Case Else
                sc_count(0) = 0
                sc_count(1) = 0
                sc_count(2) = 0
                sc_count(3) = 0
                sc_count(4) = 0
                sc_count(5) = 0
                sc_count(6) = 0
        End Select
    MC_ITEMS = sc_count()
End Function

