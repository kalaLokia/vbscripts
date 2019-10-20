Attribute VB_Name = "BOM"
Sub BOMTEST()

Dim artNo, artColor, artCat, s2, s1, brandSize As String
Dim siz, c, r, l As Integer
Dim mc, sc, fu, mpu, pcs, ccp, ccp1, ccs, fcs, fcs1, fcs2, scs1, scs2, fcm, mcs, scs, fcmx, art As String
Dim scf(2) As String
Dim scCount() As Integer
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
fcm = "4-FCM"
mcs = "4-MCS-"
fcs = "4-FCS-"
scs = "4-SCS-"
fcs1 = "4-FCS1-"
scs1 = "4-SCS1-"
fcs2 = "4-FCS2-"
scs2 = "4-SCS2-"
scf(0) = "4-SCF-"
scf(1) = "4-SCF1-"
Dim n As Integer
r = 3
c = n = 0
artNo = Worksheets("BOM").Range("D3")
artColor = Worksheets("BOM").Range("D4")
artCat = Worksheets("BOM").Range("D5")
s1 = SIZE1(Worksheets("BOM").Range("D7"))
s2 = SIZE2(Worksheets("BOM").Range("D7"))
siz = s2 - s1
art = artNo & "-" & artColor & "-" & artCat
brandSize = Worksheets("BOM").Range("D7")
scCount = MC_ITEMS(brandSize) 'small carton counts in mc
Dim cellX() As Long
'cellX = CELL_X("")
'Master Carton - MC
cellX = CELL_X("MC")

Set rng = Worksheets("test111").Range("A" & r & ":A" & cellX(1) + siz + r + 1)
rng.Value = mc & art & 1
For i = 0 To siz
Worksheets("test111").Range("B" & r + i).Value = i
Worksheets("test111").Range("C" & r + i).Value = sc & art & WorksheetFunction.Text(s1 + i, "00")
Worksheets("test111").Range("D" & r + i).Value = scCount(i)
Next i
r = r + i
For j = 0 To cellX(1) - 1
    If IsEmpty(Worksheets("BOM").Range("D" & cellX(0) + j).Value) = False Then
        Worksheets("test111").Range("B" & r + n).Value = i + n
        Worksheets("test111").Range("C" & r + n).Value = Worksheets("BOM").Range("D" & cellX(0) + j)
        Worksheets("test111").Range("D" & r + n).Value = Worksheets("BOM").cellS(cellX(0) + j, 6)
        n = n + 1
    End If
Next j
r = r + n
Worksheets("test111").Range("B" & r).Value = i + n
Worksheets("test111").Range("C" & r).Value = "FGMC_OH"
Worksheets("test111").Range("D" & r).Value = scCount(i)
r = r + 1
'Small Carton - SC
cellX = CELL_X("SC")
For i = 0 To siz
    c = 0
    Worksheets("test111").Range("B" & r).Value = c
    Worksheets("test111").Range("C" & r).Value = mpu & art & WorksheetFunction.Text(s1 + i, "00")
    Worksheets("test111").Range("D" & r).Value = 1
    
    For j = 0 To cellX(1) - 1
    If IsEmpty(Worksheets("BOM").Range("D" & cellX(0) + j).Value) = False Then
        c = c + 1
        Worksheets("test111").Range("B" & c + r).Value = c
        Worksheets("test111").Range("C" & c + r).Value = Worksheets("BOM").Range("D" & cellX(0) + j)
        Worksheets("test111").Range("D" & c + r).Value = Worksheets("BOM").cellS(cellX(0) + j, 6)
        End If
    Next j
    
Worksheets("test111").Range("B" & c + r + 1).Value = c + 1
Worksheets("test111").Range("C" & c + r + 1).Value = "FGSC_OH"
Worksheets("test111").Range("D" & c + r + 1).Value = 1

Set rng = Worksheets("test111").Range("A" & r & ":A" & cellX(1) + siz + r + 2)
rng.Value = sc & art & WorksheetFunction.Text(s1 + i, "00")
r = r + c + 2
Next i

'MPU
cellX = CELL_X("SW")
For i = 0 To siz
    Worksheets("test111").Range("A" & r).Value = mpu & art & WorksheetFunction.Text(s1 + i, "00")
    Worksheets("test111").Range("B" & r).Value = 0
    Worksheets("test111").Range("C" & r).Value = fu & art & WorksheetFunction.Text(s1 + i, "00")
    Worksheets("test111").Range("D" & r).Value = 1
    
    Worksheets("test111").Range("A" & r + 1).Value = mpu & art & WorksheetFunction.Text(s1 + i, "00")
    Worksheets("test111").Range("B" & r + 1).Value = 1
    Worksheets("test111").Range("C" & r + 1).Value = "4-PUX-0003"
    Worksheets("test111").Range("D" & r + 1).Value = Worksheets("BOM").cellS(cellX(0), i + 6)
    
    Worksheets("test111").Range("A" & r + 2).Value = mpu & art & WorksheetFunction.Text(s1 + i, "00")
    Worksheets("test111").Range("B" & r + 2).Value = 2
    Worksheets("test111").Range("C" & r + 2).Value = "MPU_OH"
    Worksheets("test111").Range("D" & r + 2).Value = 1
    r = r + 3
Next i

'FU
cellX = CELL_X("FU")
Let x = r
For i = 0 To siz
    c = 0
  If C_LOOK("CCP") > 0 Then
    Worksheets("test111").Range("B" & r + c).Value = c
    Worksheets("test111").Range("C" & r + c).Value = pcs & art & WorksheetFunction.Text(s1 + i, "00")
    Worksheets("test111").Range("D" & r + c).Value = 1
    c = c + 1
  End If
   If C_LOOK("CCP1") > 0 Then
    Worksheets("test111").Range("B" & r + c).Value = c
    Worksheets("test111").Range("C" & r + c).Value = pcs1 & art & WorksheetFunction.Text(s1 + i, "00")
    Worksheets("test111").Range("D" & r + c).Value = 1
    c = c + 1
  End If
  If C_LOOK("CCS") > 0 Then
    Worksheets("test111").Range("B" & r + c).Value = c
    Worksheets("test111").Range("C" & r + c).Value = ccs & art & WorksheetFunction.Text(s1 + i, "00")
    Worksheets("test111").Range("D" & r + c).Value = 1
    c = c + 1
  End If
    If C_LOOK("MCS") > 0 Then
    Worksheets("test111").Range("B" & r + c).Value = c
    Worksheets("test111").Range("C" & r + c).Value = mcs & art & WorksheetFunction.Text(s1 + i, "00")
    Worksheets("test111").Range("D" & r + c).Value = 1
    c = c + 1
  End If
  If C_LOOK("FCS") > 0 Then
    Worksheets("test111").Range("B" & r + c).Value = c
    Worksheets("test111").Range("C" & r + c).Value = fcs & art
    Worksheets("test111").Range("D" & r + c).Value = Worksheets("BOM").cellS(C_LOOK("FCS"), i + 6)
    c = c + 1
  End If
  If C_LOOK("FCS1") > 0 Then
    Worksheets("test111").Range("B" & r + c).Value = c
    Worksheets("test111").Range("C" & r + c).Value = fcs1 & art
    Worksheets("test111").Range("D" & r + c).Value = Worksheets("BOM").cellS(C_LOOK("FCS1"), i + 6)
    c = c + 1
  End If
    If C_LOOK("FCS2") > 0 Then
    Worksheets("test111").Range("B" & r + c).Value = c
    Worksheets("test111").Range("C" & r + c).Value = fcs2 & art
    Worksheets("test111").Range("D" & r + c).Value = Worksheets("BOM").cellS(C_LOOK("FCS2"), i + 6)
    c = c + 1
  End If
    If C_LOOK("SCS") > 0 Then
    Worksheets("test111").Range("B" & r + c).Value = c
    Worksheets("test111").Range("C" & r + c).Value = scs & art
     Worksheets("test111").Range("D" & r + c).Value = Worksheets("BOM").cellS(C_LOOK("SCS"), i + 6)
    c = c + 1
  End If
  If C_LOOK("SCS1") > 0 Then
    Worksheets("test111").Range("B" & r + c).Value = c
    Worksheets("test111").Range("C" & r + c).Value = scs1 & art
     Worksheets("test111").Range("D" & r + c).Value = Worksheets("BOM").cellS(C_LOOK("SCS1"), i + 6)
    c = c + 1
  End If
  If C_LOOK("SCS2") > 0 Then
    Worksheets("test111").Range("B" & r + c).Value = c
    Worksheets("test111").Range("C" & r + c).Value = scs2 & art
     Worksheets("test111").Range("D" & r + c).Value = Worksheets("BOM").cellS(C_LOOK("SCS2"), i + 6)
    c = c + 1
  End If
  For j = 0 To cellX(1) - 1
    If IsEmpty(Worksheets("BOM").Range("D" & cellX(0) + j).Value) = False Then
        Worksheets("test111").Range("B" & r + c).Value = c
        Worksheets("test111").Range("C" & r + c).Value = Worksheets("BOM").Range("D" & cellX(0) + j)
        Worksheets("test111").Range("D" & c + r).Value = Worksheets("BOM").cellS(cellX(0) + j, i + 6)
        c = c + 1
     End If
  Next j
    r = r + c
        Worksheets("test111").Range("B" & r).Value = c
        Worksheets("test111").Range("C" & r).Value = "STITCHING-CHARGES"
        Worksheets("test111").Range("B" & r + 1).Value = c + 1
        Worksheets("test111").Range("C" & r + 1).Value = "STITCH-OH"
        Worksheets("test111").Range("D" & r).Value = 1
        Worksheets("test111").Range("D" & r + 1).Value = 1
      r = r + 2
    Set rng = Worksheets("test111").Range("A" & x & ":A" & r - 1)
    rng.Value = fu & art & WorksheetFunction.Text(s1 + i, "00")
    x = r
Next i

If C_LOOK("CCP") > 0 Then
'PCS
cellX = CELL_X("CCP")
For i = 0 To siz
    Worksheets("test111").Range("A" & r).Value = pcs & art & WorksheetFunction.Text(s1 + i, "00")
    Worksheets("test111").Range("B" & r).Value = 0
    Worksheets("test111").Range("C" & r).Value = ccp & art & WorksheetFunction.Text(s1 + i, "00")
    Worksheets("test111").Range("D" & r).Value = 1
    
    Worksheets("test111").Range("A" & r + 1).Value = pcs & art & WorksheetFunction.Text(s1 + i, "00")
    Worksheets("test111").Range("B" & r + 1).Value = 1
    Worksheets("test111").Range("C" & r + 1).Value = "PRINTING-CHARGES"
    Worksheets("test111").Range("D" & r + 1).Value = 1
    r = r + 2
Next i
'CCP
For i = 0 To siz
n = 0
    For j = 0 To cellX(1) - 1
    If IsEmpty(Worksheets("BOM").Range("D" & cellX(0) + j).Value) = False Then
        Worksheets("test111").Range("B" & r + n).Value = n
        Worksheets("test111").Range("C" & r + n).Value = Worksheets("BOM").Range("D" & cellX(0) + j)
        Worksheets("test111").Range("D" & r + n).Value = Worksheets("BOM").cellS(cellX(0) + j, 6)
        n = n + 1
        End If
    Next j
    
Worksheets("test111").Range("B" & r + n).Value = n
Worksheets("test111").Range("C" & r + n).Value = "CLICK_OH"
Worksheets("test111").Range("D" & r + j).Value = 1

Set rng = Worksheets("test111").Range("A" & r & ":A" & r + n)
rng.Value = ccp & art & WorksheetFunction.Text(s1 + i, "00")
r = r + n + 1
Next i
End If

If C_LOOK("CCP1") > 0 Then
'PCS1
cellX = CELL_X("CCP1")
For i = 0 To siz
    Worksheets("test111").Range("A" & r).Value = pcs1 & art & WorksheetFunction.Text(s1 + i, "00")
    Worksheets("test111").Range("B" & r).Value = 0
    Worksheets("test111").Range("C" & r).Value = ccp1 & art & WorksheetFunction.Text(s1 + i, "00")
    Worksheets("test111").Range("D" & r).Value = 1
    
    Worksheets("test111").Range("A" & r + 1).Value = pcs1 & art & WorksheetFunction.Text(s1 + i, "00")
    Worksheets("test111").Range("B" & r + 1).Value = 1
    Worksheets("test111").Range("C" & r + 1).Value = "PRINTING-CHARGES"
    Worksheets("test111").Range("D" & r + 1).Value = 1
    r = r + 2
Next i

'CCP1
For i = 0 To siz
n = 0
    For j = 0 To cellX(1) - 1
        If IsEmpty(Worksheets("BOM").Range("D" & cellX(0) + j).Value) = False Then
        Worksheets("test111").Range("B" & r + n).Value = n
        Worksheets("test111").Range("C" & r + n).Value = Worksheets("BOM").Range("D" & cellX(0) + j)
        Worksheets("test111").Range("D" & r + n).Value = Worksheets("BOM").cellS(cellX(0) + j, 6)
        n = n + 1
        End If
    Next j
    
Worksheets("test111").Range("B" & r + n).Value = n
Worksheets("test111").Range("C" & r + n).Value = "CLICK_OH"
Worksheets("test111").Range("D" & r + n).Value = 1

Set rng = Worksheets("test111").Range("A" & r & ":A" & r + n)
rng.Value = ccp1 & art & WorksheetFunction.Text(s1 + i, "00")
r = r + n + 1
Next i
End If
If C_LOOK("CCS") > 0 Then
'CCS
cellX = CELL_X("CCS")
For i = 0 To siz
n = 0
    For j = 0 To cellX(1) - 1
    If IsEmpty(Worksheets("BOM").Range("D" & cellX(0) + j).Value) = False Then
        Worksheets("test111").Range("B" & r + n).Value = n
        Worksheets("test111").Range("C" & r + n).Value = Worksheets("BOM").Range("D" & cellX(0) + j)
        Worksheets("test111").Range("D" & r + n).Value = Worksheets("BOM").cellS(cellX(0) + j, 6)
        n = n + 1
        End If
    Next j
    
Worksheets("test111").Range("B" & r + n).Value = n
Worksheets("test111").Range("C" & r + n).Value = "CLICK_OH"
Worksheets("test111").Range("D" & r + n).Value = 1

Set rng = Worksheets("test111").Range("A" & r & ":A" & r + n)
rng.Value = ccs & art & WorksheetFunction.Text(s1 + i, "00")
r = r + n + 1
Next i
End If

'MCS
If C_LOOK("MCS") > 0 Then
cellX = CELL_X("MCS")
For i = 0 To siz
n = 0
    For j = 0 To cellX(1) - 1
         If IsEmpty(Worksheets("BOM").Range("C" & cellX(0) + j).Value) = False Then
        Worksheets("test111").Range("B" & r + n).Value = n
        Worksheets("test111").Range("C" & r + n).Value = "4-" & Worksheets("BOM").Range("C" & cellX(0) + j) & "-" & art
        Worksheets("test111").Range("D" & r + n).Value = Worksheets("BOM").cellS(cellX(0) + j, 6)
        n = n + 1
        End If
    Next j
    
Worksheets("test111").Range("B" & r + n).Value = n
Worksheets("test111").Range("C" & r + n).Value = "MARKING-CHARGES"
Worksheets("test111").Range("D" & r + n).Value = 1

Set rng = Worksheets("test111").Range("A" & r & ":A" & r + n)
rng.Value = mcs & art & WorksheetFunction.Text(s1 + i, "00")
r = r + n + 1
Next i

'FCM
n = 0
    For j = 0 To cellX(1) - 1
         If IsEmpty(Worksheets("BOM").Range("C" & cellX(0) + j).Value) = False Then
         
        Worksheets("test111").Range("A" & r).Value = "4-" & Worksheets("BOM").Range("C" & cellX(0) + j) & "-" & art
        Worksheets("test111").Range("B" & r).Value = 0
        Worksheets("test111").Range("C" & r).Value = scf(n) & art
        Worksheets("test111").Range("D" & r).Value = 1
        
        Worksheets("test111").Range("A" & r + 1).Value = scf(n) & art
        Worksheets("test111").Range("B" & r + 1).Value = 0
        Worksheets("test111").Range("C" & r + 1).Value = Worksheets("BOM").Range("C" & C_LOOK("SCF" & n))
        Worksheets("test111").Range("D" & r + 1).Value = Worksheets("BOM").cellS(C_LOOK("SCF" & n), 6)
        
        Worksheets("test111").Range("A" & r + 2).Value = scf(n) & art
        Worksheets("test111").Range("B" & r + 2).Value = 1
        Worksheets("test111").Range("C" & r + 2).Value = "SLITT-OH"
        Worksheets("test111").Range("D" & r + 2).Value = 1
        n = n + 1
        r = r + 3
        End If
    Next j
End If
If C_LOOK("FCS") > 0 Then
n = 0
cellX = CELL_X("FCS")
'FCS
    For j = 0 To cellX(1) - 1
         If IsEmpty(Worksheets("BOM").Range("C" & cellX(0) + j).Value) = False Then
         
        Worksheets("test111").Range("A" & r).Value = "4-" & Worksheets("BOM").Range("C" & cellX(0) + j) & "-" & art
        Worksheets("test111").Range("B" & r).Value = 0
        Worksheets("test111").Range("C" & r).Value = scf(n) & art
        Worksheets("test111").Range("D" & r).Value = 1
        
        Worksheets("test111").Range("A" & r + 1).Value = scf(n) & art
        Worksheets("test111").Range("B" & r + 1).Value = 0
        Worksheets("test111").Range("C" & r + 1).Value = Worksheets("BOM").Range("C" & C_LOOK("SCF" & n))
        Worksheets("test111").Range("D" & r + 1).Value = Worksheets("BOM").cellS(C_LOOK("SCF" & n), 6)
        
        Worksheets("test111").Range("A" & r + 2).Value = scf(n) & art
        Worksheets("test111").Range("B" & r + 2).Value = 1
        Worksheets("test111").Range("C" & r + 2).Value = "SLITT-OH"
        Worksheets("test111").Range("D" & r + 2).Value = 1
        n = n + 1
        r = r + 3
        End If
    Next j

 End If
       
  If C_LOOK("SCS") > 0 Then
cellX = CELL_X("SCS")
'SCS
n = 0
    For j = 0 To cellX(1) - 1
     If IsEmpty(Worksheets("BOM").Range("C" & cellX(0) + j).Value) = False Then
        Worksheets("test111").Range("A" & r).Value = scf(n) & art
        Worksheets("test111").Range("B" & r).Value = 0
        Worksheets("test111").Range("C" & r).Value = Worksheets("BOM").Range("C" & C_LOOK("SCF" & n))
        Worksheets("test111").Range("D" & r).Value = Worksheets("BOM").cellS(C_LOOK("SCF" & n), 6)
        
        Worksheets("test111").Range("A" & r + 1).Value = scf(n) & art
        Worksheets("test111").Range("B" & r + 1).Value = 1
        Worksheets("test111").Range("C" & r + 1).Value = "SLITT-OH"
        Worksheets("test111").Range("D" & r + 1).Value = 1
        n = n + 1
        r = r + 2
 End If
    Next j

 End If





End Sub

Function SIZE1(siz As String)
Dim RE As Object
Dim MATCH As Object
Set RE = CreateObject("VBScript.regexp")
RE.Pattern = "(\d{1,2})"
RE.Global = True
RE.ignoreCase = False
Set MATCH = RE.Execute(siz)
If MATCH.Count <> 0 Then
result = MATCH.Item(0).submatches.Item(0)
End If
SIZE1 = result
End Function

Function SIZE2(siz As String)
Dim RE As Object
Dim MATCH As Object
Set RE = CreateObject("VBScript.regexp")
RE.Pattern = "(\d{1,2})"
RE.Global = True
RE.ignoreCase = False
Set MATCH = RE.Execute(siz)
If MATCH.Count <> 0 Then
result = MATCH.Item(1).submatches.Item(0)
End If
SIZE2 = result
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

Function C_LOOK(cell_name As String)
Dim cellS As Long
With Worksheets("BOM")
    On Error Resume Next
    cellS = Application.WorksheetFunction.MATCH(cell_name, .Range("B:B"), 0)
    On Error GoTo 0
End With
C_LOOK = cellS
End Function

Function MC_ITEMS(sizee As String) As Integer()
Dim sc_count(5) As Integer
    
    Select Case sizee
        Case "6X10"
        'todo
            sc_count(0) = 3
            sc_count(4) = 3
            sc_count(1) = 6
            sc_count(2) = 6
            sc_count(3) = 6
            sc_count(5) = 24
        Case "5X9"
            sc_count(0) = 7
            sc_count(1) = 7
            sc_count(2) = 7
            sc_count(3) = 7
            sc_count(4) = 2
            sc_count(5) = 30
        Case Else
            sc_count(0) = sc_count(1) = sc_count(2) = sc_count(3) = sc_count(4) = sc_count(5) = 0
    End Select
MC_ITEMS = sc_count()
End Function

