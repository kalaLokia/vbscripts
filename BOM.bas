Attribute VB_Name = "BOM"
Sub BOM_LINE()
'test sample, requires "LINE" sheet to excecute


Dim artNo, artColor, artCat, s2, s1, brandSize As String
Dim siz, c, r, l As Integer
Dim mc, sc, fu, mpu, pcs, ccp, ccp1, ccs, fcmx, mcs, art As String
Dim scf(2) As String
Dim scs(2), fcs(2), fcm(2) As String
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

'Master Carton - MC
    cellX = CELL_X("MC")

    Set rng = Worksheets("LINE").Range("A" & r & ":A" & cellX(1) + siz + r + 1)
    rng.Value = mc & art & 1
    For i = 0 To siz
        Worksheets("LINE").Range("B" & r + i).Value = i
        Worksheets("LINE").Range("C" & r + i).Value = sc & art & WorksheetFunction.Text(s1 + i, "00")
        Worksheets("LINE").Range("D" & r + i).Value = scCount(i)
        Worksheets("LINE").Range("H" & r + i).Value = 4
    Next i
    r = r + i
    For j = 0 To cellX(1) - 1
        If IsEmpty(Worksheets("BOM").Range("D" & cellX(0) + j).Value) = False Then
            Worksheets("LINE").Range("B" & r + n).Value = i + n
            Worksheets("LINE").Range("C" & r + n).Value = Worksheets("BOM").Range("D" & cellX(0) + j)
            Worksheets("LINE").Range("D" & r + n).Value = Worksheets("BOM").cellS(cellX(0) + j, 6)
            Worksheets("LINE").Range("H" & r + n).Value = 4
            n = n + 1
        End If
    Next j
    r = r + n
    Worksheets("LINE").Range("B" & r).Value = i + n
    Worksheets("LINE").Range("C" & r).Value = "FGMC_OH"
    Worksheets("LINE").Range("D" & r).Value = scCount(i)
    Worksheets("LINE").Range("H" & r).Value = 290
    r = r + 1
'Small Carton - SC
    cellX = CELL_X("SC")
    For i = 0 To siz
        n = 0
        Worksheets("LINE").Range("B" & r).Value = n
        Worksheets("LINE").Range("C" & r).Value = mpu & art & WorksheetFunction.Text(s1 + i, "00")
        Worksheets("LINE").Range("D" & r).Value = 1
        Worksheets("LINE").Range("H" & r).Value = 4
    
        For j = 0 To cellX(1) - 1
            If IsEmpty(Worksheets("BOM").Range("D" & cellX(0) + j).Value) = False Then
                n = n + 1
                Worksheets("LINE").Range("B" & r + n).Value = n
                Worksheets("LINE").Range("C" & r + n).Value = Worksheets("BOM").Range("D" & cellX(0) + j)
                Worksheets("LINE").Range("D" & r + n).Value = Worksheets("BOM").cellS(cellX(0) + j, 6)
                Worksheets("LINE").Range("H" & r + n).Value = 4
            End If
        Next j
    
        Worksheets("LINE").Range("B" & r + n + 1).Value = n + 1
        Worksheets("LINE").Range("C" & r + n + 1).Value = "FGSC_OH"
        Worksheets("LINE").Range("D" & r + n + 1).Value = 1
        Worksheets("LINE").Range("H" & r + n + 1).Value = 290

        Set rng = Worksheets("LINE").Range("A" & r & ":A" & cellX(1) + siz + r + 2)
        rng.Value = sc & art & WorksheetFunction.Text(s1 + i, "00")
        r = r + n + 2
    Next i

'MPU
    cellX = CELL_X("SW")
    For i = 0 To siz
        Worksheets("LINE").Range("A" & r).Value = mpu & art & WorksheetFunction.Text(s1 + i, "00")
        Worksheets("LINE").Range("B" & r).Value = 0
        Worksheets("LINE").Range("C" & r).Value = fu & art & WorksheetFunction.Text(s1 + i, "00")
        Worksheets("LINE").Range("D" & r).Value = 1
        Worksheets("LINE").Range("H" & r).Value = 4
    
        Worksheets("LINE").Range("A" & r + 1).Value = mpu & art & WorksheetFunction.Text(s1 + i, "00")
        Worksheets("LINE").Range("B" & r + 1).Value = 1
        Worksheets("LINE").Range("C" & r + 1).Value = "4-PUX-0003"
        Worksheets("LINE").Range("D" & r + 1).Value = Worksheets("BOM").cellS(cellX(0), i + 6)
        Worksheets("LINE").Range("H" & r + 1).Value = 4
        
        Worksheets("LINE").Range("A" & r + 2).Value = mpu & art & WorksheetFunction.Text(s1 + i, "00")
        Worksheets("LINE").Range("B" & r + 2).Value = 2
        Worksheets("LINE").Range("C" & r + 2).Value = "MPU_OH"
        Worksheets("LINE").Range("D" & r + 2).Value = 1
        Worksheets("LINE").Range("H" & r + 2).Value = 290
        r = r + 3
    Next i

'FU
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

'PCS n CCP
    If C_LOOK("CCP", "B") > 0 Then
        cellX = CELL_X("CCP")
        For i = 0 To siz
            Worksheets("LINE").Range("A" & r).Value = pcs & art & WorksheetFunction.Text(s1 + i, "00")
            Worksheets("LINE").Range("B" & r).Value = 0
            Worksheets("LINE").Range("C" & r).Value = ccp & art & WorksheetFunction.Text(s1 + i, "00")
            Worksheets("LINE").Range("D" & r).Value = 1
            Worksheets("LINE").Range("H" & r).Value = 4
    
            Worksheets("LINE").Range("A" & r + 1).Value = pcs & art & WorksheetFunction.Text(s1 + i, "00")
            Worksheets("LINE").Range("B" & r + 1).Value = 1
            Worksheets("LINE").Range("C" & r + 1).Value = "PRINTING-CHARGES"
            Worksheets("LINE").Range("D" & r + 1).Value = 1
            Worksheets("LINE").Range("H" & r + 1).Value = 290
            r = r + 2
        Next i
'CCP
        For i = 0 To siz
            n = 0
            For j = 0 To cellX(1) - 1
                If IsEmpty(Worksheets("BOM").Range("D" & cellX(0) + j).Value) = False Then
                    Worksheets("LINE").Range("B" & r + n).Value = n
                    Worksheets("LINE").Range("C" & r + n).Value = Worksheets("BOM").Range("D" & cellX(0) + j)
                    Worksheets("LINE").Range("D" & r + n).Value = Worksheets("BOM").cellS(cellX(0) + j, 6)
                    Worksheets("LINE").Range("H" & r + n).Value = 4
                    n = n + 1
                End If
            Next j
    
            Worksheets("LINE").Range("B" & r + n).Value = n
            Worksheets("LINE").Range("C" & r + n).Value = "CLICK_OH"
            Worksheets("LINE").Range("D" & r + n).Value = 1
            Worksheets("LINE").Range("H" & r + n).Value = 290

            Set rng = Worksheets("LINE").Range("A" & r & ":A" & r + n)
            rng.Value = ccp & art & WorksheetFunction.Text(s1 + i, "00")
            r = r + n + 1
        Next i
    End If
'PCS1 n CCP1
    If C_LOOK("CCP1", "B") > 0 Then

        cellX = CELL_X("CCP1")
        For i = 0 To siz
            Worksheets("LINE").Range("A" & r).Value = pcs1 & art & WorksheetFunction.Text(s1 + i, "00")
            Worksheets("LINE").Range("B" & r).Value = 0
            Worksheets("LINE").Range("C" & r).Value = ccp1 & art & WorksheetFunction.Text(s1 + i, "00")
            Worksheets("LINE").Range("D" & r).Value = 1
            Worksheets("LINE").Range("H" & r).Value = 4
    
            Worksheets("LINE").Range("A" & r + 1).Value = pcs1 & art & WorksheetFunction.Text(s1 + i, "00")
            Worksheets("LINE").Range("B" & r + 1).Value = 1
            Worksheets("LINE").Range("C" & r + 1).Value = "PRINTING-CHARGES"
            Worksheets("LINE").Range("D" & r + 1).Value = 1
            Worksheets("LINE").Range("H" & r + 1).Value = 290
            r = r + 2
        Next i

        For i = 0 To siz
            n = 0
            For j = 0 To cellX(1) - 1
                If IsEmpty(Worksheets("BOM").Range("D" & cellX(0) + j).Value) = False Then
                    Worksheets("LINE").Range("B" & r + n).Value = n
                    Worksheets("LINE").Range("C" & r + n).Value = Worksheets("BOM").Range("D" & cellX(0) + j)
                    Worksheets("LINE").Range("D" & r + n).Value = Worksheets("BOM").cellS(cellX(0) + j, 6)
                    Worksheets("LINE").Range("H" & r + n).Value = 4
                    n = n + 1
                End If
            Next j
    
            Worksheets("LINE").Range("B" & r + n).Value = n
            Worksheets("LINE").Range("C" & r + n).Value = "CLICK_OH"
            Worksheets("LINE").Range("D" & r + n).Value = 1
            Worksheets("LINE").Range("H" & r + n).Value = 290

            Set rng = Worksheets("LINE").Range("A" & r & ":A" & r + n)
            rng.Value = ccp1 & art & WorksheetFunction.Text(s1 + i, "00")
            r = r + n + 1
        Next i
    End If
    
'CCS
    If C_LOOK("CCS", "B") > 0 Then
        cellX = CELL_X("CCS")
        For i = 0 To siz
            n = 0
            For j = 0 To cellX(1) - 1
                If IsEmpty(Worksheets("BOM").Range("D" & cellX(0) + j).Value) = False Then
                    Worksheets("LINE").Range("B" & r + n).Value = n
                    Worksheets("LINE").Range("C" & r + n).Value = Worksheets("BOM").Range("D" & cellX(0) + j)
                    Worksheets("LINE").Range("D" & r + n).Value = Worksheets("BOM").cellS(cellX(0) + j, 6)
                    Worksheets("LINE").Range("H" & r + n).Value = 4
                    n = n + 1
                End If
            Next j
    
            Worksheets("LINE").Range("B" & r + n).Value = n
            Worksheets("LINE").Range("C" & r + n).Value = "CLICK_OH"
            Worksheets("LINE").Range("D" & r + n).Value = 1
            Worksheets("LINE").Range("H" & r + n).Value = 290

            Set rng = Worksheets("LINE").Range("A" & r & ":A" & r + n)
            rng.Value = ccs & art & WorksheetFunction.Text(s1 + i, "00")
            r = r + n + 1
        Next i
    End If

'MARK
    If C_LOOK("MARK", "B") > 0 Then
        cellX = CELL_X("MARK")
        For i = 0 To siz
            n = 0
            For j = 0 To cellX(1) - 1
                If IsEmpty(Worksheets("BOM").Range("C" & cellX(0) + j).Value) = False Then
                    Worksheets("LINE").Range("B" & r + n).Value = n
                    Worksheets("LINE").Range("C" & r + n).Value = "4-" & Worksheets("BOM").Range("C" & cellX(0) + j) & "-" & art
                    Worksheets("LINE").Range("D" & r + n).Value = Worksheets("BOM").cellS(cellX(0) + j, 6)
                    Worksheets("LINE").Range("H" & r + n).Value = 4
                    n = n + 1
                End If
            Next j
    
            Worksheets("LINE").Range("B" & r + n).Value = n
            Worksheets("LINE").Range("C" & r + n).Value = "MARKING-CHARGES"
            Worksheets("LINE").Range("D" & r + n).Value = 1
            Worksheets("LINE").Range("H" & r + n).Value = 290

            Set rng = Worksheets("LINE").Range("A" & r & ":A" & r + n)
            rng.Value = mcs & art & WorksheetFunction.Text(s1 + i, "00")
            r = r + n + 1
        Next i

'FCM
        If C_LOOK("FCM", "C") > 0 Then
            Worksheets("LINE").Range("A" & r).Value = fcm(0) & art
            Worksheets("LINE").Range("B" & r).Value = 0
            Worksheets("LINE").Range("C" & r).Value = scf(0) & art
            Worksheets("LINE").Range("D" & r).Value = 1
            Worksheets("LINE").Range("H" & r).Value = 4
        
            Worksheets("LINE").Range("A" & r + 1).Value = scf(0) & art
            Worksheets("LINE").Range("B" & r + 1).Value = 0
            Worksheets("LINE").Range("C" & r + 1).Value = Worksheets("BOM").Range("D" & C_LOOK("FCM", "C"))
            Worksheets("LINE").Range("D" & r + 1).Value = Worksheets("BOM").cellS(C_LOOK("FCM", "C"), 11)
            Worksheets("LINE").Range("H" & r + 1).Value = 4
        
            Worksheets("LINE").Range("A" & r + 2).Value = scf(0) & art
            Worksheets("LINE").Range("B" & r + 2).Value = 1
            Worksheets("LINE").Range("C" & r + 2).Value = "SLITT-OH"
            Worksheets("LINE").Range("D" & r + 2).Value = 1
            Worksheets("LINE").Range("H" & r + 2).Value = 290
    
            r = r + 3
        End If
    
'FCM1
        If C_LOOK("FCM1", "C") > 0 Then
            Worksheets("LINE").Range("A" & r).Value = fcm(1) & art
            Worksheets("LINE").Range("B" & r).Value = 0
            Worksheets("LINE").Range("C" & r).Value = scf(1) & art
            Worksheets("LINE").Range("D" & r).Value = 1
            Worksheets("LINE").Range("H" & r).Value = 4
        
            Worksheets("LINE").Range("A" & r + 1).Value = scf(1) & art
            Worksheets("LINE").Range("B" & r + 1).Value = 0
            Worksheets("LINE").Range("C" & r + 1).Value = Worksheets("BOM").Range("D" & C_LOOK("FCM1", "C"))
            Worksheets("LINE").Range("D" & r + 1).Value = Worksheets("BOM").cellS(C_LOOK("FCM1", "C"), 11)
            Worksheets("LINE").Range("H" & r + 1).Value = 4
        
            Worksheets("LINE").Range("A" & r + 2).Value = scf(1) & art
            Worksheets("LINE").Range("B" & r + 2).Value = 1
            Worksheets("LINE").Range("C" & r + 2).Value = "SLITT-OH"
            Worksheets("LINE").Range("D" & r + 2).Value = 1
            Worksheets("LINE").Range("H" & r + 2).Value = 290
    
            r = r + 3
        End If
    End If
    
'FOLD
    If C_LOOK("FOLD", "B") > 0 Then
        If C_LOOK("FCS", "C") > 0 Then
            Worksheets("LINE").Range("A" & r).Value = fcs(0) & art
            Worksheets("LINE").Range("B" & r).Value = 0
            Worksheets("LINE").Range("C" & r).Value = scf(0) & art
            Worksheets("LINE").Range("D" & r).Value = 1
            Worksheets("LINE").Range("H" & r).Value = 4
        
            Worksheets("LINE").Range("A" & r + 1).Value = scf(0) & art
            Worksheets("LINE").Range("B" & r + 1).Value = 0
            Worksheets("LINE").Range("C" & r + 1).Value = Worksheets("BOM").Range("D" & C_LOOK("FCS", "C"))
            Worksheets("LINE").Range("D" & r + 1).Value = Worksheets("BOM").cellS(C_LOOK("FCS", "C"), 11)
            Worksheets("LINE").Range("H" & r + 1).Value = 4
        
            Worksheets("LINE").Range("A" & r + 2).Value = scf(0) & art
            Worksheets("LINE").Range("B" & r + 2).Value = 1
            Worksheets("LINE").Range("C" & r + 2).Value = "SLITT-OH"
            Worksheets("LINE").Range("D" & r + 2).Value = 1
            Worksheets("LINE").Range("H" & r + 2).Value = 290
    
            r = r + 3
        End If

        If C_LOOK("FCS1", "C") > 0 Then
            Worksheets("LINE").Range("A" & r).Value = fcs(1) & art
            Worksheets("LINE").Range("B" & r).Value = 0
            Worksheets("LINE").Range("C" & r).Value = scf(1) & art
            Worksheets("LINE").Range("D" & r).Value = 1
            Worksheets("LINE").Range("H" & r).Value = 4
        
            Worksheets("LINE").Range("A" & r + 1).Value = scf(1) & art
            Worksheets("LINE").Range("B" & r + 1).Value = 0
            Worksheets("LINE").Range("C" & r + 1).Value = Worksheets("BOM").Range("D" & C_LOOK("FCS1", "C"))
            Worksheets("LINE").Range("D" & r + 1).Value = Worksheets("BOM").cellS(C_LOOK("FCS1", "C"), 11)
            Worksheets("LINE").Range("H" & r + 1).Value = 4
        
            Worksheets("LINE").Range("A" & r + 2).Value = scf(1) & art
            Worksheets("LINE").Range("B" & r + 2).Value = 1
            Worksheets("LINE").Range("C" & r + 2).Value = "SLITT-OH"
            Worksheets("LINE").Range("D" & r + 2).Value = 1
            Worksheets("LINE").Range("H" & r + 2).Value = 290
    
            r = r + 3
        End If
    End If

'SLIT
    If C_LOOK("SLIT", "B") > 0 Then

        If C_LOOK("SCS", "C") > 0 Then
        
            Worksheets("LINE").Range("A" & r).Value = scs(0) & art
            Worksheets("LINE").Range("B" & r).Value = 0
            Worksheets("LINE").Range("C" & r).Value = Worksheets("BOM").Range("D" & C_LOOK("SCS", "C"))
            Worksheets("LINE").Range("D" & r).Value = Worksheets("BOM").cellS(C_LOOK("SCS", "C"), 11)
            Worksheets("LINE").Range("H" & r).Value = 4
            
            Worksheets("LINE").Range("A" & r + 1).Value = scs(0) & art
            Worksheets("LINE").Range("B" & r + 1).Value = 1
            Worksheets("LINE").Range("C" & r + 1).Value = "SLITT-OH"
            Worksheets("LINE").Range("D" & r + 1).Value = 1
            Worksheets("LINE").Range("H" & r + 1).Value = 290
    
            r = r + 2
        End If
        
        If C_LOOK("SCS1", "C") > 0 Then
        
            Worksheets("LINE").Range("A" & r).Value = scs(1) & art
            Worksheets("LINE").Range("B" & r).Value = 0
            Worksheets("LINE").Range("C" & r).Value = Worksheets("BOM").Range("D" & C_LOOK("SCS1", "C"))
            Worksheets("LINE").Range("D" & r).Value = Worksheets("BOM").cellS(C_LOOK("SCS1", "C"), 11)
            Worksheets("LINE").Range("H" & r).Value = 4
            
            Worksheets("LINE").Range("A" & r + 1).Value = scs(1) & art
            Worksheets("LINE").Range("B" & r + 1).Value = 1
            Worksheets("LINE").Range("C" & r + 1).Value = "SLITT-OH"
            Worksheets("LINE").Range("D" & r + 1).Value = 1
            Worksheets("LINE").Range("H" & r + 1).Value = 290
    
            r = r + 2
        End If
        If C_LOOK("SCS2", "C") > 0 Then
        
            Worksheets("LINE").Range("A" & r).Value = scs(2) & art
            Worksheets("LINE").Range("B" & r).Value = 0
            Worksheets("LINE").Range("C" & r).Value = Worksheets("BOM").Range("D" & C_LOOK("SCS2", "C"))
            Worksheets("LINE").Range("D" & r).Value = Worksheets("BOM").cellS(C_LOOK("SCS2", "C"), 11)
            Worksheets("LINE").Range("H" & r).Value = 4
            
            Worksheets("LINE").Range("A" & r + 1).Value = scs(2) & art
            Worksheets("LINE").Range("B" & r + 1).Value = 1
            Worksheets("LINE").Range("C" & r + 1).Value = "SLITT-OH"
            Worksheets("LINE").Range("D" & r + 1).Value = 1
            Worksheets("LINE").Range("H" & r + 1).Value = 290
    
            r = r + 2
        End If
    End If


'Other constants to rest range
    Set rng = Worksheets("LINE").Range("E3:E" & r - 1)
    rng.Value = "FB/PF001"
    Set rng = Worksheets("LINE").Range("F3:F" & r - 1)
    rng.Value = "INR"
    Set rng = Worksheets("LINE").Range("G3:G" & r - 1)
    rng.Value = "B"
    Worksheets("LINE").Range("A1").Value = "ParentKey"
    Worksheets("LINE").Range("A2").Value = "TreeCode"
    Worksheets("LINE").Range("B1").Value = "LineNum"
    Worksheets("LINE").Range("B2").Value = "LineNum"
    Worksheets("LINE").Range("C1").Value = "ItemCode"
    Worksheets("LINE").Range("C2").Value = "Code"
    Worksheets("LINE").Range("D1").Value = "Quantity"
    Worksheets("LINE").Range("D2").Value = "Quantity"
    Worksheets("LINE").Range("E1").Value = "Warehouse"
    Worksheets("LINE").Range("E2").Value = "Warehouse"
    Worksheets("LINE").Range("F1").Value = "Currency"
    Worksheets("LINE").Range("F2").Value = "Currency"
    Worksheets("LINE").Range("G1").Value = "IssueMethod"
    Worksheets("LINE").Range("G2").Value = "IssueMthd"
    Worksheets("LINE").Range("H1").Value = "Type"
    Worksheets("LINE").Range("H2").Value = "Type"


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
            Case "1X3"
                sc_count(0) = 10
                sc_count(1) = 10
                sc_count(2) = 10
                sc_count(3) = 30
                sc_count(4) = 0
                sc_count(5) = 0
            Case "1X5"
                sc_count(0) = 6
                sc_count(1) = 6
                sc_count(2) = 6
                sc_count(3) = 6
                sc_count(4) = 6
                sc_count(5) = 30
            Case "11X13"
                sc_count(0) = 12
                sc_count(1) = 12
                sc_count(2) = 12
                sc_count(3) = 36
                sc_count(4) = 0
                sc_count(5) = 0
            Case "8X10"
                sc_count(0) = 12
                sc_count(1) = 12
                sc_count(2) = 12
                sc_count(3) = 36
                sc_count(4) = 0
                sc_count(5) = 0
            Case "11X12"
                sc_count(0) = 6
                sc_count(1) = 6
                sc_count(2) = 12
                sc_count(3) = 0
                sc_count(4) = 0
                sc_count(5) = 0
            Case Else
                sc_count(0) = 0
                sc_count(1) = 0
                sc_count(2) = 0
                sc_count(3) = 0
                sc_count(4) = 0
                sc_count(5) = 0
        End Select
    MC_ITEMS = sc_count()
End Function
