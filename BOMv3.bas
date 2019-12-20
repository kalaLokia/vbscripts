Attribute VB_Name = "BOMv3"
'Version 0.3
'A Product from, FORTUNE ELASTOMERS BRANCH, KINALOOR #OM-DEPT
'Created by kalaLokia #4442   ;-)
'DISCLAIMER: USE IT ON YOUR OWN RISK, DO NOT BLAME ON US �\_(?)_/�

Public artSize() As Long
Public siz, row_count, itemCount, scCount() As Integer
Public article, brandSize As String
Public cellX() As Long

Sub BOM()
'TEST SAMPLE, requires "BOM", "LINE" & "TREE" sheets to excecute

    Dim artNo, artColor, artCat As String
    row_count = 3
    
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
'SC
     If C_LOOK("SC", "B") > 0 Then
        SMALL_CARTON "3-fb-" & article
    End If
'MPU
     If C_LOOK("MPU", "B") > 0 Then
        MOULDED_PU "4-mpu-" & article
    End If
'FU
     If C_LOOK("FU", "B") > 0 Then
        FINISHED_UPPER "4-fu-" & article
    End If
'PCS
    If C_LOOK("CCP", "B") > 0 Or C_LOOK("P", "A") > 0 Then
        PRINTING_UPPER "4-pcs-" & article
    End If
'PCS1
    If C_LOOK("CCP1", "B") > 0 Or C_LOOK("P1", "A") > 0 Then
        PRINTING_UPPER1 "4-pcs1-" & article
    End If
'MCS
    If C_LOOK("FCM", "B") > 0 Or C_LOOK("M", "A") > 0 Then
        MARKING_UPPER "4-mcs-" & article
    End If
'CCP
    If C_LOOK("CCP", "B") > 0 Then
        CLICKING_UPPER "4-ccp-" & article
    End If
'CCP
    If C_LOOK("CCP1", "B") > 0 Then
        CLICKING_UPPER "4-ccp1-" & article
    End If
'CCS
    If C_LOOK("CCS", "B") > 0 Then
        CLICKING_UPPER "4-ccs-" & article
    End If
'FCM
    If C_LOOK("FCM", "B") > 0 Then
        FOLDED_UPPER "4-fcm-" & article
    End If
'FCM1
    If C_LOOK("FCM1", "B") > 0 Then
        FOLDED_UPPER "4-fcm1-" & article
    End If
'FCM2
    If C_LOOK("FCM2", "B") > 0 Then
        FOLDED_UPPER "4-fcm2-" & article
    End If
'FCS
    If C_LOOK("FCS", "B") > 0 Then
        FOLDED_UPPER "4-fcs-" & article
    End If
'FCS1
    If C_LOOK("FCS1", "B") > 0 Then
        FOLDED_UPPER "4-fcs1-" & article
    End If
'FCS2
    If C_LOOK("FCS2", "B") > 0 Then
        FOLDED_UPPER "4-fcs-" & article
    End If
'SCS
    If C_LOOK("SCS", "B") > 0 Then
        SLITTED_UPPER "4-scs-" & article
    End If
'SCS1
    If C_LOOK("SCS1", "B") > 0 Then
        SLITTED_UPPER "4-scs1-" & article
    End If
'SCS2
    If C_LOOK("SCS2", "B") > 0 Then
        SLITTED_UPPER "4-scs2-" & article
    End If
    
LINE_TREE_TITLES
BUILD_TREE
    

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

'Finished Upper - FU
Sub FINISHED_UPPER(ite As String)
   cellX = CELL_X("FU")
    
    For i = 0 To siz
    itemCount = 0
    If C_LOOK("CCP", "B") > 0 Then
        If UCase(Worksheets("BOM").Range("A" & C_LOOK("CCP", "B"))) <> "M" Then
            LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, "4-pcs-" & article & WorksheetFunction.Text(artSize(0) + i, "00"), 1, 4
        End If
    End If
    If C_LOOK("CCP1", "B") > 0 Then
        If UCase(Worksheets("BOM").Range("A" & C_LOOK("CCP1", "B"))) <> "M" Then
            LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, "4-pcs1-" & article & WorksheetFunction.Text(artSize(0) + i, "00"), 1, 4
        End If
    End If
      
    If C_LOOK("CCS", "B") > 0 Then
        LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, "4-ccs-" & article & WorksheetFunction.Text(artSize(0) + i, "00"), 1, 4
    End If
    If C_LOOK("FCM", "B") > 0 Then
        If InStr(1, Worksheets("BOM").Range("A" & C_LOOK("FCM", "B")), "P", vbTextCompare) = 0 Then
            LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, "4-mcs-" & article & WorksheetFunction.Text(artSize(0) + i, "00"), 1, 4
        
        End If
    End If
    If C_LOOK("FCS", "B") > 0 Then
         If InStr(1, Worksheets("BOM").Range("A" & C_LOOK("FCS", "B")), "P", vbTextCompare) = 0 Then
            LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, "4-fcs-" & article, Worksheets("BOM").cellS(C_LOOK("FCS", "B"), i + 6), 4
        End If
    End If
    If C_LOOK("FCS1", "B") > 0 Then
        If InStr(1, Worksheets("BOM").Range("A" & C_LOOK("FCS1", "B")), "P", vbTextCompare) = 0 Then
            LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, "4-fcs1-" & article, Worksheets("BOM").cellS(C_LOOK("FCS1", "B"), i + 6), 4
        End If
    End If
    If C_LOOK("FCS2", "B") > 0 Then
        If InStr(1, Worksheets("BOM").Range("A" & C_LOOK("FCS2", "B")), "P", vbTextCompare) = 0 Then
            LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, "4-fcs2-" & article, Worksheets("BOM").cellS(C_LOOK("FCS2", "B"), i + 6), 4
        End If
    End If
    If C_LOOK("SCS", "B") > 0 Then
        If InStr(1, Worksheets("BOM").Range("A" & C_LOOK("SCS", "B")), "P", vbTextCompare) = 0 Then
            LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, "4-scs-" & article, Worksheets("BOM").cellS(C_LOOK("SCS", "B"), i + 6), 4
        End If
    End If
    If C_LOOK("SCS1", "B") > 0 Then
        If InStr(1, Worksheets("BOM").Range("A" & C_LOOK("SCS1", "B")), "P", vbTextCompare) = 0 Then
            LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, "4-scs1-" & article, Worksheets("BOM").cellS(C_LOOK("SCS1", "B"), i + 6), 4
        End If
    End If
    If C_LOOK("SCS2", "B") > 0 Then
        If InStr(1, Worksheets("BOM").Range("A" & C_LOOK("SCS2", "B")), "P", vbTextCompare) = 0 Then
            LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, "4-scs2-" & article, Worksheets("BOM").cellS(C_LOOK("SCS2", "B"), i + 6), 4
        End If
    End If
            
        For j = 0 To cellX(1) - 1
            If IsEmpty(Worksheets("BOM").Range("D" & cellX(0) + j).Value) = False Then
                LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, Worksheets("BOM").Range("D" & cellX(0) + j), Worksheets("BOM").cellS(cellX(0) + j, i + 6), 4
            End If
        Next j
        LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, "STITCHING-CHARGES", 1, 290
        LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, "STITCH-OH", 1, 290
          
        row_count = row_count + itemCount
    Next i

End Sub

'Printing Upper - PCS
Sub PRINTING_UPPER(ite As String)
    
    For i = 0 To siz
    itemCount = 0
    If C_LOOK("CCP", "B") > 0 Then
        If UCase(Worksheets("BOM").Range("A" & C_LOOK("CCP", "B"))) <> "M" Then
            LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, "4-ccp-" & article & WorksheetFunction.Text(artSize(0) + i, "00"), 1, 4
        End If
    End If
     If C_LOOK("FCM", "B") > 0 Then
        If UCase(Worksheets("BOM").Range("A" & C_LOOK("FCM", "B"))) = "P" Then
            LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, "4-mcs-" & article & WorksheetFunction.Text(artSize(0) + i, "00"), 1, 4
        End If
    End If
    If C_LOOK("FCS", "B") > 0 Then
        If UCase(Worksheets("BOM").Range("A" & C_LOOK("FCS", "B"))) = "P" Then
            LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, "4-fcs-" & article, Worksheets("BOM").cellS(C_LOOK("FCS", "B"), i + 6), 4
        End If
    End If
    If C_LOOK("FCS1", "B") > 0 Then
        If UCase(Worksheets("BOM").Range("A" & C_LOOK("FCS1", "B"))) = "P" Then
            LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, "4-fcs1-" & article, Worksheets("BOM").cellS(C_LOOK("FCS1", "B"), i + 6), 4
        End If
    End If
    If C_LOOK("FCS2", "B") > 0 Then
        If UCase(Worksheets("BOM").Range("A" & C_LOOK("FCS2", "B"))) = "P" Then
            LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, "4-fcs2-" & article, Worksheets("BOM").cellS(C_LOOK("FCS2", "B"), i + 6), 4
        End If
    End If
    If C_LOOK("SCS", "B") > 0 Then
        If UCase(Worksheets("BOM").Range("A" & C_LOOK("SCS", "B"))) = "P" Then
            LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, "4-scs-" & article, Worksheets("BOM").cellS(C_LOOK("SCS", "B"), i + 6), 4
        End If
    End If
    If C_LOOK("SCS1", "B") > 0 Then
        If UCase(Worksheets("BOM").Range("A" & C_LOOK("SCS1", "B"))) = "P" Then
            LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, "4-scs1-" & article, Worksheets("BOM").cellS(C_LOOK("SCS1", "B"), i + 6), 4
        End If
    End If
    If C_LOOK("SCS2", "B") > 0 Then
        If UCase(Worksheets("BOM").Range("A" & C_LOOK("SCS2", "B"))) = "P" Then
            LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, "4-scs2-" & article, Worksheets("BOM").cellS(C_LOOK("SCS2", "B"), i + 6), 4
        End If
    End If
    If itemCount > 0 Then
        LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, "PRINTING-CHARGES", 1, 290
    End If
     row_count = row_count + itemCount
    Next i
    
End Sub

'Printing Upper1 - PCS1
Sub PRINTING_UPPER1(ite As String)
    
    For i = 0 To siz
        itemCount = 0
        If C_LOOK("CCP1", "B") > 0 Then
            If UCase(Worksheets("BOM").Range("A" & C_LOOK("CCP1", "B"))) <> "M" Then
                LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, "4-ccp1-" & article & WorksheetFunction.Text(artSize(0) + i, "00"), 1, 4
            End If
        End If
         If C_LOOK("FCM", "B") > 0 Then
            If UCase(Worksheets("BOM").Range("A" & C_LOOK("FCM", "B"))) = "P1" Then
                LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, "4-mcs-" & article & WorksheetFunction.Text(artSize(0) + i, "00"), 1, 4
            End If
        End If
        If C_LOOK("FCS", "B") > 0 Then
            If UCase(Worksheets("BOM").Range("A" & C_LOOK("FCS", "B"))) = "P1" Then
                LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, "4-fcs-" & article, Worksheets("BOM").cellS(C_LOOK("FCS", "B"), i + 6), 4
            End If
        End If
        If C_LOOK("FCS1", "B") > 0 Then
            If UCase(Worksheets("BOM").Range("A" & C_LOOK("FCS1", "B"))) = "P1" Then
                LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, "4-fcs1-" & article, Worksheets("BOM").cellS(C_LOOK("FCS1", "B"), i + 6), 4
            End If
        End If
        If C_LOOK("FCS2", "B") > 0 Then
            If UCase(Worksheets("BOM").Range("A" & C_LOOK("FCS2", "B"))) = "P1" Then
                LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, "4-fcs2-" & article, Worksheets("BOM").cellS(C_LOOK("FCS2", "B"), i + 6), 4
            End If
        End If
        If C_LOOK("SCS", "B") > 0 Then
            If UCase(Worksheets("BOM").Range("A" & C_LOOK("SCS", "B"))) = "P1" Then
                LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, "4-scs-" & article, Worksheets("BOM").cellS(C_LOOK("SCS", "B"), i + 6), 4
            End If
        End If
        If C_LOOK("SCS1", "B") > 0 Then
            If UCase(Worksheets("BOM").Range("A" & C_LOOK("SCS1", "B"))) = "P1" Then
                LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, "4-scs1-" & article, Worksheets("BOM").cellS(C_LOOK("SCS1", "B"), i + 6), 4
            End If
        End If
        If C_LOOK("SCS2", "B") > 0 Then
            If UCase(Worksheets("BOM").Range("A" & C_LOOK("SCS2", "B"))) = "P1" Then
                LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, "4-scs2-" & article, Worksheets("BOM").cellS(C_LOOK("SCS2", "B"), i + 6), 4
            End If
        End If
        If itemCount > 0 Then
            LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, "PRINTING-CHARGES", 1, 290
        End If
        row_count = row_count + itemCount
    Next i
    
End Sub


'Marking Upper - MCS
Sub MARKING_UPPER(ite As String)
    For i = 0 To siz
        itemCount = 0
        If C_LOOK("FCM", "B") > 0 Then
            If UCase(Worksheets("BOM").Range("A" & C_LOOK("FCM", "B"))) <> "P" Then
                LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, "4-fcm-" & article, Worksheets("BOM").cellS(C_LOOK("FCM", "B"), i + 6), 4
            End If
        End If
        If C_LOOK("FCM1", "B") > 0 Then
            If UCase(Worksheets("BOM").Range("A" & C_LOOK("FCM", "B"))) <> "P" Then
                LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, "4-fcm1-" & article, Worksheets("BOM").cellS(C_LOOK("FCM1", "B"), i + 6), 4
            End If
        End If
        If C_LOOK("FCM2", "B") > 0 Then
            If UCase(Worksheets("BOM").Range("A" & C_LOOK("FCM", "B"))) <> "P" Then
                LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, "4-fcm2-" & article, Worksheets("BOM").cellS(C_LOOK("FCM2", "B"), i + 6), 4
            End If
        End If
        If C_LOOK("CCP", "B") > 0 Then
            If UCase(Worksheets("BOM").Range("A" & C_LOOK("CCP", "B"))) = "M" Then
                LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, "4-pcs-" & article & WorksheetFunction.Text(artSize(0) + i, "00"), 1, 4
            End If
        End If
        If C_LOOK("CCP1", "B") > 0 Then
            If UCase(Worksheets("BOM").Range("A" & C_LOOK("CCP1", "B"))) = "M" Then
                LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, "4-pcs1-" & article & WorksheetFunction.Text(artSize(0) + i, "00"), 1, 4
            End If
        End If
        If itemCount > 0 Then
            LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, "MARKING-CHARGES", 1, 290
        End If
        row_count = row_count + itemCount
    Next i
End Sub

'Clicked component #Printing|Stitching - CCP|CCS
Sub CLICKING_UPPER(ite As String)
    Select Case ite
        Case "4-ccp-" & article
            cellX = CELL_X("CCP")
        Case "4-ccp1-" & article
            cellX = CELL_X("CCP1")
        Case Else
            cellX = CELL_X("CCS")
    End Select
    For i = 0 To siz
        itemCount = 0
        For j = 0 To cellX(1) - 1
            If IsEmpty(Worksheets("BOM").Range("D" & cellX(0) + j).Value) = False Then
                LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, Worksheets("BOM").Range("D" & cellX(0) + j), Worksheets("BOM").cellS(cellX(0) + j, 6 + i), 4
            End If
        Next j
        If itemCount > 0 Then
            LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, "CLICK-OH", 1, 290
            row_count = row_count + itemCount
        End If
    Next i
End Sub

'Folded component #Marking|Folding|Slitting - FCM|FCS
Sub FOLDED_UPPER(ite As String)
    itemCount = 0
    Dim slit As String
    Select Case ite
        Case "4-fcm-" & article
            cellX = CELL_X("FCM")
            slit = "4-scf-"
        Case "4-fcm1-" & article
            cellX = CELL_X("FCM1")
            slit = "4-scf2-"
        Case "4-fcm2-" & article
            cellX = CELL_X("FCM2")
            slit = "4-scf2-"
        Case "4-fcs-" & article
            cellX = CELL_X("FCS")
            slit = "4-scf-"
        Case "4-fcs1-" & article
            cellX = CELL_X("FCS1")
            slit = "4-scf1-"
        Case "4-fcs2-" & article
            cellX = CELL_X("FCS2")
            slit = "4-scf2-"
        Case Else
            MsgBox "Folding is undefined"
            Return
    End Select
    LINE_CELLS ite, itemCount, slit & article, 1, 4
    If cellX(1) > 1 Then
        If IsEmpty(Worksheets("BOM").Range("D" & cellX(0) + 1)) = False Then
            LINE_CELLS ite, itemCount, Worksheets("BOM").Range("D" & cellX(0) + 1), Worksheets("BOM").cellS(cellX(0) + 1, 7 + siz), 4
        End If
    End If
    row_count = row_count + itemCount
    itemCount = 0
     LINE_CELLS slit & article, itemCount, Worksheets("BOM").Range("D" & cellX(0)), Worksheets("BOM").cellS(cellX(0), 7 + siz), 4
     LINE_CELLS slit & article, itemCount, "SLITT-OH", 1, 290
     row_count = row_count + itemCount
End Sub

'Slitted component SCS
Sub SLITTED_UPPER(ite As String)
    itemCount = 0
    Select Case ite
        Case "4-scs-" & article
            cellX = CELL_X("SCS")
        Case "4-scs1-" & article
            cellX = CELL_X("SCS1")
        Case Else
            cellX = CELL_X("SCS2")
    End Select
    LINE_CELLS ite, itemCount, Worksheets("BOM").Range("D" & cellX(0)), Worksheets("BOM").cellS(cellX(0), 7 + siz), 4
    LINE_CELLS ite, itemCount, "SLITT-OH", 1, 290
    
End Sub

Sub LINE_CELLS(valueA, valueB, valueC, valueD, valueH As String)
    Worksheets("LINE").Range("A" & row_count + itemCount).Value = UCase(valueA)
    Worksheets("LINE").Range("B" & row_count + itemCount).Value = UCase(valueB)
    Worksheets("LINE").Range("C" & row_count + itemCount).Value = UCase(valueC)
    Worksheets("LINE").Range("D" & row_count + itemCount).Value = UCase(valueD)
    Worksheets("LINE").Range("E" & row_count + itemCount).Value = "FB/PF001"
    Worksheets("LINE").Range("F" & row_count + itemCount).Value = "INR"
    Worksheets("LINE").Range("G" & row_count + itemCount).Value = "B"
    Worksheets("LINE").Range("H" & row_count + itemCount).Value = UCase(valueH)
    itemCount = itemCount + 1
End Sub

Sub TREE_CELLS(valueA, valueB As String, rc As Long)
    Worksheets("TREE").Range("B" & rc).Value = "P"
    Worksheets("TREE").Range("C" & rc).Value = 1
    Worksheets("TREE").Range("D" & rc).Value = "FB/PF001"
    Worksheets("TREE").Range("E" & rc).Value = UCase(valueA)
    Worksheets("TREE").Range("F" & rc).Value = UCase(valueB)
    itemCount = itemCount + 1
End Sub

Sub LINE_TREE_TITLES()
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
    Worksheets("TREE").Range("A1").Value = "TreeCode"
    Worksheets("TREE").Range("B1").Value = "TreeType"
    Worksheets("TREE").Range("C1").Value = "Quantity"
    Worksheets("TREE").Range("D1").Value = "Warehouse"
    Worksheets("TREE").Range("E1").Value = "U_ProcessType"
    Worksheets("TREE").Range("F1").Value = "U_OhApp"
    Worksheets("TREE").Range("A2").Value = "Code"
    Worksheets("TREE").Range("B2").Value = "TreeType"
    Worksheets("TREE").Range("C2").Value = "Qauntity"
    Worksheets("TREE").Range("D2").Value = "ToWH"
    Worksheets("TREE").Range("E2").Value = "U_ProcessType"
    Worksheets("TREE").Range("F2").Value = "U_OhApp"
End Sub

Sub BUILD_TREE()
    'Requires a sheet with name "TREE"
Dim lastRow, r As Long
r = 2
lastRow = Worksheets("LINE").cellS(Rows.Count, 1).End(xlUp).Row
For i = 3 To lastRow
    If Worksheets("LINE").Range("A" & i) <> Worksheets("TREE").Range("A" & r) Then
        r = r + 1
        Worksheets("TREE").Range("A" & r).Value = Worksheets("LINE").Range("A" & i)
        
        If InStr(1, Worksheets("TREE").Range("A" & r), "2-FB-", vbTextCompare) > 0 Then
           TREE_CELLS "FG_MC", "Y", r
        ElseIf InStr(1, Worksheets("TREE").Range("A" & r), "3-FB-", vbTextCompare) > 0 Then
            TREE_CELLS "FG_SC", "Y", r
         ElseIf InStr(1, Worksheets("TREE").Range("A" & r), "4-MPU-", vbTextCompare) > 0 Then
           TREE_CELLS "MPU", "Y", r
        ElseIf InStr(1, Worksheets("TREE").Range("A" & r), "4-FU-", vbTextCompare) > 0 Then
            TREE_CELLS "STITCH", "Y", r
         ElseIf InStr(1, Worksheets("TREE").Range("A" & r), "4-PCS", vbTextCompare) > 0 Then
            TREE_CELLS "PRINT", "Y", r
        ElseIf InStr(1, Worksheets("TREE").Range("A" & r), "4-CC", vbTextCompare) > 0 Then
            TREE_CELLS "CLICK", "Y", r
        ElseIf InStr(1, Worksheets("TREE").Range("A" & r), "4-MCS-", vbTextCompare) > 0 Then
            TREE_CELLS "MARK", "Y", r
        ElseIf InStr(1, Worksheets("TREE").Range("A" & r), "4-FC", vbTextCompare) > 0 Then
            TREE_CELLS "FOLD", "N", r
        ElseIf InStr(1, Worksheets("TREE").Range("A" & r), "4-SC", vbTextCompare) > 0 Then
            TREE_CELLS "SLIT", "Y", r
        End If
    End If
Next i

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

