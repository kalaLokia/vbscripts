Attribute VB_Name = "BOM"
'Version 0.7.260820 blaze : supports both shoes and slippers
'A MACRO from, FORTUNE ELASTOMERS BRANCH, KINALOOR #OM-DEPT
'Created by kalaLokia #4442   ;-)
'DISCLAIMER: USE IT ON YOUR OWN RISK, DO NOT BLAME ON US ?\_(?)_/?

Public artSize() As Long
Public siz, row_count, itemCount, rwc, scCount() As Integer
Public article, brandSize As String
Public cellX() As Long

Sub BOM()
'TEST SAMPLE, requires "BOM", "LINE" & "TREE" sheets to excecute

'CLEARS ALL DATA IN SHEETS LINE & TREE
    Worksheets("LINE").cellS.Clear
    Worksheets("TREE").cellS.Clear
    
    Dim artNo, artColor, artCat, artBrand As String
    row_count = 3
    
    artNo = Worksheets("BOM").Range("D3")
    artColor = Worksheets("BOM").Range("D4")
    artCat = Worksheets("BOM").Range("D5")
    artBrand = Worksheets("BOM").Range("D2")
    artSize = SIZE_DECODE(Worksheets("BOM").Range("D7"))
    siz = artSize(1) - artSize(0)
    article = artNo & "-" & artColor & "-" & artCat
    If InStr(1, artNo, "Z", vbTextCompare) > 0 Then
        brandSize = UCase(Worksheets("BOM").Range("D7") & "Z")
    Else
        brandSize = UCase(Worksheets("BOM").Range("D7"))
    End If
    If Worksheets("BOM").Range("D6").Value = "SHOES" Then
        brandSize = brandSize & "_SHOES"
    End If
    If artBrand = "SMARTAK" Then
        brandSize = artCat & brandSize & "_SMART"
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
     If C_LOOK("FU", "B") > 0 Or C_LOOK("CCP", "B") > 0 Then
        FINISHED_UPPER "4-fu-" & article
    End If
'PCM
    If C_LOOK("PCM", "B") > 0 Then
        PRINTING_MPUCOM "4-pcm-" & article
    End If
'PCS
    If C_LOOK("CCP", "B") > 0 Or C_LOOK("P", "A") > 0 Then
        PRINTING_UPPER "4-pcs-" & article
    End If
'PCS1
    If C_LOOK("CCP1", "B") > 0 Or C_LOOK("P1", "A") > 0 Or C_LOOK("PCS1", "B") > 0 Then
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
'CCP1
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
        FOLDED_UPPER "4-fcs2-" & article
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
    
LINE_TREE_TITLES 666
BUILD_TREE 666
    

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

'Moulded PU
Sub MOULDED_PU(ite As String)

    cellX = CELL_X("MPU")
    cllX = CELL_X("PCM")
    For i = 0 To siz
        itemCount = 0
         If C_LOOK("FU", "B") > 0 Or C_LOOK("CCP", "B") > 0 Then
            LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, "4-fu-" & article & WorksheetFunction.Text(artSize(0) + i, "00"), 1, 4
         End If
        For j = 0 To cllX(1) - 1
            If C_LOOK("PCM", "B") > 0 And IsEmpty(Worksheets("BOM").cellS(cllX(0) + j, 6 + i).Value) = False Then
                LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, "4-pcm-" & article & WorksheetFunction.Text(artSize(0) + i, "00"), 1, 4
                Exit For
            End If
        Next j
        For j = 0 To cellX(1) - 1
            
            If IsEmpty(Worksheets("BOM").Range("D" & cellX(0) + j).Value) = False And IsEmpty(Worksheets("BOM").cellS(cellX(0) + j, 6 + i).Value) = False Then
                LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, Worksheets("BOM").Range("D" & cellX(0) + j), Worksheets("BOM").cellS(cellX(0) + j, 6 + i), 4
            End If
            
        Next j
        If itemCount > 0 Then
            LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, "MPU-OH", 1, 290
            row_count = row_count + itemCount
        End If
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
    ElseIf C_LOOK("P1", "A") > 0 Then
        LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, "4-pcs1-" & article & WorksheetFunction.Text(artSize(0) + i, "00"), 1, 4
    ElseIf C_LOOK("PCS1", "B") > 0 Then
         LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, "4-pcs1-" & article & WorksheetFunction.Text(artSize(0) + i, "00"), 1, 4
    End If
      
    If C_LOOK("CCS", "B") > 0 Then
        LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, "4-ccs-" & article & WorksheetFunction.Text(artSize(0) + i, "00"), 1, 4
    End If
    If C_LOOK("FCM", "B") > 0 Then
        If InStr(1, Worksheets("BOM").Range("A" & C_LOOK("FCM", "B")), "P", vbTextCompare) = 0 Then
            LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, "4-mcs-" & article & WorksheetFunction.Text(artSize(0) + i, "00"), 1, 4
        End If
    ElseIf C_LOOK("M", "A") > 0 Then
        LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, "4-mcs-" & article & WorksheetFunction.Text(artSize(0) + i, "00"), 1, 4
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

'Printing Components on MPU
Sub PRINTING_MPUCOM(ite As String)

    cellX = CELL_X("PCM")
    For i = 0 To siz
        itemCount = 0
        For j = 0 To cellX(1) - 1
            If IsEmpty(Worksheets("BOM").Range("D" & cellX(0) + j).Value) = False And IsEmpty(Worksheets("BOM").cellS(cellX(0) + j, 6 + i).Value) = False Then
                LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, Worksheets("BOM").Range("D" & cellX(0) + j), Worksheets("BOM").cellS(cellX(0) + j, 6 + i), 4
            End If
        Next j
        If itemCount > 0 Then
            LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, "PRINTING-CHARGES", 1, 290
            row_count = row_count + itemCount
        End If
    Next i

End Sub


'Printing Upper - PCS
Sub PRINTING_UPPER(ite As String)
    
    For i = 0 To siz
    itemCount = 0
    If C_LOOK("CCP", "B") > 0 Then
    
            LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, "4-ccp-" & article & WorksheetFunction.Text(artSize(0) + i, "00"), 1, 4
        
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
    'COMPONENT
    If C_LOOK("COM", "B") > 0 Then
        If UCase(Worksheets("BOM").Range("A" & C_LOOK("COM", "B"))) = "P" Then
        cellX = CELL_X("COM")
            For j = 0 To cellX(1) - 1
            If IsEmpty(Worksheets("BOM").Range("D" & cellX(0) + j).Value) = False Then
                LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, Worksheets("BOM").Range("D" & cellX(0) + j), Worksheets("BOM").cellS(cellX(0) + j, 6 + i), 4
            End If
        Next j
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
                LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, "4-ccp1-" & article & WorksheetFunction.Text(artSize(0) + i, "00"), 1, 4
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
        
         If C_LOOK("COM", "B") > 0 Then
        If UCase(Worksheets("BOM").Range("A" & C_LOOK("COM", "B"))) = "P1" Then
        cellX = CELL_X("COM")
            For j = 0 To cellX(1) - 1
            If IsEmpty(Worksheets("BOM").Range("D" & cellX(0) + j).Value) = False Then
                LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, Worksheets("BOM").Range("D" & cellX(0) + j), Worksheets("BOM").cellS(cellX(0) + j, 6 + i), 4
            End If
        Next j
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
            
                LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, "4-fcm-" & article, Worksheets("BOM").cellS(C_LOOK("FCM", "B"), i + 6), 4
           
        End If
        If C_LOOK("FCM1", "B") > 0 Then
            
                LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, "4-fcm1-" & article, Worksheets("BOM").cellS(C_LOOK("FCM1", "B"), i + 6), 4
            
        End If
        If C_LOOK("FCM2", "B") > 0 Then
           
                LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, "4-fcm2-" & article, Worksheets("BOM").cellS(C_LOOK("FCM2", "B"), i + 6), 4
          
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
        If C_LOOK("COM", "B") > 0 Then
        If UCase(Worksheets("BOM").Range("A" & C_LOOK("COM", "B"))) = "M" Then
        cellX = CELL_X("COM")
            For j = 0 To cellX(1) - 1
            If IsEmpty(Worksheets("BOM").Range("D" & cellX(0) + j).Value) = False Then
                LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, Worksheets("BOM").Range("D" & cellX(0) + j), Worksheets("BOM").cellS(cellX(0) + j, 6 + i), 4
            End If
        Next j
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
        If itemCount > 0 And ite <> "CCP1" Then
            LINE_CELLS ite & WorksheetFunction.Text(artSize(0) + i, "00"), itemCount, "CLICK-OH", 1, 290
            row_count = row_count + itemCount
        End If
    Next i
End Sub

'Folded component #Marking|Folding|Slitting - FCM|FCS n SCF
Sub FOLDED_UPPER(ite As String)
    itemCount = 0
    Dim slit As String
    Select Case ite
        Case "4-fcm-" & article
            cellX = CELL_X("FCM")
            slit = "4-scf-"
        Case "4-fcm1-" & article
            cellX = CELL_X("FCM1")
            slit = "4-scf1-"
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
    'Excludes slitting item if no value in the right most cell after the size columns
    If Not (IsEmpty(Worksheets("BOM").cellS(cellX(0), 7 + siz))) Then
        LINE_CELLS ite, itemCount, slit & article, 1, 4
    End If
    'Items under folding, consumption per mtr length of folded component
    If cellX(1) > 1 Then
        i = cellX(0) + 1
        j = cellX(0) + cellX(1) - 1
        For k = i To j
            If IsEmpty(Worksheets("BOM").Range("D" & k)) = False Then
                LINE_CELLS ite, itemCount, Worksheets("BOM").Range("D" & k), Worksheets("BOM").cellS(k, 7 + siz), 4
            End If
        Next k
    End If
   
    row_count = row_count + itemCount
    itemCount = 0
    If Not (IsEmpty(Worksheets("BOM").cellS(cellX(0), 7 + siz))) Then
        LINE_CELLS slit & article, itemCount, Worksheets("BOM").Range("D" & cellX(0)), Worksheets("BOM").cellS(cellX(0), 7 + siz), 4
        If slit = "4-scf-" Then
           LINE_CELLS slit & article, itemCount, "SLITT-OH", 1, 290
        End If
        row_count = row_count + itemCount
     End If
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
    If ite = "4-scs-" & article Then
        LINE_CELLS ite, itemCount, "SLITT-OH", 1, 290
    End If
   row_count = row_count + itemCount
    
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

Sub LINE_TREE_TITLES(val As Integer)
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

Sub BUILD_TREE(val As Integer)
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
         ElseIf InStr(1, Worksheets("TREE").Range("A" & r), "4-PC", vbTextCompare) > 0 Then
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
'##########################################################################################################
'#####################               MPU ITEMS FROM SOLE                ###################################
'##########################################################################################################
Sub SOLE_ITEMS()

    cellX = CELL_X("MPU")
    rwc = cellX(0)
    'SHOES
    If C_LOOK("JBLD", "C") > 0 Then
        BOM_CELLS "OUTTER SOLE", "4-PUX-0040", Worksheets("BOM").Range("E" & C_LOOK("SOLE", "B")).Value, Worksheets("BOM").Range("F" & C_LOOK("SOLE", "B")).Value, Worksheets("BOM").Range("G" & C_LOOK("SOLE", "B")).Value, Worksheets("BOM").Range("H" & C_LOOK("SOLE", "B")).Value, Worksheets("BOM").Range("I" & C_LOOK("SOLE", "B")).Value, Worksheets("BOM").Range("J" & C_LOOK("SOLE", "B")).Value
        BOM_CELLS "MID SOLE[i]", "5-PO01-0043", "PU-JBLD ISO MDI 2509/IN", 72 / 175 * Worksheets("BOM").Range("F" & 1 + C_LOOK("SOLE", "B")).Value, 72 / 175 * Worksheets("BOM").Range("G" & 1 + C_LOOK("SOLE", "B")).Value, 72 / 175 * Worksheets("BOM").Range("H" & 1 + C_LOOK("SOLE", "B")).Value, 72 / 175 * Worksheets("BOM").Range("I" & 1 + C_LOOK("SOLE", "B")).Value, 72 / 175 * Worksheets("BOM").Range("J" & 1 + C_LOOK("SOLE", "B")).Value
        BOM_CELLS "MID SOLE[p]", "5-PO01-0044", "PU-JBLD POLYOL 721", 100 / 175 * 200 / 225 * Worksheets("BOM").Range("F" & 1 + C_LOOK("SOLE", "B")).Value, 100 / 175 * 200 / 225 * Worksheets("BOM").Range("G" & 1 + C_LOOK("SOLE", "B")).Value, 100 / 175 * 200 / 225 * Worksheets("BOM").Range("H" & 1 + C_LOOK("SOLE", "B")).Value, 100 / 175 * 200 / 225 * Worksheets("BOM").Range("I" & 1 + C_LOOK("SOLE", "B")).Value, 100 / 175 * 200 / 225 * Worksheets("BOM").Range("J" & 1 + C_LOOK("SOLE", "B")).Value
        BOM_CELLS "ADDITIVE[p]", "6-CHM-0146", "JBLD ADDITIVE P721/3/200", 100 / 175 * 25 / 225 * Worksheets("BOM").Range("F" & 1 + C_LOOK("SOLE", "B")).Value, 100 / 175 * 25 / 225 * Worksheets("BOM").Range("G" & 1 + C_LOOK("SOLE", "B")).Value, 100 / 175 * 25 / 225 * Worksheets("BOM").Range("H" & 1 + C_LOOK("SOLE", "B")).Value, 100 / 175 * 25 / 225 * Worksheets("BOM").Range("I" & 1 + C_LOOK("SOLE", "B")).Value, 100 / 175 * 25 / 225 * Worksheets("BOM").Range("J" & 1 + C_LOOK("SOLE", "B")).Value
        BOM_CELLS "MID SOLE[c]", "6-CHM-0019", "PIGMENT KC 1871 WHITE", 3 / 175 * Worksheets("BOM").Range("F" & 1 + C_LOOK("SOLE", "B")).Value, 3 / 175 * Worksheets("BOM").Range("G" & 1 + C_LOOK("SOLE", "B")).Value, 3 / 175 * Worksheets("BOM").Range("H" & 1 + C_LOOK("SOLE", "B")).Value, 3 / 175 * Worksheets("BOM").Range("I" & 1 + C_LOOK("SOLE", "B")).Value, 3 / 175 * Worksheets("BOM").Range("J" & 1 + C_LOOK("SOLE", "B")).Value
        BOM_CELLS "IMC -WH", "6-CHM-0156", "Water Base IMC White KB 4505", 0.005, 0.005, 0.005, 0.005, 0.005
        BOM_CELLS "WB 07A", "6-CHM-0126", "RELEASE AGENT W.B 711/07A", 0.001, 0.001, 0.001, 0.001, 0.001
        BOM_CELLS "1602", "6-CHM-0010", "RELEASE AGENT KECK� 1602/18", 0.001, 0.001, 0.001, 0.001, 0.001
    'DOUBLE COLOR
    ElseIf C_LOOK("2 COLOR", "C") > 0 Then
        BOM_CELLS "OUTTER SOLE", "4-PUX-0004", Worksheets("BOM").Range("E" & C_LOOK("SOLE", "B")).Value, Worksheets("BOM").Range("F" & C_LOOK("SOLE", "B")).Value, Worksheets("BOM").Range("G" & C_LOOK("SOLE", "B")).Value, Worksheets("BOM").Range("H" & C_LOOK("SOLE", "B")).Value, Worksheets("BOM").Range("I" & C_LOOK("SOLE", "B")).Value, Worksheets("BOM").Range("J" & C_LOOK("SOLE", "B")).Value
        BOM_CELLS "MID SOLE[i]", "5-PO01-0004", "ISO 163", 51 / 154 * Worksheets("BOM").Range("F" & 1 + C_LOOK("SOLE", "B")).Value, 51 / 154 * Worksheets("BOM").Range("G" & 1 + C_LOOK("SOLE", "B")).Value, 51 / 154 * Worksheets("BOM").Range("H" & 1 + C_LOOK("SOLE", "B")).Value, 51 / 154 * Worksheets("BOM").Range("I" & 1 + C_LOOK("SOLE", "B")).Value, 51 / 154 * Worksheets("BOM").Range("J" & 1 + C_LOOK("SOLE", "B")).Value
        BOM_CELLS "MID SOLE[p]", "5-PO01-0042", "POLY VB1", 100 / 154 * Worksheets("BOM").Range("F" & 1 + C_LOOK("SOLE", "B")).Value, 100 / 154 * Worksheets("BOM").Range("G" & 1 + C_LOOK("SOLE", "B")).Value, 100 / 154 * Worksheets("BOM").Range("H" & 1 + C_LOOK("SOLE", "B")).Value, 100 / 154 * Worksheets("BOM").Range("I" & 1 + C_LOOK("SOLE", "B")).Value, 100 / 154 * Worksheets("BOM").Range("J" & 1 + C_LOOK("SOLE", "B")).Value
        BOM_CELLS "MID SOLE[c]", "6-CHM-0019", "PIGMENT KC 1871 WHITE", 3 / 154 * Worksheets("BOM").Range("F" & 1 + C_LOOK("SOLE", "B")).Value, 3 / 154 * Worksheets("BOM").Range("G" & 1 + C_LOOK("SOLE", "B")).Value, 3 / 154 * Worksheets("BOM").Range("H" & 1 + C_LOOK("SOLE", "B")).Value, 3 / 154 * Worksheets("BOM").Range("I" & 1 + C_LOOK("SOLE", "B")).Value, 3 / 154 * Worksheets("BOM").Range("J" & 1 + C_LOOK("SOLE", "B")).Value
        BOM_CELLS "IMC -WH", "6-CHM-0156", "Water Base IMC White KB 4505", 0.004, 0.004, 0.004, 0.004, 0.004
        BOM_CELLS "WB 07A", "6-CHM-0126", "RELEASE AGENT W.B 711/07A", 0.001, 0.001, 0.001, 0.001, 0.001
        BOM_CELLS "1602", "6-CHM-0010", "RELEASE AGENT KECK� 1602/18", 0.001, 0.001, 0.001, 0.001, 0.001
        
    Else:
        BOM_CELLS "OUTTER SOLE", "4-PUX-0004", Worksheets("BOM").Range("E" & C_LOOK("SOLE", "B")).Value, Worksheets("BOM").Range("F" & C_LOOK("SOLE", "B")).Value, Worksheets("BOM").Range("G" & C_LOOK("SOLE", "B")).Value, Worksheets("BOM").Range("H" & C_LOOK("SOLE", "B")).Value, Worksheets("BOM").Range("I" & C_LOOK("SOLE", "B")).Value, Worksheets("BOM").Range("J" & C_LOOK("SOLE", "B")).Value
        'SOFT
        If C_LOOK("SOFT", "C") > 0 Then
            BOM_CELLS "MID SOLE[i]", "5-PO01-0004", "ISO  99055290 SHISO GE -163 (DOW)", 34 / 134 * Worksheets("BOM").Range("F" & 1 + C_LOOK("SOLE", "B")).Value, 34 / 134 * Worksheets("BOM").Range("G" & 1 + C_LOOK("SOLE", "B")).Value, 34 / 134 * Worksheets("BOM").Range("H" & 1 + C_LOOK("SOLE", "B")).Value, 34 / 134 * Worksheets("BOM").Range("I" & 1 + C_LOOK("SOLE", "B")).Value, 34 / 134 * Worksheets("BOM").Range("J" & 1 + C_LOOK("SOLE", "B")).Value
            BOM_CELLS "MID SOLE[p]", "5-PO01-0018", "VORALAST SOFT POLYOL GM899-DOW", 100 / 134 * 200 / 202.5 * Worksheets("BOM").Range("F" & 1 + C_LOOK("SOLE", "B")).Value, 100 / 134 * 200 / 202.5 * Worksheets("BOM").Range("G" & 1 + C_LOOK("SOLE", "B")).Value, 100 / 134 * 200 / 202.5 * Worksheets("BOM").Range("H" & 1 + C_LOOK("SOLE", "B")).Value, 100 / 134 * 200 / 202.5 * Worksheets("BOM").Range("I" & 1 + C_LOOK("SOLE", "B")).Value, 100 / 134 * 200 / 202.5 * Worksheets("BOM").Range("J" & 1 + C_LOOK("SOLE", "B")).Value
            BOM_CELLS "ADDITIVE[p]", "6-CHM-0115", "VORALAST NATURAL 817 ADDICTIVE", 100 / 134 * 2.5 / 202.5 * Worksheets("BOM").Range("F" & 1 + C_LOOK("SOLE", "B")).Value, 100 / 134 * 2.5 / 202.5 * Worksheets("BOM").Range("G" & 1 + C_LOOK("SOLE", "B")).Value, 100 / 134 * 2.5 / 202.5 * Worksheets("BOM").Range("H" & 1 + C_LOOK("SOLE", "B")).Value, 100 / 134 * 2.5 / 202.5 * Worksheets("BOM").Range("I" & 1 + C_LOOK("SOLE", "B")).Value, 100 / 134 * 2.5 / 202.5 * Worksheets("BOM").Range("J" & 1 + C_LOOK("SOLE", "B")).Value
            
        End If
        BOM_CELLS "ELFI", "6-ADH-0029", "ELFI", 0.0003, 0.0003, 0.0003, 0.0003, 0.0003
        BOM_CELLS "WB 07A", "6-CHM-0126", "RELEASE AGENT W.B 711/07A", 0.0008, 0.0008, 0.0008, 0.0008, 0.0008
    End If
    
End Sub
'############################################################################################################
'############################################################################################################


Sub BOM_CELLS(valueC, valueD, valueE, valueF, valueG, valueH, valueI, valueJ As String)
    Worksheets("BOM").Range("C" & rwc).Value = UCase(valueC)
    Worksheets("BOM").Range("D" & rwc).Value = UCase(valueD)
    Worksheets("BOM").Range("E" & rwc).Value = UCase(valueE)
    Worksheets("BOM").Range("F" & rwc).Value = Round(valueF, 4)
    Worksheets("BOM").Range("G" & rwc).Value = Round(valueG, 4)
    Worksheets("BOM").Range("H" & rwc).Value = Round(valueH, 4)
    Worksheets("BOM").Range("I" & rwc).Value = Round(valueI, 4)
    Worksheets("BOM").Range("J" & rwc).Value = Round(valueJ, 4)
    rwc = rwc + 1
End Sub

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

Function MC_ITEMS(SIZEE As String) As Integer()
    Dim sc_count(6) As Integer
        
        Select Case SIZEE
            Case "6X10"
                sc_count(0) = 4
                sc_count(4) = 5
                sc_count(1) = 5
                sc_count(2) = 5
                sc_count(3) = 5
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
            Case "6X10_SHOES"
                sc_count(0) = 3
                sc_count(1) = 3
                sc_count(2) = 4
                sc_count(3) = 4
                sc_count(4) = 4
                sc_count(5) = 18
                sc_count(6) = 1
            Case "7X10Z"
                sc_count(0) = 3
                sc_count(1) = 4
                sc_count(2) = 3
                sc_count(3) = 2
                sc_count(4) = 12
                sc_count(5) = 0
                sc_count(6) = 2
                
            Case "5X9"
                sc_count(0) = 6
                sc_count(1) = 6
                sc_count(2) = 6
                sc_count(3) = 6
                sc_count(4) = 6
                sc_count(5) = 30
                sc_count(6) = 1
            Case "5X9Z"
                sc_count(0) = 2
                sc_count(1) = 3
                sc_count(2) = 3
                sc_count(3) = 2
                sc_count(4) = 2
                sc_count(5) = 12
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
            Case "9X11Z"
                sc_count(0) = 4
                sc_count(1) = 4
                sc_count(2) = 4
                sc_count(3) = 12
                sc_count(4) = 0
                sc_count(5) = 0
                sc_count(6) = 52
            Case "11X12"
                sc_count(0) = 6
                sc_count(1) = 6
                sc_count(2) = 12
                sc_count(3) = 0
                sc_count(4) = 0
                sc_count(5) = 0
                sc_count(6) = 1
             Case "11X12Z"
                sc_count(0) = 6
                sc_count(1) = 6
                sc_count(2) = 12
                sc_count(3) = 0
                sc_count(4) = 0
                sc_count(5) = 0
                sc_count(6) = 1
            Case "L6X9_SMART"
                sc_count(0) = 1
                sc_count(1) = 2
                sc_count(2) = 2
                sc_count(3) = 1
                sc_count(4) = 6
                sc_count(5) = 0
                sc_count(6) = 5
            Case "L5X8_SMART"
                sc_count(0) = 1
                sc_count(1) = 2
                sc_count(2) = 2
                sc_count(3) = 1
                sc_count(4) = 6
                sc_count(5) = 0
                sc_count(6) = 2
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


'###############################################
'#####               UPDATES                 ###
'###############################################

'@  Made SLIT-OH only to main slitting (not for scs1,scs2,scf1,scf2)       #xx/xx/2020
'@ Made CLICK-OH only to main clicking (not for ccp1)                      #03/06/2020



