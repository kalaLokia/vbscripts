Attribute VB_Name = "PACKING_DB"
' Version 0.4

Sub PACKING_DB()

    Dim excel As excel.Application
    Dim wb As excel.Workbook
    Dim sht As excel.Worksheet
    Dim shet As excel.Worksheet
    Dim lookItem, lookUpValue, sc_raw, mc_raw As Variant
    Dim title(11), items(11), itemDesc(11), msc(2) As Variant
    Dim r As Integer
    
    sc_raw = MSC_LOOK("SC")
    mc_raw = MSC_LOOK("MC")
    
    lookUpValue = UCase(Worksheets("BOM").Range("D2") & "_" & Worksheets("BOM").Range("D5"))
    If UCase(Worksheets("BOM").Range("D6")) = "SHOES" Then
        lookUpValue = lookUpValue & "SHOE"
    End If

    If IsEmpty(Worksheets("BOM").Range("D1")) = False Then
        lookUpValue = UCase(lookUpValue & Worksheets("BOM").Range("D1"))
    End If

    
    Set excel = CreateObject("excel.Application")
    excel.Visible = False
    Set wb = excel.Workbooks.Open("E:\SOLID_DATA\PACKING_MATERIALS.xlsx")
    Set sht = wb.Worksheets("ITEMS")
    Set shet = wb.Worksheets("DB")
    
    sht.Activate
    'MsgBox sht.Range("A5")
 
        lookItem = db_LOOK(lookUpValue, sht)
    
        items(0) = hasValue(sht.cellS(lookItem, 3))     'MC
        items(1) = hasValue(sht.cellS(lookItem, 4))     'STICKER MC
        items(2) = hasValue(sht.cellS(lookItem, 5))     'BELT
        items(3) = hasValue(sht.cellS(lookItem, 6))     'ADH, TAPE
        items(4) = hasValue(sht.cellS(lookItem, 7))     'VENT. STICKER
        items(5) = hasValue(sht.cellS(lookItem, 8))     'WEIGHT STICKER
        items(6) = hasValue(sht.cellS(lookItem, 10))     'SC
        items(7) = hasValue(sht.cellS(lookItem, 11))    'PRICE STICKER
        items(8) = hasValue(sht.cellS(lookItem, 12))    'TAG LOOP
        items(9) = hasValue(sht.cellS(lookItem, 13))    'TAG
        items(10) = hasValue(sht.cellS(lookItem, 14))    'TISSUE PAPER
        items(11) = hasValue(sht.cellS(lookItem, 15))   'SILICA GEL
        
        title(0) = hasValue(sht.cellS(1, 3))        'MC
        title(1) = hasValue(sht.cellS(1, 4))        'STICKER MC
        title(2) = hasValue(sht.cellS(1, 5))        'BELT
        title(3) = hasValue(sht.cellS(1, 6))        'ADH, TAPE
        title(4) = hasValue(sht.cellS(1, 7))        'VENT. STICKER
        title(5) = hasValue(sht.cellS(1, 8))        'WEIGHT STICKER
        title(6) = hasValue(sht.cellS(1, 10))        'SC
        title(7) = hasValue(sht.cellS(1, 11))       'PRICE STICKER
        title(8) = hasValue(sht.cellS(1, 12))       'TAG LOOP
        title(9) = hasValue(sht.cellS(1, 13))       'TAG
        title(10) = hasValue(sht.cellS(1, 14))       'TISSUE PAPER
        title(11) = hasValue(sht.cellS(1, 15))      'SILICA GEL
    
        itemDesc(0) = isInDb(items(0), shet)     'MC
        itemDesc(1) = isInDb(items(1), shet)     'STICKER MC
        itemDesc(2) = isInDb(items(2), shet)     'BELT
        itemDesc(3) = isInDb(items(3), shet)     'ADH, TAPE
        itemDesc(4) = isInDb(items(4), shet)     'VENT. STICKER
        itemDesc(5) = isInDb(items(5), shet)     'VENT. STICKER
        itemDesc(6) = isInDb(items(6), shet)     'SC
        itemDesc(7) = isInDb(items(7), shet)     'PRICE STICKER
        itemDesc(8) = isInDb(items(8), shet)     'TAG LOOP
        itemDesc(9) = isInDb(items(9), shet)     'TAG
        itemDesc(10) = isInDb(items(10), shet)     'TISSUE PAPER
        itemDesc(11) = isInDb(items(11), shet)    'SILICA GEL
        
    'Master box BELT and ADH. TAPE values
        If items(0) <> 0 Then
            msc(0) = shet.cellS(db_LOOK(items(0), shet), 6)
            msc(1) = shet.cellS(db_LOOK(items(0), shet), 7)
        End If
      
    wb.Close SaveChanges:=False
    'MC
    r = 0
    If mc_raw > 0 Then
    For i = 0 To 5
       
        If items(i) > 0 Then
            Worksheets("BOM").Range("c" & mc_raw + r).Value = title(i)
            Worksheets("BOM").Range("d" & mc_raw + r).Value = items(i)
            Worksheets("BOM").Range("e" & mc_raw + r).Value = itemDesc(i)
            If Worksheets("BOM").Range("d" & mc_raw + r).Value = "7-AT-0015" Then
                 Worksheets("BOM").Range("f" & mc_raw + r).Value = msc(1)
            ElseIf Worksheets("BOM").Range("d" & mc_raw + r).Value = "7-BT-0001" Then
                Worksheets("BOM").Range("f" & mc_raw + r).Value = msc(0)
    
            Else:
            Worksheets("BOM").Range("f" & mc_raw + r).Value = 1
            End If
            'If i > 0 Then
             '   Worksheets("BOM").Range("f" & mc_raw + 2).Value = msc(0)
             '   Worksheets("BOM").Range("f" & mc_raw + 3).Value = msc(1)
            'End If
            r = r + 1
            End If
    Next i
    End If
    r = 0
    If sc_raw > 0 Then
    For i = 6 To 11
        If items(i) > 0 Then
            Worksheets("BOM").Range("c" & sc_raw + r).Value = title(i)
            Worksheets("BOM").Range("d" & sc_raw + r).Value = items(i)
            Worksheets("BOM").Range("e" & sc_raw + r).Value = itemDesc(i)
            Worksheets("BOM").Range("f" & sc_raw + r).Value = 1
            If items(i) = "7-OT-0007" Then
                Worksheets("BOM").Range("f" & sc_raw + r).Value = 0.002
            End If
            If items(i) = "7-PS-0014" And InStr(1, lookUpValue, "SMARTAK", vbTextCompare) > 0 Then
                Worksheets("BOM").Range("f" & sc_raw + r).Value = 2
            End If
            r = r + 1
        End If
    Next i
    End If
End Sub

Function hasValue(val As Variant)
    If IsEmpty(val) Then
        hasValue = 0
    Else:
        hasValue = val
    End If
End Function


Function isInDb(ite As Variant, st As Worksheet)
    If ite <> 0 Then
        isInDb = st.cellS(db_LOOK(ite, st), 3)
    Else:
        isInDb = 0
    End If
End Function


Function db_LOOK(lookValue As Variant, st As Worksheet)
    Dim look As Variant
   With st
        On Error Resume Next
        look = Application.WorksheetFunction.MATCH(lookValue, .Range("A:A"), 0)
        On Error GoTo 0
    End With
    If look = 0 Then
        look = 1
    End If
    db_LOOK = look
End Function

Function MSC_LOOK(lookValue As Variant)
    Dim look As Variant
   With Worksheets("BOM")
        On Error Resume Next
        look = Application.WorksheetFunction.MATCH(lookValue, .Range("B:B"), 0)
        On Error GoTo 0
    End With
    MSC_LOOK = look
End Function

