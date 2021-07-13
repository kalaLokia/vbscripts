VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsClicking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private target_sheet As String

Private c_process As String
Private c_artno As String
Private c_color As String
Private c_category As String
Private c_jobno As String

Private c_plan As Long
Private c_planqty As Long


' Initializing the object
Private Sub Class_Initialize()
    target_sheet = "datas"
End Sub

Public Property Get Artno() As String
    Artno = c_artno
End Property
Public Property Let Artno(ByVal sArtno As String)
    c_artno = UCase(sArtno)
End Property

Public Property Get Color() As String
    Color = c_color
End Property
Public Property Let Color(ByVal sColor As String)
    c_color = UCase(sColor)
End Property

Public Property Get Category() As String
    Category = c_category
End Property
Public Property Let Category(ByVal sCategory As String)
    c_category = UCase(sCategory)
End Property

Public Property Get Jobno() As String
    Jobno = c_jobno
End Property
Public Property Let Jobno(ByVal sJobno As String)
    c_jobno = UCase(sJobno)
End Property

Public Property Get Plan() As Long
    Plan = c_plan
End Property
Public Property Let Plan(ByVal sPlan As Long)
    c_plan = sPlan
End Property

Public Property Get PlanQty() As Long
    PlanQty = c_planqty
End Property
Public Property Let PlanQty(ByVal sPlanQty As Long)
    c_planqty = sPlanQty
End Property

Public Property Get Process() As String
    Process = c_process
End Property
Public Property Let Process(ByVal sProcess As String)
    c_process = sProcess
End Property

' Getting SAP article model
Public Function Article() As String
        Article = c_artno & "-" & c_color & "-" & c_category
End Function

' Differentiate between Gents vs Kids
Public Function ArticleCategory() As String
    If c_category = "B" Then
        ArticleCategory = c_artno & c_category
    Else:
        ArticleCategory = c_artno
    End If
End Function

' Getting SAP CLK item code for the article
Public Function ArticleItem(Optional ByVal sSize As Integer = 0) As String
    Select Case c_process & "-" & c_artno & c_color & "-" & c_category & sSize
        ' 3391 Boys angutta is same as 3391 N Blue Gents : CCS-3391-NB-G
        Case "CCS-3391GR-B0"
        ArticleItem = "3391-NB-G"
        Case "CCS-3391RD-B0"
        ArticleItem = "3391-NB-G"
        Case "CCS-3391GY-B0"
        ArticleItem = "3391-NB-G"
        Case "CCS-3391NB-B0"
        ArticleItem = "3391-NB-G"

        ' 3391 N Blue Gents angutta is common in all colors except for black : CCS-3391-NB-G
        Case "CCS-3391GR-G0"
        ArticleItem = "3391-NB-G"
        Case "CCS-3391RD-G0"
        ArticleItem = "3391-NB-G"
        Case "CCS-3391GY-G0"
        ArticleItem = "3391-NB-G"
        Case "CCS-DG9133NB-G0"
        ArticleItem = "3391-NB-G"
        Case "CCS-DG9133GR-G0"
        ArticleItem = "3391-NB-G"
        Case "CCS-DG9133RD-G0"
        ArticleItem = "3391-NB-G"
        Case "CCS-DG9133GY-G0"
        ArticleItem = "3391-NB-G"
        ' Other common DG9119-NB
        Case "CCS-DG9119NB-G0"
        ArticleItem = "3391-NB-G"
        Case "CCS-3394GY-G0"
        ArticleItem = "3391-NB-G"

        ' Cushion Pad is common in all colors of 3391 boys article : CCS-3391-NB-B{size}
        Case "CCS-3391GR-B1"
        ArticleItem = "3391-NB-B"
        Case "CCS-3391GR-B2"
        ArticleItem = "3391-NB-B"
        Case "CCS-3391GR-B3"
        ArticleItem = "3391-NB-B"
        Case "CCS-3391GR-B4"
        ArticleItem = "3391-NB-B"
        Case "CCS-3391GR-B5"
        ArticleItem = "3391-NB-B"

        Case "CCS-3391RD-B1"
        ArticleItem = "3391-NB-B"
        Case "CCS-3391RD-B2"
        ArticleItem = "3391-NB-B"
        Case "CCS-3391RD-B3"
        ArticleItem = "3391-NB-B"
        Case "CCS-3391RD-B4"
        ArticleItem = "3391-NB-B"
        Case "CCS-3391RD-B5"
        ArticleItem = "3391-NB-B"
        
        Case "CCS-3391GY-B1"
        ArticleItem = "3391-NB-B"
        Case "CCS-3391GY-B2"
        ArticleItem = "3391-NB-B"
        Case "CCS-3391GY-B3"
        ArticleItem = "3391-NB-B"
        Case "CCS-3391GY-B4"
        ArticleItem = "3391-NB-B"
        Case "CCS-3391GY-B5"
        ArticleItem = "3391-NB-B"
        
        Case "CCS-3374BR-B1"
        ArticleItem = "3391-NB-B"
        Case "CCS-3374BR-B2"
        ArticleItem = "3391-NB-B"
        Case "CCS-3374BR-B3"
        ArticleItem = "3391-NB-B"
        Case "CCS-3374BR-B4"
        ArticleItem = "3391-NB-B"
        Case "CCS-3374BR-B5"
        ArticleItem = "3391-NB-B"
        
        '3290 PATTIES COMMON FOR GIANTS ALSO
        Case "CCP1-3290BK-X0"
        ArticleItem = "3290-BK-G"
        Case "CCP1-3290BR-X0"
        ArticleItem = "3290-BR-G"
        
        '3059 NB PATTIES SAME AS DG9540 NB
        'Case "CCP1-DG9540NB-G0"
        'ArticleItem = "3059-NB-G"

        Case Else
        ArticleItem = Article
    End Select
End Function

' Writting to sheet datas
Public Sub WriteToSheet(rowNo As Integer, colNo As Integer, Optional ByVal sSize As Integer = 0)
    
    Worksheets(target_sheet).Range("A" & rowNo).Value = sSize
     Worksheets(target_sheet).Range("B" & rowNo).Value = c_jobno
     
     Worksheets(target_sheet).Range("E" & rowNo).Value = "FB/CF001"
     Worksheets(target_sheet).Range("F" & rowNo).Value = "FB/CF001"
     Worksheets(target_sheet).Range("G" & rowNo).Value = Article
     
     Worksheets(target_sheet).Range("J" & rowNo).Value = "=CLICKING!$T$" & colNo

     If sSize = 0 Then
          Worksheets(target_sheet).Range("C" & rowNo).Value = "4-" & c_process & "-" & ArticleItem(sSize)
          Worksheets(target_sheet).Range("D" & rowNo).Value = Worksheets("CLICKING").Range("U" & colNo).Value
          Worksheets(target_sheet).Range("I" & rowNo).Value = Worksheets("CLICKING").Range("U" & colNo).Value
     Else:
          Worksheets(target_sheet).Range("C" & rowNo).Value = "4-" & c_process & "-" & ArticleItem(sSize) & WorksheetFunction.Text(sSize, "00")
          Worksheets(target_sheet).Range("D" & rowNo).Value = "=I" & rowNo & "*J" & rowNo
          Worksheets(target_sheet).Range("I" & rowNo).Value = Worksheets("CLICKING").cellS(colNo, sSize + 6)
     End If
     
End Sub

'Write headers to the datas sheet
Public Sub WriteHeaders()
     Worksheets(target_sheet).Range("A1").Value = "SIZE"
     Worksheets(target_sheet).Range("B1").Value = "JOB NO."
     Worksheets(target_sheet).Range("C1").Value = "SAP ITEM CODE"
     Worksheets(target_sheet).Range("D1").Value = "QTY"
     Worksheets(target_sheet).Range("E1").Value = "H. WHR"
     Worksheets(target_sheet).Range("F1").Value = "C. WHR"
     Worksheets(target_sheet).Range("I1").Value = "planqty"
     Worksheets(target_sheet).Range("J1").Value = "plan"
End Sub

' Created by kalaLokia
