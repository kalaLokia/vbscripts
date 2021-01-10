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
    c_artno = Ucase(sArtno)
End Property

Public Property Get Color() As String
    Color = c_color
End Property
Public Property Let Color(ByVal sColor As String)
    c_color = Ucase(sColor)
End Property

Public Property Get Category() As String
    Category = c_category
End Property
Public Property Let Category(ByVal sCategory As String)
    c_category = Ucase(sCategory)
End Property

Public Property Get Jobno() As String
    Jobno = c_jobno
End Property
Public Property Let Jobno(ByVal sJobno As String)
    c_jobno = Ucase(sJobno)
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
    Select Case c_process & "-" & c_artno & "-" & c_category & sSize
        Case "CCS-3391-B0"
        ArticleItem = "3391-NB-G"
        Case "CCS-3391-B1"
        ArticleItem = "3391-NB-B"
        Case "CCS-3391-B2"
        ArticleItem = "3391-NB-B"
        Case "CCS-3391-B3"
        ArticleItem = "3391-NB-B"
        Case "CCS-3391-B4"
        ArticleItem = "3391-NB-B"
        Case "CCS-3391-B5"
        ArticleItem = "3391-NB-B"
        Case Else
        ArticleItem = Article
    End Select
End Function

' Writting to sheet datas
Public Sub WriteToSheet(rowNo As Integer, colNo As Integer,  Optional ByVal sSize As Integer = 0)
     Worksheets(target_sheet).Range("A" & rowNo).Value = sSize
     Worksheets(target_sheet).Range("B" & rowNo).Value = c_jobno
     
     Worksheets(target_sheet).Range("E" & rowNo).Value = "FB/CF001"
     Worksheets(target_sheet).Range("F" & rowNo).Value = "FB/CF001"
     Worksheets(target_sheet).Range("G" & rowNo).Value = Article
     
     Worksheets(target_sheet).Range("J" & rowNo).Value = "=CLICKING!$T$" & colNo

     if sSize = 0 Then
          Worksheets(target_sheet).Range("C" & rowNo).Value = "4-"& c_process & "-" & ArticleItem(sSize)
          Worksheets(target_sheet).Range("D" & rowNo).Value = Worksheets("CLICKING").Range("U" & colNo).Value
          Worksheets(target_sheet).Range("I" & rowNo).Value = Worksheets("CLICKING").Range("U" & colNo).Value
     Else:
          Worksheets(target_sheet).Range("C" & rowNo).Value = "4-"& c_process & "-" & ArticleItem(sSize) & WorksheetFunction.Text(sSize, "00")
          Worksheets(target_sheet).Range("D" & rowNo).Value = "=I" & rowNo & "*J" & rowNo
          Worksheets(target_sheet).Range("I" & rowNo).Value = Worksheets("CLICKING").Cells(colNo, sSize + 6)
     End if
     
End Sub

'Write headers to the datas sheet
Public Sub WriteHeaders()
     Worksheets(target_sheet).Range("B1").Value = "SIZE"
     Worksheets(target_sheet).Range("B1").Value = "JOB NO."
     Worksheets(target_sheet).Range("C1").Value = "SAP ITEM CODE"
     Worksheets(target_sheet).Range("D1").Value = "QTY"
     Worksheets(target_sheet).Range("E1").Value = "H. WHR"
     Worksheets(target_sheet).Range("F1").Value = "C. WHR"
     Worksheets(target_sheet).Range("I1").Value = "planqty"
     Worksheets(target_sheet).Range("J1").Value = "plan"
End Sub

