VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsClick"
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

Private arr_cc_item As Variant
Private arr_ins_ccp1 As Variant
Private arr_ins_ccp1_com As Variant
Private arr_ins_ccp1_sep As Variant
Private arr_up_no_ccs As Variant
Private arr_up_ccs_com As Variant
Private arr_up_ccp1 As Variant
Private arr_up_ccp1_com As Variant
Private arr_up_ccf As Variant

'Load data
Private Sub Table_To_Array()
    Dim tbl_cc_item As ListObject
    Dim tbl_ins_ccp1 As ListObject
    Dim tbl_ins_ccp1_com As ListObject
    Dim tbl_ins_ccp1_sep As ListObject
    Dim tbl_up_no_ccs As ListObject
    Dim tbl_up_ccs_com As ListObject
    Dim tbl_up_ccp1_com As ListObject
    Dim tbl_up_ccp1 As ListObject
    Dim tbl_up_ccf As ListObject
    Dim ws As Worksheet

    Set ws = Worksheets("db")
    ws.Activate

    Set tbl_cc_item = ws.ListObjects("CC_SAP_ITEM")
    Set tbl_ins_ccp1 = ws.ListObjects("INS_CCP1")
    Set tbl_ins_ccp1_com = ws.ListObjects("INS_CCP1_COM")
    Set tbl_ins_ccp1_sep = ws.ListObjects("INS_CCP1_SEP")
    Set tbl_up_no_ccs = ws.ListObjects("UP_NO_CCS")
    Set tbl_up_ccs_com = ws.ListObjects("UP_CCS_COM")
    Set tbl_up_ccp1 = ws.ListObjects("UP_CCP1")
    Set tbl_up_ccp1_com = ws.ListObjects("UP_CCP1_COM")
    Set tbl_up_ccf = ws.ListObjects("UP_CCF")

    tbl_cc_item.DataBodyRange.Select
    tbl_ins_ccp1.DataBodyRange.Select
    tbl_ins_ccp1_com.DataBodyRange.Select
    tbl_ins_ccp1_sep.DataBodyRange.Select
    tbl_up_no_ccs.DataBodyRange.Select
    tbl_up_ccs_com.DataBodyRange.Select
    tbl_up_ccp1.DataBodyRange.Select
    tbl_up_ccp1_com.DataBodyRange.Select
    tbl_up_ccf.DataBodyRange.Select

    arr_cc_item = tbl_cc_item.DataBodyRange
    arr_ins_ccp1 = tbl_ins_ccp1.DataBodyRange
    arr_ins_ccp1_com = tbl_ins_ccp1_com.DataBodyRange
    arr_ins_ccp1_sep = tbl_ins_ccp1_sep.DataBodyRange
    arr_up_no_ccs = tbl_up_no_ccs.DataBodyRange
    arr_up_ccs_com = tbl_up_ccs_com.DataBodyRange
    arr_up_ccp1 = tbl_up_ccp1.DataBodyRange
    arr_up_ccp1_com = tbl_up_ccp1_com.DataBodyRange
    arr_up_ccf = tbl_up_ccf.DataBodyRange

End Sub


' Initializing the object
Private Sub Class_Initialize()
    target_sheet = "datas"
    Table_To_Array ' Load array values from table
End Sub


Public Property Get Arr_IP1() As Variant
Arr_IP1 = arr_ins_ccp1
End Property
Public Property Get Arr_IP1_COM() As Variant
Arr_IP1_COM = arr_ins_ccp1_com
End Property
Public Property Get Arr_IP1_SEP() As Variant
Arr_IP1_SEP = arr_ins_ccp1_sep
End Property
Public Property Get Arr_UNOS() As Variant
Arr_UNOS = arr_up_no_ccs
End Property
Public Property Get Arr_US_COM() As Variant
Arr_US_COM = arr_up_ccs_com
End Property
Public Property Get Arr_UP1() As Variant
Arr_UP1 = arr_up_ccp1
End Property
Public Property Get Arr_UP1_COM() As Variant
Arr_UP1_COM = arr_up_ccp1_com
End Property
Public Property Get Arr_UF() As Variant
Arr_UF = arr_up_ccf
End Property

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
        ArticleCategory = c_artno & "-" & c_category
        Else:
        ArticleCategory = c_artno
    End If
End Function

' Getting SAP CLK item code for the article
Public Function ArticleItem(Optional ByVal sSize As Integer = 0) As String
    lookUpValue = c_process & "-" & c_artno & "-" & c_color & "-" & c_category & sSize
    ArticleItem = Article
    For i = 1 To UBound(arr_cc_item)
        If StrComp(lookUpValue, arr_cc_item(i, 1)) = 0 Then
            ArticleItem = arr_cc_item(i, 2)
            Exit For
        End If
        Next
        'Debug.Print lookUpValue & "  :  " & ArticleItem
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
    Worksheets(target_sheet).Range("G1").Value = "ARTICLE"
    Worksheets(target_sheet).Range("I1").Value = "planqty"
    Worksheets(target_sheet).Range("J1").Value = "plan"
End Sub

Public Function MatchUp(arr As Variant, lookUpValue As String) As Boolean
    ' Looks up a string in an array
    For i = 1 To UBound(arr)
        MatchUp = False
        If StrComp(lookUpValue, arr(i, 1)) = 0 Then
            MatchUp = True
            Exit For
        End If
        Next
End Function


' Created by kalaLokia
