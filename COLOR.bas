Attribute VB_Name = "COLOR"
Function FROMSAPCOLOR(col As String)

    col = LCase(col)
    Select Case colour
        Case "bk"
        col = "black"
        Case "br"
        col = "brown"
        Case "bl"
        col = "blue"
        Case "rd"
        col = "red"
        Case "pk"
        col = "pink"
        Case "ta"
        col = "tan"
        Case "pe"
        col = "pink blue"
        Case "lr"
        col = "blue red"
        Case "gy"
        col = "grey"
        Case "gd"
        col = "gold"
        Case "co"
        col = "copper"
        Case "wt"
        col = "white"
        Case "gr"
        col = "green"
        Case "or"
        col = "orange"
        Case "nb"
        col = "navy blue"
        Case "dn"
        col = "dark green"
        Case "mh"
        col = "mehandi"
        Case "ph"
        col = "peach"
        Case "black white"
        col = "wk"
        Case "ov"
        col = "olive"
        Case "sk"
        col = "school black"
        Case "tb"
        col = "tan black"
        Case "mr"
        col = "maroon"
        Case "st"
        col = "special tan"
        Case "sa"
        col = "special black"
        Case "nr"
        col = "navy blue red"
        Case "ng"
        col = "navy blue grey"
        Case "kg"
        col = "black grey"
        Case "gd"
        col = "gold"
        Case "tr"
        col = "tan brown"
        Case "ny"
        col = "navy"
        Case "rk"
        col = "red black"
        Case "dr"
        col = "dark brown"
        Case Else
        col = "NOT-FOUND"
    End Select
  FROMSAPCOLOR = col
End Function
