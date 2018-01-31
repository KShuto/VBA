Attribute VB_Name = "Module1"
Sub HowToAutoFiler()

    Dim TgtSheet As Object
    Dim TgtCell As Range
    
    Set TgtSheet = Worksheets("01")
    Set TgtCell = TgtSheet.Range("A10:D10")
    
    ' オートフィルタをかける
    With TgtCell
        .AutoFilter _
            Field:=Range("B4"), _
            Criteria1:=Range("B5"), _
            Operator:=xlOr, _
            Criteria2:=Range("B7")
'        .AutoFilter _
'            Field:=4, _
'            Operator:=xlTop10Items
    End With
    
End Sub
