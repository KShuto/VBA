Attribute VB_Name = "Module1"
Sub HowToAutoFiler()

    Dim TgtSheet As Object
    Dim TgtCell As Range
    
    Set TgtSheet = Worksheets("01")
    Set TgtCell = TgtSheet.Range("A6:D6")
    
    
    
    
    ' オートフィルタをかける
    With TgtCell
        .AutoFilter _
            Field:=3, _
            Criteria1:=Range("B3"), _
            Operator:=xlOr, _
            Criteria2:=Range("B4")
'        .AutoFilter _
'            Field:=4, _
'            Operator:=xlTop10Items
    End With
    
End Sub
