Attribute VB_Name = "Module1"
'Like演算子：正規表現を使った文字列検査
Sub Sample1()

    Dim arr() As String, str As String, pattern As String, msg As String
    
    str = "侍エンジニア塾"
    pattern = "[エンジニア]"
    
    '配列に文字列を1文字ずつ格納
    Dim i As Integer
    ReDim arr(1 To Len(str))
    
    'UBoundは、配列の最大インデックスを返す
    For i = 1 To UBound(arr)
        arr(i) = Mid(str, i, 1)
    Next i
    
    For Each ele In arr
        If ele Like pattern Then
            msg = msg & ele & ", "
        End If
    Next ele
    
    MsgBox msg
End Sub

'Like演算子：b2セルから一列に日付けがある場合、日にちが1から6のセルの右隣に〇を記入
Sub Sample2()
    Dim i As Long
    For i = 2 To Cells(Rows.Count, 2).End(xlUp).Row
        ' Day関数は日付だけを返す
        ' Day関数の他にも、Month関数、Year関数がある
        If Day(Cells(i, 2)) Like "[1-6]" Then Cells(i, 3) = "○"
    Next i
End Sub

'Like演算子を使わないとこうなる
Sub Sample2_2()
    Dim i As Long
    For i = 2 To Cells(Rows.Count, 2).End(xlUp).Row
        Select Case Day(Cells(i, 2))
        Case 1 To 6
            Cells(i, 3) = "○"
        End Select
    Next i
End Sub

Option Explicit

Private Const SHEET_NAME As String = "Sheet1"
Private Const CELL_POS As String = "A2"

Public Sub Main()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)

    Dim tbl As Range
    Set tbl = ws.Range(CELL_POS).CurrentRegion

    Debug.Print "表の範囲 - " & tbl.Address

    Set tbl = tbl.Offset(1, 1).Resize(tbl.Rows.Count - 1, tbl.Columns.Count - 1)

    Debug.Print "表の範囲(リサイズ後) - " & tbl.Address

    Dim r As Range
    Dim c As Range
    For Each r In tbl.Rows
        Debug.Print "-----"
        For Each c In r.Columns
            Debug.Print c.Value
        Next
    Next
End Sub
