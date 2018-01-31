Attribute VB_Name = "Module1"
'Like���Z�q�F���K�\�����g���������񌟍�
Sub Sample1()

    Dim arr() As String, str As String, pattern As String, msg As String
    
    str = "���G���W�j�A�m"
    pattern = "[�G���W�j�A]"
    
    '�z��ɕ������1�������i�[
    Dim i As Integer
    ReDim arr(1 To Len(str))
    
    'UBound�́A�z��̍ő�C���f�b�N�X��Ԃ�
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

'Like���Z�q�Fb2�Z��������ɓ��t��������ꍇ�A���ɂ���1����6�̃Z���̉E�ׂɁZ���L��
Sub Sample2()
    Dim i As Long
    For i = 2 To Cells(Rows.Count, 2).End(xlUp).Row
        ' Day�֐��͓��t������Ԃ�
        ' Day�֐��̑��ɂ��AMonth�֐��AYear�֐�������
        If Day(Cells(i, 2)) Like "[1-6]" Then Cells(i, 3) = "��"
    Next i
End Sub

'Like���Z�q���g��Ȃ��Ƃ����Ȃ�
Sub Sample2_2()
    Dim i As Long
    For i = 2 To Cells(Rows.Count, 2).End(xlUp).Row
        Select Case Day(Cells(i, 2))
        Case 1 To 6
            Cells(i, 3) = "��"
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

    Debug.Print "�\�͈̔� - " & tbl.Address

    Set tbl = tbl.Offset(1, 1).Resize(tbl.Rows.Count - 1, tbl.Columns.Count - 1)

    Debug.Print "�\�͈̔�(���T�C�Y��) - " & tbl.Address

    Dim r As Range
    Dim c As Range
    For Each r In tbl.Rows
        Debug.Print "-----"
        For Each c In r.Columns
            Debug.Print c.Value
        Next
    Next
End Sub
