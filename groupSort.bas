Attribute VB_Name = "groupSort"
Sub �{�^��1_Click()
    Call groupSort
End Sub
Private Sub groupSort()
    Dim ws As Worksheet
    Dim fieldRng As Range
    Dim keyRng As Range
    
    ' �V�[�g�Ɣ͈͂��w��
    Set ws = ThisWorkbook.Sheets("�e�X�g����")
    Set fieldRng = ws.Range("B4:E15")
    Set keyRng = ws.Range("B4")

    ' ���ёւ����s
    With ws.Sort
        With .SortFields
            .Clear
            .Add Key:=keyRng, Order:=xlAscending ', DataOption:=xlSortNormal
        End With
        .SetRange fieldRng
        .Header = xlNo ' �w�b�_�[���Ȃ��ꍇ�� xlNo�A�w�b�_�[������ꍇ�� xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        '.SortMethod = xlPinYin ' ���{��̕��ёւ��ɑΉ�
        .Apply
    End With
End Sub
