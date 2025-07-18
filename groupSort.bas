Attribute VB_Name = "groupSort"
Sub ボタン1_Click()
    Call groupSort
End Sub
Private Sub groupSort()
    Dim ws As Worksheet
    Dim fieldRng As Range
    Dim keyRng As Range
    
    ' シートと範囲を指定
    Set ws = ThisWorkbook.Sheets("テスト名簿")
    Set fieldRng = ws.Range("B4:E15")
    Set keyRng = ws.Range("B4")

    ' 並び替え実行
    With ws.Sort
        With .SortFields
            .Clear
            .Add Key:=keyRng, Order:=xlAscending ', DataOption:=xlSortNormal
        End With
        .SetRange fieldRng
        .Header = xlNo ' ヘッダーがない場合は xlNo、ヘッダーがある場合は xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        '.SortMethod = xlPinYin ' 日本語の並び替えに対応
        .Apply
    End With
End Sub
