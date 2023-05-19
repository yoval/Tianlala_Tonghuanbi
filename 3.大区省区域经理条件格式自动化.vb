
Sub SetConditionalFormat()
    Dim rng As Range
    Dim col As Variant '修改这里
    Dim lastVal As Variant
    '获取已有选区范围
    Set rng = Selection
    '定义一个数组，存储要作用的列号
    Dim colsArray(1 To 4) As Variant '修改这里
    colsArray(1) = 9
    colsArray(2) = 12
    colsArray(3) = 15
    colsArray(4) = 18
    '遍历数组中的每个列号
    For Each col In colsArray
        '获取最下方的值
        lastVal = rng.Cells(rng.Rows.Count, col).Value
        '如果是数字，则设置条件格式
        If IsNumeric(lastVal) Then
            '添加条件格式规则，如果小于最下方的值，则字体为红色
            rng.Columns(col).FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:=lastVal
            rng.Columns(col).FormatConditions(1).Font.Color = vbRed
        End If
    Next col
End Sub

