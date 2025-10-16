Attribute VB_Name = "DataAnalysis"
' =========================================
' VBA数据分析模块 (Data Analysis Module)
' =========================================
' 包含数据统计和分析功能
' Contains data statistics and analysis functions

Option Explicit

' 计算范围统计信息 (Calculate range statistics)
Function GetRangeStatistics(rng As Range) As String
    Dim result As String
    Dim count As Long, sum As Double
    Dim avg As Double, minVal As Double, maxVal As Double
    
    On Error Resume Next
    
    count = Application.WorksheetFunction.Count(rng)
    sum = Application.WorksheetFunction.sum(rng)
    avg = Application.WorksheetFunction.Average(rng)
    minVal = Application.WorksheetFunction.Min(rng)
    maxVal = Application.WorksheetFunction.Max(rng)
    
    result = "统计信息:" & vbCrLf & _
             "数量: " & count & vbCrLf & _
             "总和: " & sum & vbCrLf & _
             "平均值: " & avg & vbCrLf & _
             "最小值: " & minVal & vbCrLf & _
             "最大值: " & maxVal
    
    On Error GoTo 0
    GetRangeStatistics = result
End Function

' 创建数据透视表 (Create pivot table)
Sub CreatePivotTable(sourceRange As Range, targetCell As Range, _
                     rowField As String, columnField As String, dataField As String)
    Dim pc As PivotCache
    Dim pt As PivotTable
    
    Application.ScreenUpdating = False
    
    ' 创建数据透视缓存
    Set pc = ActiveWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=sourceRange.Address(True, True, xlR1C1, True))
    
    ' 创建数据透视表
    Set pt = pc.CreatePivotTable( _
        TableDestination:=targetCell.Address(True, True, xlR1C1, True), _
        TableName:="数据透视表_" & Format(Now, "yyyymmddhhnnss"))
    
    ' 设置行字段
    With pt.PivotFields(rowField)
        .Orientation = xlRowField
        .Position = 1
    End With
    
    ' 设置列字段
    If columnField <> "" Then
        With pt.PivotFields(columnField)
            .Orientation = xlColumnField
            .Position = 1
        End With
    End If
    
    ' 设置数据字段
    With pt.PivotFields(dataField)
        .Orientation = xlDataField
        .Function = xlSum
    End With
    
    Application.ScreenUpdating = True
    MsgBox "数据透视表创建完成", vbInformation
End Sub

' 数据分组 (Group data by column)
Function GroupDataByColumn(sourceRange As Range, keyColumn As Long) As Object
    Dim dict As Object
    Dim cell As Range
    Dim key As Variant
    Dim row As Long
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    For Each cell In sourceRange.Columns(keyColumn).Cells
        If Not IsEmpty(cell.Value) Then
            key = cell.Value
            row = cell.Row
            
            If Not dict.Exists(key) Then
                dict.Add key, New Collection
            End If
            
            dict(key).Add row
        End If
    Next cell
    
    Set GroupDataByColumn = dict
End Function

' 计算同比增长率 (Calculate year-over-year growth)
Function CalculateYoYGrowth(currentValue As Double, previousValue As Double) As Double
    If previousValue = 0 Then
        CalculateYoYGrowth = 0
    Else
        CalculateYoYGrowth = (currentValue - previousValue) / previousValue
    End If
End Function

' 查找异常值 (Find outliers using IQR method)
Sub FindOutliers(rng As Range, highlightColor As Long)
    Dim values() As Double
    Dim cell As Range
    Dim i As Long, count As Long
    Dim q1 As Double, q3 As Double, iqr As Double
    Dim lowerBound As Double, upperBound As Double
    
    ' 收集数值
    ReDim values(1 To rng.Cells.Count)
    count = 0
    
    For Each cell In rng
        If IsNumeric(cell.Value) And Not IsEmpty(cell.Value) Then
            count = count + 1
            values(count) = cell.Value
        End If
    Next cell
    
    If count < 4 Then
        MsgBox "数据量不足，无法计算异常值", vbExclamation
        Exit Sub
    End If
    
    ' 排序数组
    ReDim Preserve values(1 To count)
    Call QuickSort(values, 1, count)
    
    ' 计算四分位数
    q1 = values(Application.WorksheetFunction.Quartile(values, 1))
    q3 = values(Application.WorksheetFunction.Quartile(values, 3))
    iqr = q3 - q1
    
    lowerBound = q1 - 1.5 * iqr
    upperBound = q3 + 1.5 * iqr
    
    ' 标记异常值
    Application.ScreenUpdating = False
    
    For Each cell In rng
        If IsNumeric(cell.Value) And Not IsEmpty(cell.Value) Then
            If cell.Value < lowerBound Or cell.Value > upperBound Then
                cell.Interior.Color = highlightColor
            End If
        End If
    Next cell
    
    Application.ScreenUpdating = True
    MsgBox "异常值已标记", vbInformation
End Sub

' 快速排序辅助函数 (Quick sort helper)
Private Sub QuickSort(arr() As Double, left As Long, right As Long)
    Dim pivot As Double
    Dim i As Long, j As Long
    Dim temp As Double
    
    If left < right Then
        pivot = arr((left + right) \ 2)
        i = left
        j = right
        
        Do
            Do While arr(i) < pivot: i = i + 1: Loop
            Do While arr(j) > pivot: j = j - 1: Loop
            
            If i <= j Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
                i = i + 1
                j = j - 1
            End If
        Loop While i <= j
        
        If left < j Then QuickSort arr, left, j
        If i < right Then QuickSort arr, i, right
    End If
End Sub

' 计算移动平均 (Calculate moving average)
Sub CalculateMovingAverage(dataRange As Range, period As Integer, outputRange As Range)
    Dim i As Long
    Dim lastRow As Long
    Dim avg As Double
    
    lastRow = dataRange.Rows.Count
    
    Application.ScreenUpdating = False
    
    For i = period To lastRow
        avg = Application.WorksheetFunction.Average( _
            dataRange.Cells(i - period + 1, 1).Resize(period, 1))
        outputRange.Cells(i, 1).Value = avg
    Next i
    
    Application.ScreenUpdating = True
    MsgBox period & "期移动平均计算完成", vbInformation
End Sub

' 数据去重统计 (Count unique values)
Function CountUnique(rng As Range) As Long
    Dim dict As Object
    Dim cell As Range
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    For Each cell In rng
        If Not IsEmpty(cell.Value) Then
            dict(cell.Value) = 1
        End If
    Next cell
    
    CountUnique = dict.Count
End Function

' 创建频率分布表 (Create frequency distribution)
Sub CreateFrequencyDistribution(dataRange As Range, bins As Variant, outputRange As Range)
    Dim freq As Variant
    Dim i As Long
    
    freq = Application.WorksheetFunction.Frequency(dataRange, bins)
    
    outputRange.Cells(1, 1).Value = "区间"
    outputRange.Cells(1, 2).Value = "频数"
    
    For i = 1 To UBound(bins) + 1
        If i <= UBound(bins) Then
            outputRange.Cells(i + 1, 1).Value = "≤" & bins(i)
        Else
            outputRange.Cells(i + 1, 1).Value = ">" & bins(UBound(bins))
        End If
        outputRange.Cells(i + 1, 2).Value = freq(i, 1)
    Next i
    
    MsgBox "频率分布表创建完成", vbInformation
End Sub

' 数据排名 (Rank data)
Sub RankData(dataRange As Range, outputRange As Range, Optional ascending As Boolean = False)
    Dim cell As Range
    Dim i As Long
    
    Application.ScreenUpdating = False
    
    i = 1
    For Each cell In dataRange
        If Not IsEmpty(cell.Value) Then
            If ascending Then
                outputRange.Cells(i, 1).Value = Application.WorksheetFunction.Rank(cell.Value, dataRange, 1)
            Else
                outputRange.Cells(i, 1).Value = Application.WorksheetFunction.Rank(cell.Value, dataRange, 0)
            End If
        End If
        i = i + 1
    Next cell
    
    Application.ScreenUpdating = True
    MsgBox "数据排名完成", vbInformation
End Sub
