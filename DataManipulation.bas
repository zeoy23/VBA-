Attribute VB_Name = "DataManipulation"
' =========================================
' VBA数据处理模块 (Data Manipulation Module)
' =========================================
' 包含各种常用的数据处理功能
' Contains various common data processing functions

Option Explicit

' 删除重复行 (Remove duplicate rows)
' 从指定范围删除重复数据
Sub RemoveDuplicateRows(rng As Range, Optional keyColumn As Long = 1)
    Dim dict As Object
    Dim cell As Range
    Dim key As Variant
    Dim rowsToDelete As Collection
    Dim i As Long
    
    Set dict = CreateObject("Scripting.Dictionary")
    Set rowsToDelete = New Collection
    
    ' 遍历范围并标记重复行
    For Each cell In rng.Columns(keyColumn).Cells
        If Not IsEmpty(cell.Value) Then
            key = cell.Value
            If dict.Exists(key) Then
                rowsToDelete.Add cell.Row
            Else
                dict.Add key, cell.Row
            End If
        End If
    Next cell
    
    ' 从后往前删除行，避免索引问题
    For i = rowsToDelete.Count To 1 Step -1
        Rows(rowsToDelete(i)).Delete
    Next i
    
    MsgBox "已删除 " & rowsToDelete.Count & " 个重复行", vbInformation
End Sub

' 数据清洗：去除空白字符 (Data cleaning: trim whitespace)
Sub TrimAllCells(rng As Range)
    Dim cell As Range
    Dim count As Long
    
    count = 0
    Application.ScreenUpdating = False
    
    For Each cell In rng
        If Not IsEmpty(cell.Value) And VarType(cell.Value) = vbString Then
            If cell.Value <> Trim(cell.Value) Then
                cell.Value = Trim(cell.Value)
                count = count + 1
            End If
        End If
    Next cell
    
    Application.ScreenUpdating = True
    MsgBox "已清理 " & count & " 个单元格", vbInformation
End Sub

' 批量替换 (Batch replace)
Sub BatchReplace(rng As Range, findText As String, replaceText As String)
    Dim count As Long
    
    Application.ScreenUpdating = False
    
    count = rng.Replace(What:=findText, _
                       Replacement:=replaceText, _
                       LookAt:=xlPart, _
                       SearchOrder:=xlByRows, _
                       MatchCase:=False)
    
    Application.ScreenUpdating = True
    MsgBox "替换了 " & count & " 处内容", vbInformation
End Sub

' 分列数据 (Split data to columns)
Sub SplitDataToColumns(sourceRange As Range, delimiter As String, targetColumn As Long)
    Dim arr As Variant
    Dim i As Long, j As Long
    Dim cell As Range
    Dim ws As Worksheet
    
    Set ws = sourceRange.Worksheet
    Application.ScreenUpdating = False
    
    For Each cell In sourceRange
        If Not IsEmpty(cell.Value) Then
            arr = Split(cell.Value, delimiter)
            For j = 0 To UBound(arr)
                ws.Cells(cell.Row, targetColumn + j).Value = Trim(arr(j))
            Next j
        End If
    Next cell
    
    Application.ScreenUpdating = True
    MsgBox "分列完成", vbInformation
End Sub

' 合并列数据 (Merge column data)
Function MergeColumns(rng As Range, delimiter As String) As String
    Dim cell As Range
    Dim result As String
    
    result = ""
    For Each cell In rng
        If Not IsEmpty(cell.Value) Then
            If result <> "" Then
                result = result & delimiter
            End If
            result = result & cell.Value
        End If
    Next cell
    
    MergeColumns = result
End Function

' 提取数字 (Extract numbers from text)
Function ExtractNumbers(text As String) As String
    Dim i As Long
    Dim result As String
    Dim char As String
    
    result = ""
    For i = 1 To Len(text)
        char = Mid(text, i, 1)
        If IsNumeric(char) Or char = "." Or char = "-" Then
            result = result & char
        End If
    Next i
    
    ExtractNumbers = result
End Function

' 提取中文字符 (Extract Chinese characters)
Function ExtractChinese(text As String) As String
    Dim i As Long
    Dim result As String
    Dim char As String
    
    result = ""
    For i = 1 To Len(text)
        char = Mid(text, i, 1)
        If Asc(char) < 0 Then ' 中文字符的ASCII值为负数
            result = result & char
        End If
    Next i
    
    ExtractChinese = result
End Function

' 数据转置 (Transpose data)
Sub TransposeData(sourceRange As Range, targetCell As Range)
    Dim transposed As Variant
    
    Application.ScreenUpdating = False
    
    transposed = Application.WorksheetFunction.Transpose(sourceRange.Value)
    targetCell.Resize(UBound(transposed, 1), UBound(transposed, 2)).Value = transposed
    
    Application.ScreenUpdating = True
    MsgBox "转置完成", vbInformation
End Sub
