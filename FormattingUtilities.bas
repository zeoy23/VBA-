Attribute VB_Name = "FormattingUtilities"
' =========================================
' VBA格式化工具模块 (Formatting Utilities Module)
' =========================================
' 包含单元格格式化和样式设置功能
' Contains cell formatting and styling functions

Option Explicit

' 自动调整列宽 (Auto-fit columns)
Sub AutoFitColumns(ws As Worksheet)
    ws.Columns.AutoFit
    MsgBox "列宽已自动调整", vbInformation
End Sub

' 设置表格边框 (Add table borders)
Sub AddTableBorders(rng As Range, Optional borderStyle As XlLineStyle = xlContinuous)
    With rng.Borders
        .LineStyle = borderStyle
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    MsgBox "已添加表格边框", vbInformation
End Sub

' 隔行填充颜色 (Alternate row colors)
Sub AlternateRowColors(rng As Range, color1 As Long, color2 As Long)
    Dim row As Range
    Dim i As Long
    
    Application.ScreenUpdating = False
    
    i = 0
    For Each row In rng.Rows
        If i Mod 2 = 0 Then
            row.Interior.Color = color1
        Else
            row.Interior.Color = color2
        End If
        i = i + 1
    Next row
    
    Application.ScreenUpdating = True
    MsgBox "隔行颜色设置完成", vbInformation
End Sub

' 创建标准表格样式 (Create standard table style)
Sub CreateStandardTable(rng As Range)
    With rng
        ' 添加边框
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        
        ' 设置标题行
        With .Rows(1)
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(68, 114, 196)
            .HorizontalAlignment = xlCenter
        End With
        
        ' 自动筛选
        .AutoFilter
        
        ' 冻结首行
        rng.Worksheet.Activate
        rng.Worksheet.Rows(2).Select
        ActiveWindow.FreezePanes = True
    End With
    
    MsgBox "标准表格创建完成", vbInformation
End Sub

' 条件格式：突出显示重复值 (Conditional formatting: highlight duplicates)
Sub HighlightDuplicates(rng As Range, highlightColor As Long)
    Dim cell As Range
    Dim dict As Object
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    Application.ScreenUpdating = False
    
    ' 清除现有格式
    rng.Interior.ColorIndex = xlNone
    
    ' 找出重复值
    For Each cell In rng
        If Not IsEmpty(cell.Value) Then
            If dict.Exists(cell.Value) Then
                cell.Interior.Color = highlightColor
                dict(cell.Value).Interior.Color = highlightColor
            Else
                dict.Add cell.Value, cell
            End If
        End If
    Next cell
    
    Application.ScreenUpdating = True
    MsgBox "重复值已标记", vbInformation
End Sub

' 根据数值设置颜色渐变 (Color scale based on values)
Sub ApplyColorScale(rng As Range)
    Dim minVal As Double, maxVal As Double
    Dim cell As Range
    Dim ratio As Double
    
    On Error Resume Next
    minVal = Application.WorksheetFunction.Min(rng)
    maxVal = Application.WorksheetFunction.Max(rng)
    On Error GoTo 0
    
    If minVal = maxVal Then Exit Sub
    
    Application.ScreenUpdating = False
    
    For Each cell In rng
        If IsNumeric(cell.Value) And Not IsEmpty(cell.Value) Then
            ratio = (cell.Value - minVal) / (maxVal - minVal)
            
            ' 从红色渐变到绿色
            If ratio < 0.5 Then
                cell.Interior.Color = RGB(255, CInt(ratio * 2 * 255), 0)
            Else
                cell.Interior.Color = RGB(CInt((1 - ratio) * 2 * 255), 255, 0)
            End If
        End If
    Next cell
    
    Application.ScreenUpdating = True
    MsgBox "颜色渐变已应用", vbInformation
End Sub

' 清除所有格式 (Clear all formatting)
Sub ClearAllFormatting(rng As Range)
    rng.ClearFormats
    MsgBox "格式已清除", vbInformation
End Sub

' 设置数字格式 (Set number format)
Sub SetNumberFormat(rng As Range, formatString As String)
    On Error Resume Next
    rng.NumberFormat = formatString
    If Err.Number = 0 Then
        MsgBox "数字格式已设置", vbInformation
    Else
        MsgBox "格式字符串无效", vbCritical
    End If
    On Error GoTo 0
End Sub

' 设置货币格式 (Set currency format)
Sub SetCurrencyFormat(rng As Range, Optional currencySymbol As String = "¥")
    rng.NumberFormat = currencySymbol & "#,##0.00"
    MsgBox "货币格式已设置", vbInformation
End Sub

' 设置百分比格式 (Set percentage format)
Sub SetPercentageFormat(rng As Range, Optional decimalPlaces As Integer = 2)
    rng.NumberFormat = "0." & String(decimalPlaces, "0") & "%"
    MsgBox "百分比格式已设置", vbInformation
End Sub

' 设置日期格式 (Set date format)
Sub SetDateFormat(rng As Range, Optional dateFormat As String = "yyyy-mm-dd")
    rng.NumberFormat = dateFormat
    MsgBox "日期格式已设置", vbInformation
End Sub

' 合并单元格并居中 (Merge and center cells)
Sub MergeAndCenter(rng As Range)
    With rng
        .Merge
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    MsgBox "单元格已合并并居中", vbInformation
End Sub

' 批量设置字体 (Batch set font)
Sub SetFont(rng As Range, fontName As String, fontSize As Integer, Optional isBold As Boolean = False)
    With rng.Font
        .Name = fontName
        .Size = fontSize
        .Bold = isBold
    End With
    MsgBox "字体已设置", vbInformation
End Sub

' 设置文本对齐方式 (Set text alignment)
Sub SetAlignment(rng As Range, hAlign As XlHAlign, vAlign As XlVAlign)
    With rng
        .HorizontalAlignment = hAlign
        .VerticalAlignment = vAlign
    End With
    MsgBox "对齐方式已设置", vbInformation
End Sub
