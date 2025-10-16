Attribute VB_Name = "Examples"
' =========================================
' 使用示例 (Usage Examples)
' =========================================
' 这个模块包含各种功能的使用示例
' This module contains usage examples for various features

Option Explicit

' ==================== 数据处理示例 ====================

' 示例：批量清理数据
Sub Example_DataCleaning()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' 去除空白字符
    Call TrimAllCells(ws.Range("A1:E100"))
    
    ' 删除重复行（基于第1列）
    Call RemoveDuplicateRows(ws.Range("A1:E100"), 1)
    
    ' 批量替换
    Call BatchReplace(ws.Range("A1:E100"), "旧文本", "新文本")
End Sub

' 示例：数据提取和转换
Sub Example_DataExtraction()
    Dim ws As Worksheet
    Dim cell As Range
    
    Set ws = ActiveSheet
    
    ' 在B列提取A列中的数字
    For Each cell In ws.Range("A1:A100")
        If Not IsEmpty(cell.Value) Then
            cell.Offset(0, 1).Value = ExtractNumbers(cell.Value)
        End If
    Next cell
    
    ' 在C列提取A列中的中文字符
    For Each cell In ws.Range("A1:A100")
        If Not IsEmpty(cell.Value) Then
            cell.Offset(0, 2).Value = ExtractChinese(cell.Value)
        End If
    Next cell
End Sub

' 示例：数据分列和合并
Sub Example_SplitAndMerge()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' 将A列的数据按逗号分列到B、C、D列
    Call SplitDataToColumns(ws.Range("A1:A10"), ",", 2)
    
    ' 将B、C、D列合并到E列，用" | "分隔
    Dim cell As Range
    For Each cell In ws.Range("A1:A10")
        cell.Offset(0, 4).Value = MergeColumns(cell.Offset(0, 1).Resize(1, 3), " | ")
    Next cell
End Sub

' ==================== 文件操作示例 ====================

' 示例：完整的数据导出流程
Sub Example_ExportWorkflow()
    Dim ws As Worksheet
    Dim exportPath As String
    
    Set ws = ActiveSheet
    exportPath = ThisWorkbook.Path & "\导出数据_" & Format(Now, "yyyymmdd") & ".csv"
    
    ' 先备份
    Call BackupWorkbook
    
    ' 导出CSV
    Call ExportToCSV(ws, exportPath)
End Sub

' 示例：批量处理文件夹中的Excel文件
Sub Example_BatchImport()
    Dim wsTarget As Worksheet
    Dim folderPath As String
    
    ' 创建或获取目标工作表
    On Error Resume Next
    Set wsTarget = ThisWorkbook.Worksheets("汇总数据")
    If wsTarget Is Nothing Then
        Set wsTarget = ThisWorkbook.Worksheets.Add
        wsTarget.Name = "汇总数据"
    End If
    On Error GoTo 0
    
    ' 设置标题行
    wsTarget.Range("A1").Value = "数据列1"
    wsTarget.Range("B1").Value = "数据列2"
    wsTarget.Range("C1").Value = "数据列3"
    
    ' 批量导入
    folderPath = "C:\数据文件夹\"
    Call ImportAllExcelFiles(folderPath, wsTarget)
End Sub

' 示例：按部门拆分数据到多个文件
Sub Example_SplitByDepartment()
    Dim ws As Worksheet
    Dim outputFolder As String
    
    Set ws = ActiveSheet
    outputFolder = ThisWorkbook.Path & "\拆分结果\"
    
    ' 创建输出文件夹（如果不存在）
    On Error Resume Next
    MkDir outputFolder
    On Error GoTo 0
    
    ' 按第3列（部门）拆分
    Call SplitWorkbookByColumn(ws, 3, outputFolder)
End Sub

' ==================== 格式化示例 ====================

' 示例：创建美观的报表
Sub Example_CreateBeautifulReport()
    Dim ws As Worksheet
    Dim headerRange As Range
    Dim dataRange As Range
    
    Set ws = ActiveSheet
    Set headerRange = ws.Range("A1:E1")
    Set dataRange = ws.Range("A1:E20")
    
    ' 创建标准表格格式
    Call CreateStandardTable(dataRange)
    
    ' 隔行填充颜色（白色和浅灰色）
    Call AlternateRowColors(ws.Range("A2:E20"), RGB(255, 255, 255), RGB(242, 242, 242))
    
    ' 设置货币格式（假设D列是金额）
    Call SetCurrencyFormat(ws.Range("D2:D20"), "¥")
    
    ' 设置百分比格式（假设E列是百分比）
    Call SetPercentageFormat(ws.Range("E2:E20"), 2)
    
    MsgBox "报表格式化完成！", vbInformation
End Sub

' 示例：条件格式化
Sub Example_ConditionalFormatting()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' 突出显示重复值（黄色）
    Call HighlightDuplicates(ws.Range("A2:A100"), RGB(255, 255, 0))
    
    ' 对销售额应用颜色渐变
    Call ApplyColorScale(ws.Range("C2:C100"))
End Sub

' ==================== 数据分析示例 ====================

' 示例：完整的数据分析流程
Sub Example_DataAnalysisWorkflow()
    Dim ws As Worksheet
    Dim dataRange As Range
    Dim stats As String
    
    Set ws = ActiveSheet
    Set dataRange = ws.Range("B2:B100")
    
    ' 1. 获取基本统计信息
    stats = GetRangeStatistics(dataRange)
    MsgBox stats, vbInformation, "统计信息"
    
    ' 2. 查找并标记异常值（红色）
    Call FindOutliers(dataRange, RGB(255, 0, 0))
    
    ' 3. 计算移动平均（7期）
    Call CalculateMovingAverage(dataRange, 7, ws.Range("C2"))
    
    ' 4. 计算排名
    Call RankData(dataRange, ws.Range("D2"), False)
    
    MsgBox "数据分析完成！", vbInformation
End Sub

' 示例：创建销售数据透视表
Sub Example_CreateSalesPivot()
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim sourceRange As Range
    
    Set wsSource = ThisWorkbook.Worksheets("销售数据")
    Set wsTarget = ThisWorkbook.Worksheets.Add
    wsTarget.Name = "数据透视分析"
    
    ' 获取源数据范围（假设数据从A1开始）
    Set sourceRange = wsSource.Range("A1").CurrentRegion
    
    ' 创建数据透视表
    ' 参数：源数据、目标位置、行字段、列字段、数据字段
    Call CreatePivotTable(sourceRange, wsTarget.Range("A1"), "产品类别", "月份", "销售额")
End Sub

' 示例：计算同比增长
Sub Example_YoYGrowth()
    Dim ws As Worksheet
    Dim i As Long
    Dim lastRow As Long
    Dim currentYear As Double
    Dim previousYear As Double
    Dim growth As Double
    
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' 假设B列是今年数据，C列是去年数据，D列输出增长率
    ws.Range("D1").Value = "同比增长率"
    
    For i = 2 To lastRow
        currentYear = ws.Cells(i, 2).Value
        previousYear = ws.Cells(i, 3).Value
        growth = CalculateYoYGrowth(currentYear, previousYear)
        ws.Cells(i, 4).Value = growth
    Next i
    
    ' 设置百分比格式
    Call SetPercentageFormat(ws.Range("D2:D" & lastRow), 2)
End Sub

' ==================== 工作表管理示例 ====================

' 示例：创建月度报表结构
Sub Example_CreateMonthlyStructure()
    Dim months As Variant
    
    months = Array("1月", "2月", "3月", "4月", "5月", "6月", _
                   "7月", "8月", "9月", "10月", "11月", "12月")
    
    ' 批量创建月度工作表
    Call CreateMultipleSheets(months)
    
    ' 创建目录
    Call CreateTableOfContents
    
    ' 排序工作表
    Call SortSheetsByName(True)
    
    MsgBox "月度报表结构创建完成！", vbInformation
End Sub

' 示例：工作表清理
Sub Example_WorksheetCleanup()
    Dim response As VbMsgBoxResult
    
    response = MsgBox("是否要清理工作簿？" & vbCrLf & _
                     "- 删除空白工作表" & vbCrLf & _
                     "- 显示所有隐藏的工作表" & vbCrLf & _
                     "- 按名称排序", vbYesNo + vbQuestion)
    
    If response = vbYes Then
        Call DeleteEmptySheets
        Call ShowAllSheets
        Call SortSheetsByName(True)
        MsgBox "清理完成！", vbInformation
    End If
End Sub

' 示例：导出所有工作表为单独文件
Sub Example_ExportAllSheets()
    Dim ws As Worksheet
    Dim outputFolder As String
    
    outputFolder = ThisWorkbook.Path & "\各工作表\"
    
    ' 创建输出文件夹
    On Error Resume Next
    MkDir outputFolder
    On Error GoTo 0
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    For Each ws In ThisWorkbook.Worksheets
        Call CopySheetToNewWorkbook(ws, outputFolder & ws.Name)
    Next ws
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox "所有工作表已导出！", vbInformation
End Sub

' ==================== 综合应用示例 ====================

' 示例：完整的数据处理和报表生成流程
Sub Example_CompleteWorkflow()
    Dim wsRaw As Worksheet
    Dim wsClean As Worksheet
    Dim wsReport As Worksheet
    
    Application.ScreenUpdating = False
    
    ' 1. 准备工作表
    Set wsRaw = ThisWorkbook.Worksheets("原始数据")
    
    On Error Resume Next
    Set wsClean = ThisWorkbook.Worksheets("清洗后数据")
    If wsClean Is Nothing Then
        Set wsClean = ThisWorkbook.Worksheets.Add
        wsClean.Name = "清洗后数据"
    End If
    
    Set wsReport = ThisWorkbook.Worksheets("分析报表")
    If wsReport Is Nothing Then
        Set wsReport = ThisWorkbook.Worksheets.Add
        wsReport.Name = "分析报表"
    End If
    On Error GoTo 0
    
    ' 2. 复制原始数据
    wsRaw.UsedRange.Copy wsClean.Range("A1")
    
    ' 3. 数据清洗
    Call TrimAllCells(wsClean.UsedRange)
    Call RemoveDuplicateRows(wsClean.UsedRange, 1)
    
    ' 4. 格式化清洗后的数据
    Call CreateStandardTable(wsClean.UsedRange)
    Call AlternateRowColors(wsClean.Range("A2:Z100"), RGB(255, 255, 255), RGB(242, 242, 242))
    
    ' 5. 创建数据透视表到报表工作表
    Call CreatePivotTable(wsClean.UsedRange, wsReport.Range("A1"), "类别", "月份", "金额")
    
    ' 6. 添加统计信息
    Dim stats As String
    stats = GetRangeStatistics(wsClean.Range("D2:D100"))
    wsReport.Range("H1").Value = stats
    
    ' 7. 格式化报表
    wsReport.Columns.AutoFit
    
    Application.ScreenUpdating = True
    
    MsgBox "完整数据处理流程执行完成！" & vbCrLf & _
           "请查看'清洗后数据'和'分析报表'工作表", vbInformation
End Sub

' 示例：自动化每日报表生成
Sub Example_DailyReportGeneration()
    Dim todaySheet As Worksheet
    Dim reportName As String
    
    reportName = "报表_" & Format(Date, "yyyy-mm-dd")
    
    ' 检查今日报表是否已存在
    On Error Resume Next
    Set todaySheet = ThisWorkbook.Worksheets(reportName)
    On Error GoTo 0
    
    If todaySheet Is Nothing Then
        ' 创建新报表
        Set todaySheet = ThisWorkbook.Worksheets.Add
        todaySheet.Name = reportName
        
        ' 设置报表结构
        With todaySheet
            .Range("A1").Value = "日期"
            .Range("B1").Value = "项目"
            .Range("C1").Value = "数量"
            .Range("D1").Value = "金额"
            .Range("E1").Value = "备注"
        End With
        
        ' 应用格式
        Call CreateStandardTable(todaySheet.Range("A1:E1"))
        todaySheet.Columns.AutoFit
        
        MsgBox "今日报表已创建：" & reportName, vbInformation
    Else
        MsgBox "今日报表已存在：" & reportName, vbInformation
        todaySheet.Activate
    End If
End Sub
