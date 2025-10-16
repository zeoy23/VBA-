Attribute VB_Name = "FileOperations"
' =========================================
' VBA文件操作模块 (File Operations Module)
' =========================================
' 包含文件导入导出和工作簿操作功能
' Contains file import/export and workbook operations

Option Explicit

' 导出工作表为CSV (Export worksheet to CSV)
Sub ExportToCSV(ws As Worksheet, filePath As String)
    Dim fileNum As Integer
    Dim row As Long, col As Long
    Dim rowData As String
    Dim lastRow As Long, lastCol As Long
    
    fileNum = FreeFile
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    Open filePath For Output As #fileNum
    
    For row = 1 To lastRow
        rowData = ""
        For col = 1 To lastCol
            If col > 1 Then rowData = rowData & ","
            rowData = rowData & Chr(34) & ws.Cells(row, col).Value & Chr(34)
        Next col
        Print #fileNum, rowData
    Next row
    
    Close #fileNum
    MsgBox "导出CSV成功: " & filePath, vbInformation
End Sub

' 导入CSV文件 (Import CSV file)
Sub ImportFromCSV(ws As Worksheet, filePath As String, Optional hasHeader As Boolean = True)
    Dim fileNum As Integer
    Dim lineData As String
    Dim dataArray As Variant
    Dim row As Long
    Dim col As Long
    
    fileNum = FreeFile
    row = 1
    
    ws.Cells.Clear
    
    Open filePath For Input As #fileNum
    
    Do While Not EOF(fileNum)
        Line Input #fileNum, lineData
        dataArray = Split(lineData, ",")
        
        For col = 0 To UBound(dataArray)
            ws.Cells(row, col + 1).Value = Replace(dataArray(col), Chr(34), "")
        Next col
        
        row = row + 1
    Loop
    
    Close #fileNum
    
    If hasHeader Then
        ws.Rows(1).Font.Bold = True
    End If
    
    MsgBox "导入CSV成功: " & filePath, vbInformation
End Sub

' 批量导入文件夹中的所有Excel文件 (Import all Excel files from folder)
Sub ImportAllExcelFiles(folderPath As String, targetSheet As Worksheet)
    Dim fileName As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim targetRow As Long
    Dim lastRow As Long
    
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    
    fileName = Dir(folderPath & "*.xlsx")
    targetRow = 2 ' 从第2行开始，假设第1行是标题
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Do While fileName <> ""
        Set wb = Workbooks.Open(folderPath & fileName)
        
        For Each ws In wb.Worksheets
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            
            If lastRow > 1 Then
                ws.Range("A2:Z" & lastRow).Copy
                targetSheet.Cells(targetRow, 1).PasteSpecial xlPasteValues
                targetRow = targetRow + lastRow - 1
            End If
        Next ws
        
        wb.Close SaveChanges:=False
        fileName = Dir
    Loop
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "批量导入完成", vbInformation
End Sub

' 拆分工作簿 (Split workbook by column value)
Sub SplitWorkbookByColumn(sourceSheet As Worksheet, keyColumn As Long, outputFolder As String)
    Dim dict As Object
    Dim cell As Range
    Dim key As Variant
    Dim newWb As Workbook
    Dim newWs As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim rng As Range
    Dim i As Long
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    lastRow = sourceSheet.Cells(sourceSheet.Rows.Count, keyColumn).End(xlUp).Row
    lastCol = sourceSheet.Cells(1, sourceSheet.Columns.Count).End(xlToLeft).Column
    
    If Right(outputFolder, 1) <> "\" Then outputFolder = outputFolder & "\"
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' 复制标题行
    sourceSheet.Rows(1).Copy
    
    ' 遍历数据行
    For i = 2 To lastRow
        key = sourceSheet.Cells(i, keyColumn).Value
        
        If Not dict.Exists(key) Then
            Set newWb = Workbooks.Add
            Set newWs = newWb.Sheets(1)
            newWs.Name = Left(CStr(key), 31) ' 工作表名称最多31个字符
            
            ' 粘贴标题行
            newWs.Cells(1, 1).PasteSpecial xlPasteAll
            
            dict.Add key, newWb
        Else
            Set newWb = dict(key)
            Set newWs = newWb.Sheets(1)
        End If
        
        ' 复制当前行
        sourceSheet.Rows(i).Copy
        newWs.Cells(newWs.Cells(newWs.Rows.Count, 1).End(xlUp).Row + 1, 1).PasteSpecial xlPasteAll
    Next i
    
    ' 保存所有新工作簿
    For Each key In dict.Keys
        Set newWb = dict(key)
        newWb.SaveAs outputFolder & CStr(key) & ".xlsx"
        newWb.Close
    Next
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "已创建 " & dict.Count & " 个文件", vbInformation
End Sub

' 合并多个工作簿 (Merge multiple workbooks)
Sub MergeWorkbooks(folderPath As String, targetWorkbook As Workbook)
    Dim fileName As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim targetWs As Worksheet
    
    Set targetWs = targetWorkbook.Sheets.Add
    targetWs.Name = "合并数据"
    
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    
    fileName = Dir(folderPath & "*.xlsx")
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Do While fileName <> ""
        Set wb = Workbooks.Open(folderPath & fileName)
        
        For Each ws In wb.Worksheets
            ws.Copy After:=targetWs
        Next ws
        
        wb.Close SaveChanges:=False
        fileName = Dir
    Loop
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "合并完成", vbInformation
End Sub

' 备份当前工作簿 (Backup current workbook)
Sub BackupWorkbook()
    Dim backupPath As String
    Dim originalName As String
    
    originalName = ThisWorkbook.Name
    backupPath = ThisWorkbook.Path & "\Backup_" & _
                Format(Now, "yyyymmdd_hhnnss") & "_" & originalName
    
    ThisWorkbook.SaveCopyAs backupPath
    MsgBox "备份成功: " & backupPath, vbInformation
End Sub
