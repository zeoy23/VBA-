Attribute VB_Name = "WorksheetUtilities"
' =========================================
' VBA工作表工具模块 (Worksheet Utilities Module)
' =========================================
' 包含工作表操作和管理功能
' Contains worksheet operations and management functions

Option Explicit

' 批量创建工作表 (Batch create worksheets)
Sub CreateMultipleSheets(sheetNames As Variant)
    Dim i As Long
    Dim ws As Worksheet
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    For i = LBound(sheetNames) To UBound(sheetNames)
        ' 检查工作表是否已存在
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(sheetNames(i))
        On Error GoTo 0
        
        If ws Is Nothing Then
            Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            ws.Name = sheetNames(i)
        End If
        
        Set ws = Nothing
    Next i
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "工作表创建完成", vbInformation
End Sub

' 删除空白工作表 (Delete empty worksheets)
Sub DeleteEmptySheets()
    Dim ws As Worksheet
    Dim count As Long
    
    count = 0
    Application.DisplayAlerts = False
    
    For Each ws In ThisWorkbook.Worksheets
        If Application.WorksheetFunction.CountA(ws.UsedRange) = 0 Then
            ws.Delete
            count = count + 1
        End If
    Next ws
    
    Application.DisplayAlerts = True
    MsgBox "已删除 " & count & " 个空白工作表", vbInformation
End Sub

' 复制工作表到新工作簿 (Copy worksheet to new workbook)
Sub CopySheetToNewWorkbook(ws As Worksheet, Optional newFileName As String = "")
    Dim newWb As Workbook
    
    ws.Copy
    Set newWb = ActiveWorkbook
    
    If newFileName <> "" Then
        newWb.SaveAs ThisWorkbook.Path & "\" & newFileName & ".xlsx"
        MsgBox "工作表已复制到新文件: " & newFileName & ".xlsx", vbInformation
    Else
        MsgBox "工作表已复制到新工作簿", vbInformation
    End If
End Sub

' 隐藏所有工作表除了指定工作表 (Hide all sheets except specified)
Sub HideAllExcept(sheetName As String)
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> sheetName Then
            ws.Visible = xlSheetHidden
        Else
            ws.Visible = xlSheetVisible
        End If
    Next ws
    
    MsgBox "已隐藏其他工作表", vbInformation
End Sub

' 显示所有工作表 (Show all worksheets)
Sub ShowAllSheets()
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        ws.Visible = xlSheetVisible
    Next ws
    
    MsgBox "所有工作表已显示", vbInformation
End Sub

' 按名称排序工作表 (Sort worksheets by name)
Sub SortSheetsByName(Optional ascending As Boolean = True)
    Dim i As Long, j As Long
    Dim sheetCount As Long
    
    sheetCount = ThisWorkbook.Worksheets.Count
    
    Application.ScreenUpdating = False
    
    For i = 1 To sheetCount - 1
        For j = i + 1 To sheetCount
            If ascending Then
                If ThisWorkbook.Worksheets(j).Name < ThisWorkbook.Worksheets(i).Name Then
                    ThisWorkbook.Worksheets(j).Move Before:=ThisWorkbook.Worksheets(i)
                End If
            Else
                If ThisWorkbook.Worksheets(j).Name > ThisWorkbook.Worksheets(i).Name Then
                    ThisWorkbook.Worksheets(j).Move Before:=ThisWorkbook.Worksheets(i)
                End If
            End If
        Next j
    Next i
    
    Application.ScreenUpdating = True
    MsgBox "工作表已排序", vbInformation
End Sub

' 保护所有工作表 (Protect all worksheets)
Sub ProtectAllSheets(password As String)
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        ws.Protect password:=password
    Next ws
    
    MsgBox "所有工作表已保护", vbInformation
End Sub

' 取消保护所有工作表 (Unprotect all worksheets)
Sub UnprotectAllSheets(password As String)
    Dim ws As Worksheet
    
    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        ws.Unprotect password:=password
    Next ws
    On Error GoTo 0
    
    MsgBox "所有工作表已取消保护", vbInformation
End Sub

' 为每个工作表添加目录 (Create table of contents)
Sub CreateTableOfContents()
    Dim ws As Worksheet
    Dim tocSheet As Worksheet
    Dim i As Long
    
    ' 删除已存在的目录表
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("目录").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    ' 创建新的目录表
    Set tocSheet = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
    tocSheet.Name = "目录"
    
    ' 设置标题
    With tocSheet.Range("A1")
        .Value = "工作表目录"
        .Font.Size = 16
        .Font.Bold = True
    End With
    
    ' 添加工作表链接
    i = 3
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "目录" Then
            tocSheet.Hyperlinks.Add _
                Anchor:=tocSheet.Cells(i, 1), _
                Address:="", _
                SubAddress:="'" & ws.Name & "'!A1", _
                TextToDisplay:=ws.Name
            i = i + 1
        End If
    Next ws
    
    tocSheet.Columns("A:A").AutoFit
    MsgBox "目录创建完成", vbInformation
End Sub

' 清空所有工作表内容 (Clear all worksheet contents)
Sub ClearAllSheetContents()
    Dim ws As Worksheet
    Dim response As VbMsgBoxResult
    
    response = MsgBox("确定要清空所有工作表的内容吗？此操作不可撤销！", vbYesNo + vbExclamation)
    
    If response = vbYes Then
        Application.ScreenUpdating = False
        
        For Each ws In ThisWorkbook.Worksheets
            ws.Cells.Clear
        Next ws
        
        Application.ScreenUpdating = True
        MsgBox "所有工作表内容已清空", vbInformation
    End If
End Sub

' 获取工作表统计信息 (Get worksheet statistics)
Function GetWorksheetInfo(ws As Worksheet) As String
    Dim info As String
    Dim lastRow As Long, lastCol As Long
    Dim usedCells As Long
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    usedCells = Application.WorksheetFunction.CountA(ws.UsedRange)
    
    info = "工作表信息: " & ws.Name & vbCrLf & _
           "最后使用行: " & lastRow & vbCrLf & _
           "最后使用列: " & lastCol & vbCrLf & _
           "已使用单元格: " & usedCells
    
    GetWorksheetInfo = info
End Function

' 批量重命名工作表 (Batch rename worksheets)
Sub RenameWorksheets(prefix As String)
    Dim ws As Worksheet
    Dim i As Long
    
    i = 1
    Application.ScreenUpdating = False
    
    For Each ws In ThisWorkbook.Worksheets
        ws.Name = prefix & i
        i = i + 1
    Next ws
    
    Application.ScreenUpdating = True
    MsgBox "工作表重命名完成", vbInformation
End Sub
