Option Explicit

'=============================================================================
' 脚本名称：追加式矩阵导入工具
'=============================================================================

Sub ImagesMatrixImport()
    Dim ws As Worksheet
    Dim fDialog As FileDialog
    Dim folderPath As String, fileName As String
    Dim rawName As String
    Dim parts() As String
    
    ' --- 配置变量 ---
    Dim configStr As String
    Dim configArr() As String
    Dim delimiter As String      ' 分隔符
    Dim colKeyIdx As Integer     ' 列名索引
    Dim cellHeight As Double     ' 行高
    Dim cellWidth As Double      ' 列宽
    Dim imgMargin As Double      ' 边距
    
    ' --- 逻辑变量 ---
    Dim anchorRow As Long, anchorCol As Long ' 基点坐标
    Dim rowKey As String, colKey As String
    Dim findRow As Range, findCol As Range
    Dim lastRow As Long, lastCol As Long
    Dim targetCell As Range
    Dim shp As Shape
    Dim targetW As Double, targetH As Double
    Dim i As Integer, realColIdx As Integer
    Dim searchRange As Range
    
    '=========================================================================
    ' 第一步：单界面获取配置 (记忆上次输入)
    '=========================================================================
    Dim defaultConf As String
    ' 默认配置：下划线分隔, 列名在最后(-1), 行高100, 列宽50
    defaultConf = GetSetting("ExcelMacro", "AppendMatrix", "Config", "_,-1,100,50")
    
    configStr = InputBox("请输入配置参数（逗号隔开）：" & vbCrLf & vbCrLf & _
                         "1. 分隔符 (如 _ )" & vbCrLf & _
                         "2. 列名索引 (0=开头, -1=结尾)" & vbCrLf & _
                         "3. 行高 (建议 100)" & vbCrLf & _
                         "4. 列宽 (建议 50, 注意单位是字符宽)" & vbCrLf & vbCrLf & _
                         "当前配置：", "追加导入配置", defaultConf)
    
    If Trim(configStr) = "" Then Exit Sub
    
    On Error Resume Next
    configArr = Split(Replace(configStr, "，", ","), ",")
    If UBound(configArr) < 3 Then
        MsgBox "参数数量不足4个，请检查格式。"
        Exit Sub
    End If
    
    delimiter = Trim(configArr(0))
    colKeyIdx = CInt(Trim(configArr(1)))
    cellHeight = CDbl(Trim(configArr(2)))
    cellWidth = CDbl(Trim(configArr(3)))
    imgMargin = 4
    
    If Err.Number <> 0 Then
        MsgBox "参数格式错误（高度/宽度必须是数字）。"
        Exit Sub
    End If
    On Error GoTo 0
    SaveSetting "ExcelMacro", "AppendMatrix", "Config", configStr
    
    '=========================================================================
    ' 第二步：确定基准点 (锚点)
    '=========================================================================
    Set ws = ActiveSheet
    
    ' 以当前选中的单元格作为“表格左上角”
    anchorRow = ActiveCell.Row
    anchorCol = ActiveCell.Column
    
    ' 给基准点打个标记（可选）
    If ws.Cells(anchorRow, anchorCol).Value = "" Then
        ws.Cells(anchorRow, anchorCol).Value = "分类\名称"
    End If
    
    '=========================================================================
    ' 第三步：选择文件夹
    '=========================================================================
    Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)
    fDialog.Title = "请选择图片所在的文件夹"
    If fDialog.Show = -1 Then
        folderPath = fDialog.SelectedItems(1) & "\"
    Else
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    fileName = Dir(folderPath & "*.*")
    
    '=========================================================================
    ' 第四步：循环处理
    '=========================================================================
    Do While fileName <> ""
        If IsImageFile(fileName) Then
            
            rawName = Left(fileName, InStrRev(fileName, ".") - 1)
            parts = Split(rawName, delimiter)
            
            If UBound(parts) >= 0 Then
                
                ' 1. 确定列名位置
                If colKeyIdx < 0 Then
                    realColIdx = UBound(parts)
                Else
                    realColIdx = colKeyIdx
                    If realColIdx > UBound(parts) Then realColIdx = UBound(parts)
                End If
                
                ' 2. 提取列名
                colKey = Trim(parts(realColIdx))
                
                ' 3. 提取行名 (使用用户分隔符拼接)
                rowKey = ""
                For i = 0 To UBound(parts)
                    If i <> realColIdx Then
                        ' 如果不是第一个词，先加分隔符
                        If rowKey <> "" Then rowKey = rowKey & delimiter
                        rowKey = rowKey & parts(i)
                    End If
                Next i
                rowKey = Trim(rowKey)
                If rowKey = "" Then rowKey = "通用"
                
                ' ---------------------------------------------------------
                ' 4. 基于锚点查找行和列
                ' ---------------------------------------------------------
                
                ' A. 找行 (在锚点所在的列向下找)
                ' 搜索范围：从锚点下一行开始，到表格最底部
                Set searchRange = ws.Range(ws.Cells(anchorRow + 1, anchorCol), ws.Cells(ws.Rows.Count, anchorCol))
                ' 为了防止SearchRange太大导致性能问题，可以限制一下，但在追加模式下全列搜比较稳
                
                Set findRow = searchRange.Find(What:=rowKey, LookAt:=xlWhole)
                
                If findRow Is Nothing Then
                    ' 没找到，在当前该列已有数据的最后一行下面追加
                    lastRow = ws.Cells(ws.Rows.Count, anchorCol).End(xlUp).Row
                    ' 确保不小于锚点位置
                    If lastRow < anchorRow Then lastRow = anchorRow
                    
                    Set findRow = ws.Cells(lastRow + 1, anchorCol)
                    findRow.Value = rowKey
                    findRow.RowHeight = cellHeight
                Else
                    ' 如果找到了，也要确保行高符合配置（防止旧数据行高不对）
                    findRow.RowHeight = cellHeight
                End If
                
                ' B. 找列 (在锚点所在的行向右找)
                ' 搜索范围：从锚点右侧一格开始，到最右边
                Set searchRange = ws.Range(ws.Cells(anchorRow, anchorCol + 1), ws.Cells(anchorRow, ws.Columns.Count))
                
                Set findCol = searchRange.Find(What:=colKey, LookAt:=xlWhole)
                
                If findCol Is Nothing Then
                    ' 没找到，在当前该行已有数据的最后一列右边追加
                    lastCol = ws.Cells(anchorRow, ws.Columns.Count).End(xlToLeft).Column
                    If lastCol < anchorCol Then lastCol = anchorCol
                    
                    Set findCol = ws.Cells(anchorRow, lastCol + 1)
                    findCol.Value = colKey
                    ' 【重要修复】立即设置该列的列宽
                    findCol.EntireColumn.ColumnWidth = cellWidth
                Else
                     ' 如果找到了，也强制刷新列宽
                    findCol.EntireColumn.ColumnWidth = cellWidth
                End If
                
                ' ---------------------------------------------------------
                ' 5. 插入图片
                ' ---------------------------------------------------------
                Set targetCell = ws.Cells(findRow.Row, findCol.Column)
                Call DeletePicInCell(targetCell)
                
                Set shp = ws.Shapes.AddPicture( _
                    Filename:=folderPath & fileName, _
                    LinkToFile:=msoFalse, _
                    SaveWithDocument:=msoTrue, _
                    Left:=targetCell.Left, _
                    Top:=targetCell.Top, _
                    Width:=-1, Height:=-1)
                
                With shp
                    .LockAspectRatio = msoTrue
                    
                    ' 计算尺寸
                    targetW = targetCell.Width - (imgMargin * 2)
                    targetH = targetCell.Height - (imgMargin * 2)
                    
                    If targetW > 0 And targetH > 0 Then
                        .Width = targetW
                        If .Height > targetH Then .Height = targetH
                    End If
                    
                    ' 居中
                    .Left = targetCell.Left + (targetCell.Width - .Width) / 2
                    .Top = targetCell.Top + (targetCell.Height - .Height) / 2
                    
                    .Placement = xlMoveAndSize
                End With
                
            End If
        End If
        fileName = Dir
    Loop
    
    ' 只自动调整第一列（行标题）的宽度，不动后面放图片的列
    ws.Columns(anchorCol).AutoFit
    If ws.Columns(anchorCol).ColumnWidth < 20 Then ws.Columns(anchorCol).ColumnWidth = 20
    
    Application.ScreenUpdating = True
    MsgBox "导入完成！" & vbCrLf & "图片已追加到 " & ws.Cells(anchorRow, anchorCol).Address(0, 0) & " 区域。"
End Sub

' --- 辅助函数 ---
Function IsImageFile(fName As String) As Boolean
    Dim ext As String
    ext = LCase(Right(fName, 4))
    If ext = ".jpg" Or ext = ".png" Or ext = ".bmp" Or ext = ".gif" Or Right(LCase(fName), 5) = ".jpeg" Then
        IsImageFile = True
    Else
        IsImageFile = False
    End If
End Function

Sub DeletePicInCell(rng As Range)
    Dim shp As Shape
    For Each shp In rng.Parent.Shapes
        If shp.Type = msoPicture Or shp.Type = msoLinkedPicture Then
            If Not Intersect(shp.TopLeftCell, rng) Is Nothing Then shp.Delete
        End If
    Next shp
End Sub