Option Explicit

'==========================================================
' 功能1：模糊筛选 (支持多关键词包含逻辑)
' 逻辑：输入 "A,B"，会筛选出包含 "A" 或者包含 "B" 的所有行
'==========================================================
Sub MultiFilterFuzzy()
    Dim ws As Worksheet
    Dim rng As Range, dataRng As Range
    Dim colNum As Integer
    Dim inputStr As String
    Dim keywords() As String
    Dim cellData As Variant
    Dim i As Long, j As Long
    Dim dict As Object
    Dim matchArr() As String
    Dim matchCount As Long
    Dim cellVal As String
    
    Set ws = ActiveSheet
    
    ' 1. 检查选中
    If ActiveCell Is Nothing Then
        MsgBox "请点击要筛选的那一列中的任意数据！"
        Exit Sub
    End If
    
    ' 2. 自动开启筛选并获取区域
    If Not ws.AutoFilterMode Then
        On Error Resume Next
        ActiveCell.CurrentRegion.AutoFilter
        On Error GoTo 0
    End If
    
    If ws.AutoFilterMode Then
        Set rng = ws.AutoFilter.Range
        ' 检查点击位置是否有效
        If Intersect(ActiveCell, rng) Is Nothing Then
            MsgBox "请点击筛选区域内的单元格！"
            Exit Sub
        End If
        ' 计算相对列号
        colNum = ActiveCell.Column - rng.Column + 1
        ' 获取该列数据区域（不含表头）
        Set dataRng = rng.Columns(colNum).Offset(1, 0).Resize(rng.Rows.Count - 1)
    Else
        MsgBox "无法确定筛选区域，请手动对表头开启筛选。"
        Exit Sub
    End If
    
    ' 3. 获取输入
    inputStr = InputBox("请输入包含的关键词（支持多个）：" & vbCrLf & _
                        "例如输入：A,B" & vbCrLf & _
                        "将筛选出所有【包含】A 或 【包含】B ", "模糊筛选")
    If Trim(inputStr) = "" Then Exit Sub
    
    ' 处理全角逗号
    inputStr = Replace(inputStr, "，", ",")
    keywords = Split(inputStr, ",")
    
    ' 4. 核心算法：在内存中计算所有符合模糊条件的值
    ' 使用字典去重
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' 将数据读入内存数组加速处理
    If dataRng.Cells.Count = 1 Then
        ReDim cellData(1 To 1, 1 To 1)
        cellData(1, 1) = dataRng.Value
    Else
        cellData = dataRng.Value
    End If
    
    On Error Resume Next ' 防止数据中有 #N/A 导致崩溃
    For i = 1 To UBound(cellData)
        cellVal = CStr(cellData(i, 1)) ' 强制转为字符串
        
        If Len(cellVal) > 0 Then
            For j = LBound(keywords) To UBound(keywords)
                ' InStr > 0 表示包含
                If InStr(1, cellVal, Trim(keywords(j)), vbTextCompare) > 0 Then
                    dict(cellVal) = 1 ' 记录符合条件的完整内容
                    Exit For ' 命中一个关键词即可跳出
                End If
            Next j
        End If
    Next i
    On Error GoTo 0
    
    ' 5. 应用筛选
    If dict.Count > 0 Then
        ' 将字典的 Key 转换为数组
        matchArr = Split(Join(dict.Keys, "|@|"), "|@|")
        rng.AutoFilter Field:=colNum, Criteria1:=matchArr, Operator:=xlFilterValues
        MsgBox "筛选完成，找到 " & dict.Count & " 个符合条件的项。"
    Else
        MsgBox "没有找到包含这些关键词的内容。"
        ' 清除该列筛选
        rng.AutoFilter Field:=colNum
    End If
End Sub
