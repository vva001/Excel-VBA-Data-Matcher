Sub AdvancedDataMatcher()
    '------------------------------------------------------------------------------------
    ' 功能:   根据主表与副表中的<关键字列>，将副表中的多列数据高效匹配并填充到主表中。
    ' 作者:   (你的名字或昵称)
    ' 日期:   2025-06-30
    ' 说明:   本脚本使用字典对象(Scripting.Dictionary)进行内存匹配，性能远超VLOOKUP，
    '         尤其适用于大数据量处理。用户可通过修改下方的“配置参数”部分轻松适配自己的表格。
    '------------------------------------------------------------------------------------
    
    ' --- 声明变量 ---
    Dim wsMain As Worksheet, wsSub As Worksheet
    Dim dict As Object
    Dim mainRow As Long, subRow As Long, lastMainRow As Long, lastSubRow As Long
    Dim key As Variant
    Dim dataArray() As Variant
    Dim i As Long
    Dim startTime As Double
    
    ' --- 配置参数 (请根据你的实际Excel文件修改以下内容) ---
    
    ' 1. 指定工作表
    Set wsMain = ThisWorkbook.Sheets("锆石录入20250409")  ' <-- 修改这里: 需要进行数据填充的目标表格
    Set wsSub = ThisWorkbook.Sheets("Sheet1")          ' <-- 修改这里: 提供数据的源头表格

    ' 2. 定义匹配的关键字在哪一列
    Const MAIN_MATCH_COLUMN As Long = 11  ' 主表中<关键字列>的列号 (A=1, B=2, ...)。当前为第11列(K列)。
    Const SUB_MATCH_COLUMN As Long = 1    ' 副表中<关键字列>的列号。当前为第1列(A列)。

    ' 3. 定义需要从副表复制哪些数据
    Const SUB_DATA_START_COLUMN As Long = 2  ' 从副表的第几列开始复制数据。当前为第2列(B列)。
    Const SUB_DATA_COLUMN_COUNT As Long = 8  ' 总共复制多少列数据。当前为8列(B到I列)。

    ' 4. 定义数据要写入主表的哪个位置
    Const TARGET_START_COLUMN As String = "AT" ' 匹配成功后，数据将从主表的这一列开始写入。

    ' 5. 定义数据处理的行范围 (手动指定)
    Const MAIN_START_ROW As Long = 7399 ' 主表的数据从第几行开始 (标题行不算)。
    lastMainRow = 7466                  ' 主表数据的结束行。
    
    Const SUB_START_ROW As Long = 2   ' 副表的数据从第几行开始。
    lastSubRow = 85                     ' 副表数据的结束行。
    
    ' (可选) 如果想让程序自动寻找最后一行，可以取消下面两行的注释，并注释掉上面的 lastMainRow 和 lastSubRow
    ' lastMainRow = wsMain.Cells(wsMain.Rows.Count, MAIN_MATCH_COLUMN).End(xlUp).Row
    ' lastSubRow = wsSub.Cells(wsSub.Rows.Count, SUB_MATCH_COLUMN).End(xlUp).Row
    
    ' --- 初始化 ---
    startTime = Timer
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' --- 步骤 1: 读取副表数据到字典中，建立高效的查找“手册” ---
    For subRow = SUB_START_ROW To lastSubRow
        key = wsSub.Cells(subRow, SUB_MATCH_COLUMN).Value ' 获取副表当前行的关键字
        
        ' 如果关键字不为空且字典中不存在，则处理
        If Not IsEmpty(key) And Not dict.exists(key) Then
            ' 将需要复制的多列数据打包存入一个数组中
            ReDim dataArray(1 To SUB_DATA_COLUMN_COUNT)
            For i = 1 To SUB_DATA_COLUMN_COUNT
                dataArray(i) = wsSub.Cells(subRow, SUB_DATA_START_COLUMN + i - 1).Value
            Next i
            
            ' 将关键字和对应的数据数组存入字典
            dict(key) = dataArray
        End If
    Next subRow
    
    ' --- 步骤 2: 遍历主表，进行匹配并填充数据 ---
    Dim targetStartColNum As Long
    targetStartColNum = wsMain.Range(TARGET_START_COLUMN & "1").Column
    
    For mainRow = MAIN_START_ROW To lastMainRow
        ' 从主表中获取当前行的关键字
        key = wsMain.Cells(mainRow, MAIN_MATCH_COLUMN).Value
        
        ' 检查字典中是否存在这个关键字
        If dict.exists(key) Then
            ' 如果找到，从字典中提取打包好的数据数组
            dataArray = dict(key)
            
            ' 将数组中的数据依次写入到主表的目标区域
            For i = 1 To SUB_DATA_COLUMN_COUNT
                wsMain.Cells(mainRow, targetStartColNum + i - 1).Value = dataArray(i)
            Next i
        Else
            ' (可选) 如果在副表中没有找到匹配项，则清空主表的目标单元格区域，以防有旧数据
            wsMain.Cells(mainRow, targetStartColNum).Resize(1, SUB_DATA_COLUMN_COUNT).ClearContents
        End If
    Next mainRow
    
    ' --- 清理与收尾 ---
    Set dict = Nothing
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    ' --- 提示完成 ---
    MsgBox "数据匹配与填充完成！" & vbNewLine & _
           "共处理主表 " & (lastMainRow - MAIN_START_ROW + 1) & " 行数据。" & vbNewLine & _
           "耗时 " & Round(Timer - startTime, 2) & " 秒。", vbInformation, "操作成功"
End Sub