Attribute VB_Name = "LegalRestDays"
'-----------------------------
' 函数名称: LegalRestDay
' 功能: 获取法定休息日数据
' 参数:
'   Period [Variant] 可选参数，支持以下输入类型：
'     - 日期/序列号：返回该日期之后的所有法定假期
'     - 年份(2010-2100)：返回指定年份的法定假期
'     - 负数：返回过去N个月之未来12个月内的法定假期（默认6个月）
'     - 0：返回原始数据
' 返回值:  [Variant] 日期数组或错误值
'-----------------------------
Function LegalRestDay(Optional Period As Variant = 0) As Variant
    On Error GoTo ErrorHandler  ' 统一错误处理
    
    '--------------- 局部变量声明 ----------------
    Dim ws As Worksheet                ' 数据工作表对象
    Dim dataRange As Range             ' 原始数据范围
    Dim rawData As Variant             ' 原始数据数组
    Dim result() As Double             ' 结果数组(日期序列值)
    Dim i As Long, j As Long           ' 循环计数器
    Dim targetYear As Integer          ' 目标年份缓存
    Dim startDate As Date, endDate As Date  ' 日期范围边界
    Dim cellValue As Variant           ' 临时单元格值
    Dim tempDate As Date               ' 临时日期转换
    Dim monthsOffset As Integer        ' 月份偏移量计算
    Dim targetDays As Long             ' 目标日期序列值
    
    '--------------- 数据准备阶段 ----------------
    ' 获取法定假期数据工作表
    Set ws = ThisWorkbook.Sheets("LegalDays")
    ' 动态获取数据范围(从A2开始到最后一个非空单元格)
    Set dataRange = ws.Range("A2:A" & ws.Cells(ws.rows.Count, "A").End(xlUp).row)
    rawData = dataRange.Value          ' 将数据加载到数组提升性能
    
    ' 初始化结果数组(初始大小与原始数据相同)
    ReDim result(1 To UBound(rawData))
    
    '--------------- 核心处理逻辑 ----------------
    Select Case True
        ' 模式1：日期/序列号查询（返回指定日期之后的所有假期）
        Case IsDate(Period) Or (IsNumeric(Period) And Period > 40000)
            ' 统一转换为日期序列值处理
            targetDays = CLng(IIf(IsDate(Period), CDate(Period), Period))
            
            For i = 1 To UBound(rawData)
                cellValue = rawData(i, 1)
                If IsDate(cellValue) Then
                    tempDate = CDate(cellValue)
                    ' 筛选大于目标日期的记录
                    If tempDate > CDate(targetDays) Then
                        j = j + 1
                        result(j) = CDbl(tempDate)  ' 存储为Double类型日期值
                    End If
                End If
            Next
            
        ' 模式2：年份查询（返回指定年份的假期）
        Case Period >= 2010 And Period <= 2100
            targetYear = CInt(Period)
            For i = 1 To UBound(rawData)
                cellValue = rawData(i, 1)
                If IsDate(cellValue) Then
                    tempDate = CDate(cellValue)
                    ' 精确匹配年份
                    If year(tempDate) = targetYear Then
                        j = j + 1
                        result(j) = CDbl(tempDate)
                    End If
                End If
            Next
            
        ' 模式3：负数查询（返回过去N个月之未来12个月的假期，默认6个月）
        Case Period < 0
            ' 限制最小查询范围为18个月
            monthsOffset = IIf(Period < -6, Period, -6)
            startDate = DateAdd("m", monthsOffset, Date)  ' 计算起始日期
            endDate = DateAdd("m", 18, Date)               ' 计算结束日期
            
            For i = 1 To UBound(rawData)
                cellValue = rawData(i, 1)
                If IsDate(cellValue) Then
                    tempDate = CDate(cellValue)
                    ' 筛选日期范围内的记录
                    If tempDate >= startDate And tempDate <= endDate Then
                        j = j + 1
                        result(j) = CDbl(tempDate)
                    End If
                End If
            Next
            
        ' 模式4：默认参数（返回原始数据）
        Case Period = 0
            LegalRestDay = dataRange.Value  ' 直接返回原始数据范围
            Exit Function
            
        ' 无效参数处理
        Case Else
            LegalRestDay = CVErr(xlErrValue)
            Exit Function
    End Select
    
    '--------------- 结果处理阶段 ----------------
    If j > 0 Then
        ' 精确调整数组大小并转置为垂直数组
        ReDim Preserve result(1 To j)
        LegalRestDay = Application.Transpose(result)
    Else
        ' 无匹配结果时返回NA错误
        LegalRestDay = CVErr(xlErrNA)
    End If
    Exit Function
    
ErrorHandler:
    ' 统一错误处理，返回标准错误值
    LegalRestDay = CVErr(xlErrNA)
End Function
