Attribute VB_Name = "HolidaySync"
'-----------------------------
' 模块名称: HolidaySync (合并优化版)
' 功能: 法定假期数据同步系统
'-----------------------------
Option Explicit
Option Private Module  ' 防止公共函数暴露给其他工作簿

'--------------- 全局配置 ----------------
Private Const GITHUB_REPO_URL As String = "https://raw.githubusercontent.com/laomor/holiday-data/main/All%20Years/"
Public Const CONFIG_SHEET As String = "Config"    ' 改为公共常量供其他模块访问
Public Const DATA_SHEET As String = "LegalDays"
Private Const CHECK_INTERVAL_CRITICAL As Integer = 1    ' 关键月份检查间隔

'--------------- 枚举常量 ----------------
Private Enum xlSortSettings
    xlSortOnValues = 0
    xlAscending = 1
    xlYes = 1
    xlTopToBottom = 1
End Enum

'--------------- 对象变量 ----------------
Private logBuffer As Collection  ' 日志缓冲区

'--------------- 事件处理器 ----------------
' 工作簿打开时触发（需在ThisWorkbook中绑定）
Public Sub Workbook_Open()
    If NeedCheckUpdate() Then
        Application.OnTime Now + TimeValue("00:01:00"), "CheckHolidayUpdates"
    End If
End Sub

'--------------- 核心功能 ----------------
' 检查更新必要性
Public Function NeedCheckUpdate() As Boolean
    Dim lastCheckDate As Date, currentYr As Integer, currentMonth As Integer
    Dim expectedYears As Collection, localYears As Collection
    Dim yr As Integer, elem As Variant
    
    ' 获取当前年月
    currentYr = VBA.year(Date)
    currentMonth = VBA.Month(Date)
    
    ' 生成动态年份范围（保留12月自动包含下一年逻辑）
    Set expectedYears = New Collection
    For yr = 2011 To currentYr + IIf(currentMonth = 12, 1, 0)
        expectedYears.Add CStr(yr)
    Next
    
    ' 获取本地已有数据
    Set localYears = GetLocalYears()
    
    ' 唯一判断条件：数据完整性检查
    For Each elem In expectedYears
        If Not CollectionContains(localYears, CStr(elem)) Then
            NeedCheckUpdate = True
            Exit Function
        End If
    Next
    
    ' 数据完整时直接返回不需要更新
    NeedCheckUpdate = False
    
End Function

' 主更新流程
Public Sub CheckHolidayUpdates()
    InitLogBuffer  ' 初始化日志
    
    Dim remoteYears As Collection, expectedYears As Collection
    Dim yr As Integer, targetYear As Variant, csvData As String, updateCount As Integer
    Dim currentYr As Integer, currentMonth As Integer
    
    currentYr = VBA.year(Date)
    currentMonth = VBA.Month(Date)
    
    ' 生成预期年份
    Set expectedYears = New Collection
    For yr = 2011 To currentYr + IIf(currentMonth = 12, 1, 0)
        expectedYears.Add CStr(yr)
    Next
    
    ' 获取远程数据
    Set remoteYears = GetRemoteYears()
    If remoteYears Is Nothing Then GoTo Finalize
    
    ' 遍历更新
    updateCount = 0
    For Each targetYear In expectedYears
        Dim yearStr As String: yearStr = CStr(targetYear)
        If CollectionContains(remoteYears, yearStr) And Not CollectionContains(GetLocalYears(), yearStr) Then
            csvData = DownloadCSVData(yearStr)
            If csvData <> "" Then
                MergeDataToSheet csvData, yearStr
                UpdateConfig yearStr
                updateCount = updateCount + 1
            End If
        End If
    Next

Finalize:
    FlushLogs  ' 必须执行的日志清理
    
     ' 数据保存
   If updateCount > 0 Then
        On Error Resume Next
        ThisWorkbook.Save
        Select Case Err.Number
            Case 0
                ThisWorkbook.Sheets(CONFIG_SHEET).Range("C2").Value = "最后保存: " & Now
                ThisWorkbook.Save
            Case Else
                ThisWorkbook.Sheets(CONFIG_SHEET).Range("D1").Value = "保存失败: " & Err.Description & " (" & Now & ")"
                ThisWorkbook.Save
        End Select
        On Error GoTo 0
    End If
    
    ThisWorkbook.Save
    
    'If updateCount > 0 Then MsgBox "已更新 " & updateCount & " 年数据", vbInformation
End Sub

'--------------- 网络操作 ----------------
' 获取远程年份
Private Function GetRemoteYears() As Collection
    Dim http As Object, response As String, regex As Object
    Dim years As New Collection, Match As Object
    
    Set http = CreateObject("MSXML2.ServerXMLHTTP")
    http.Open "GET", "https://api.github.com/repos/laomor/holiday-data/contents/All%20Years/", False
    http.setTimeouts 5000, 5000, 10000, 10000
    http.send
    
    If http.Status = 200 Then
        response = http.responseText
        Set regex = CreateObject("VBScript.RegExp")
        regex.Pattern = """name"":\s*""(\d{4})\.csv"""
        regex.Global = True
        
        For Each Match In regex.Execute(response)
            Dim year As String: year = CStr(Match.SubMatches(0))
            If Len(year) = 4 And IsNumeric(year) Then years.Add year
        Next
    Else
        LogError "GitHub请求失败: " & http.Status & " " & http.statusText
    End If
    
    Set GetRemoteYears = years
End Function

' 下载CSV
Private Function DownloadCSVData(year As String) As String
    Dim http As Object, url As String, statusText As String
    Dim retry As Integer, startTime As Double
    
    url = GITHUB_REPO_URL & year & ".csv"
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    
    For retry = 1 To 1  ' 为减少Excel不可操作时间故调整为1次，需要时可增加次数
        On Error Resume Next
        startTime = Timer
        http.Open "GET", url, False
        http.setTimeouts 5000, 5000, 10000, 10000
        http.setRequestHeader "User-Agent", "Mozilla/5.0"
        SetProxySettings http
        http.send
        
        If Err.Number = 0 Then
            If http.Status = 200 Then
                DownloadCSVData = http.responseText: Exit Function
            Else
                statusText = "HTTP状态码：" & http.Status & " " & http.statusText
            End If
        Else
            statusText = "系统错误: " & Err.Description & " (0x" & Hex(Err.Number) & ")"
        End If
        
        If retry < 1 Then  ' 为减少Excel不可操作时间故调整为1次，需要时可增加次数
            LogError year & ".csv 下载失败（第" & retry & "次） - 耗时" & Format(Timer - startTime, "0.00") & "秒"
            Application.Wait Now + TimeSerial(0, 0, 2 ^ retry)
        End If
        On Error GoTo 0
    Next
    
    LogError year & ".csv 最终失败 - " & statusText
    DownloadCSVData = ""
End Function

' 代理设置（原模块2）
Private Sub SetProxySettings(httpObj As Object)
    On Error Resume Next
    'httpObj.SetProxy 2, ""  ' 自动代理检测
    On Error GoTo 0
End Sub

'--------------- 数据操作 ----------------
' 合并数据
Private Sub MergeDataToSheet(csvData As String, year As String)
    If Trim(csvData) = "" Then Exit Sub
    
    Dim rows As Variant, dataArray() As Variant, i As Long, newRows As Long
    Dim existingDates As Object, currentDate As Date
    
    On Error GoTo MergeError
    
    ' 确保目标表存在
    If Not WorksheetExists(DATA_SHEET) Then
        With ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
            .Name = DATA_SHEET
            .Cells(1, 1).Value = "法定假期"
        End With
    End If
    
    ' 加载现有日期
    Set existingDates = CreateObject("Scripting.Dictionary")
    With ThisWorkbook.Sheets(DATA_SHEET)
        Dim lastRow As Long: lastRow = .Cells(.rows.Count, 1).End(xlUp).row
        If lastRow > 1 Then
            Dim cell As Range
            For Each cell In .Range("A2:A" & lastRow)
                existingDates(CDate(cell.Value)) = True
            Next
        End If
    End With
    
    ' 处理新数据
    rows = Split(csvData, vbCrLf)
    ReDim dataArray(1 To UBound(rows) + 1, 1 To 1)
    
    For i = 0 To UBound(rows)
        Dim rowData As String: rowData = Trim(rows(i))
        If rowData <> "" And IsDate(rowData) Then
            currentDate = CDate(rowData)
            If Not existingDates.Exists(currentDate) Then
                newRows = newRows + 1
                dataArray(newRows, 1) = currentDate
                existingDates(currentDate) = True
            End If
        ElseIf rowData <> "" Then
            LogError year & " 数据异常: 行 " & i + 1 & " '" & rowData & "'不是有效日期"
        End If
    Next
    
    ' 写入数据
    If newRows > 0 Then
        With ThisWorkbook.Sheets(DATA_SHEET)
            lastRow = .Cells(.rows.Count, 1).End(xlUp).row
            If lastRow = 1 And .Cells(1, 1).Value = "" Then lastRow = 0
            .Cells(lastRow + 1, 1).Resize(newRows, 1).Value = dataArray
        End With
    End If
    
    SortDates
    Exit Sub

MergeError:
    LogError year & " 合并失败: 行 " & i + 1 & " | " & Err.Description
End Sub

' 日期排序
Private Sub SortDates()
    On Error GoTo SortError
    
    With ThisWorkbook.Sheets(DATA_SHEET)
        Dim lastRow As Long
        lastRow = .Cells(.rows.Count, 1).End(xlUp).row
        If lastRow < 2 Then Exit Sub
        
        ' 显式声明排序对象
        Dim sortRange As Range
        Set sortRange = .Range("A1:A" & lastRow)
        
        ' 清除可能存在的临时排序
        .Sort.SortFields.Clear
        
        ' 使用兼容性更好的Add方法替代Add2
        .Sort.SortFields.Add Key:=sortRange.Offset(1, 0).Resize(sortRange.rows.Count - 1), _
            SortOn:=xlSortOnValues, Order:=xlAscending
        
        With .Sort
            .SetRange sortRange
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .Apply
        End With
    End With
    Exit Sub

SortError:
    LogError "排序失败: " & Err.Description & " [lastRow=" & lastRow & "]"
End Sub

'--------------- 配置管理---------------
' 初始化配置表
Public Sub InitializeConfigSheet()
    On Error GoTo ErrorHandler
    If Not WorksheetExists(CONFIG_SHEET) Then
        With ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
            .Name = CONFIG_SHEET
            .Range("A1:B1").Value = Array("Year", "LastUpdated")
            .Range("C1").Value = "LastCheckDate"
            .Range("D1").Value = "ErrorLog"
        End With
    End If
    Exit Sub
ErrorHandler:
    LogError "初始化Config失败: " & Err.Description
End Sub

' 获取本地年份
Public Function GetLocalYears() As Collection
    Dim years As New Collection, ws As Worksheet
    
    InitializeConfigSheet
    Set ws = ThisWorkbook.Sheets(CONFIG_SHEET)
    
    With ws
        Dim lastRow As Long: lastRow = .Cells(.rows.Count, "A").End(xlUp).row
        If lastRow >= 2 Then
            Dim yearCell As Range
            For Each yearCell In .Range("A2:A" & lastRow)
                If Not IsEmpty(yearCell.Value) And IsNumeric(yearCell.Value) Then
                    years.Add CStr(yearCell.Value)
                End If
            Next
        End If
    End With
    
    Set GetLocalYears = years
End Function

' 更新配置
Public Sub UpdateConfig(year As String)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(CONFIG_SHEET)
    Dim lastRow As Long: lastRow = ws.Cells(ws.rows.Count, "A").End(xlUp).row
    
    If lastRow = 1 And ws.Cells(1, 1).Value = "" Then
        ws.Cells(1, 1).Value = year
        ws.Cells(1, 2).Value = Now
    Else
        ws.Cells(lastRow + 1, 1).Value = year
        ws.Cells(lastRow + 1, 2).Value = Now
    End If
End Sub

' 集合包含检查
Private Function CollectionContains(col As Collection, item As String) As Boolean
    Dim elem As Variant
    For Each elem In col
        If CStr(elem) = item Then
            CollectionContains = True: Exit Function
        End If
    Next
End Function

' 工作表存在检查
Public Function WorksheetExists(sheetName As String) As Boolean
    On Error Resume Next
    WorksheetExists = Not ThisWorkbook.Sheets(sheetName) Is Nothing
    On Error GoTo 0
End Function

'--------------- 日志系统---------------
' 初始化日志
Public Sub InitLogBuffer()
    Set logBuffer = New Collection
End Sub

' 记录错误
Public Sub LogError(msg As String)
    If logBuffer Is Nothing Then InitLogBuffer
    logBuffer.Add Format(Now, "yyyy-mm-dd hh:mm") & " | " & Left(msg, 255)
End Sub

' 写入日志
Public Sub FlushLogs()
    On Error Resume Next
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(CONFIG_SHEET)
    If ws Is Nothing Then Exit Sub
    
    With ws
        Dim i As Long
        For i = 1 To logBuffer.Count
            .Cells(.rows.Count, "D").End(xlUp).Offset(1).Value = logBuffer(i)
        Next
        
        ' 日志轮转
        Const MAX_LOGS As Long = 30
        Dim lastLogRow As Long: lastLogRow = .Cells(.rows.Count, "D").End(xlUp).row
        If lastLogRow > MAX_LOGS + 1 Then
            .Range("D2:D" & (lastLogRow - MAX_LOGS)).Delete Shift:=xlUp
        End If
    End With
    Set logBuffer = Nothing
End Sub
