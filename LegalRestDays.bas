Attribute VB_Name = "LegalRestDays"
'-----------------------------
' ��������: LegalRestDay
' ����: ��ȡ������Ϣ������
' ����:
'   Period [Variant] ��ѡ������֧�������������ͣ�
'     - ����/���кţ����ظ�����֮������з�������
'     - ���(2010-2100)������ָ����ݵķ�������
'     - ���������ع�ȥN����֮δ��12�����ڵķ������ڣ�Ĭ��6���£�
'     - 0������ԭʼ����
' ����ֵ:  [Variant] ������������ֵ
'-----------------------------
Function LegalRestDay(Optional Period As Variant = 0) As Variant
    On Error GoTo ErrorHandler  ' ͳһ������
    
    '--------------- �ֲ��������� ----------------
    Dim ws As Worksheet                ' ���ݹ��������
    Dim dataRange As Range             ' ԭʼ���ݷ�Χ
    Dim rawData As Variant             ' ԭʼ��������
    Dim result() As Double             ' �������(��������ֵ)
    Dim i As Long, j As Long           ' ѭ��������
    Dim targetYear As Integer          ' Ŀ����ݻ���
    Dim startDate As Date, endDate As Date  ' ���ڷ�Χ�߽�
    Dim cellValue As Variant           ' ��ʱ��Ԫ��ֵ
    Dim tempDate As Date               ' ��ʱ����ת��
    Dim monthsOffset As Integer        ' �·�ƫ��������
    Dim targetDays As Long             ' Ŀ����������ֵ
    
    '--------------- ����׼���׶� ----------------
    ' ��ȡ�����������ݹ�����
    Set ws = ThisWorkbook.Sheets("LegalDays")
    ' ��̬��ȡ���ݷ�Χ(��A2��ʼ�����һ���ǿյ�Ԫ��)
    Set dataRange = ws.Range("A2:A" & ws.Cells(ws.rows.Count, "A").End(xlUp).row)
    rawData = dataRange.Value          ' �����ݼ��ص�������������
    
    ' ��ʼ���������(��ʼ��С��ԭʼ������ͬ)
    ReDim result(1 To UBound(rawData))
    
    '--------------- ���Ĵ����߼� ----------------
    Select Case True
        ' ģʽ1������/���кŲ�ѯ������ָ������֮������м��ڣ�
        Case IsDate(Period) Or (IsNumeric(Period) And Period > 40000)
            ' ͳһת��Ϊ��������ֵ����
            targetDays = CLng(IIf(IsDate(Period), CDate(Period), Period))
            
            For i = 1 To UBound(rawData)
                cellValue = rawData(i, 1)
                If IsDate(cellValue) Then
                    tempDate = CDate(cellValue)
                    ' ɸѡ����Ŀ�����ڵļ�¼
                    If tempDate > CDate(targetDays) Then
                        j = j + 1
                        result(j) = CDbl(tempDate)  ' �洢ΪDouble��������ֵ
                    End If
                End If
            Next
            
        ' ģʽ2����ݲ�ѯ������ָ����ݵļ��ڣ�
        Case Period >= 2010 And Period <= 2100
            targetYear = CInt(Period)
            For i = 1 To UBound(rawData)
                cellValue = rawData(i, 1)
                If IsDate(cellValue) Then
                    tempDate = CDate(cellValue)
                    ' ��ȷƥ�����
                    If year(tempDate) = targetYear Then
                        j = j + 1
                        result(j) = CDbl(tempDate)
                    End If
                End If
            Next
            
        ' ģʽ3��������ѯ�����ع�ȥN����֮δ��12���µļ��ڣ�Ĭ��6���£�
        Case Period < 0
            ' ������С��ѯ��ΧΪ18����
            monthsOffset = IIf(Period < -6, Period, -6)
            startDate = DateAdd("m", monthsOffset, Date)  ' ������ʼ����
            endDate = DateAdd("m", 18, Date)               ' �����������
            
            For i = 1 To UBound(rawData)
                cellValue = rawData(i, 1)
                If IsDate(cellValue) Then
                    tempDate = CDate(cellValue)
                    ' ɸѡ���ڷ�Χ�ڵļ�¼
                    If tempDate >= startDate And tempDate <= endDate Then
                        j = j + 1
                        result(j) = CDbl(tempDate)
                    End If
                End If
            Next
            
        ' ģʽ4��Ĭ�ϲ���������ԭʼ���ݣ�
        Case Period = 0
            LegalRestDay = dataRange.Value  ' ֱ�ӷ���ԭʼ���ݷ�Χ
            Exit Function
            
        ' ��Ч��������
        Case Else
            LegalRestDay = CVErr(xlErrValue)
            Exit Function
    End Select
    
    '--------------- �������׶� ----------------
    If j > 0 Then
        ' ��ȷ���������С��ת��Ϊ��ֱ����
        ReDim Preserve result(1 To j)
        LegalRestDay = Application.Transpose(result)
    Else
        ' ��ƥ����ʱ����NA����
        LegalRestDay = CVErr(xlErrNA)
    End If
    Exit Function
    
ErrorHandler:
    ' ͳһ���������ر�׼����ֵ
    LegalRestDay = CVErr(xlErrNA)
End Function
