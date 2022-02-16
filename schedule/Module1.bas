Attribute VB_Name = "Module1"
Option Explicit

' Setting
Const SCHEDULE_SETTING_SHEET As String = "schedule_macro"
Const SETTING_WB_NAME_CELL As String = "B1"
Const SETTING_WS_NAME_CELL As String = "B2"
Const SETTING_SCHEDULE_START_ROW_CELL As String = "B3"
Const SETTING_SCHEDULE_END_ROW_CELL As String = "B4"
Const SETTING_SCHEDULE_START_COL_CELL As String = "B5"
Const SETTING_SCHEDULE_END_COL_CELL As String = "B6"
Const SETTING_SCHEDULE_WORK_DAYS_ROW_CELL As String = "B7"
Const SETTING_SCHEDULE_DATE_ROW_CELL As String = "B8"
Const SETTING_SCHEDULE_WORKER_COL_CELL As String = "B9"
Const SETTING_SCHEDULE_REQUIRED_FOR_INPUT_COL_CELL As String = "B10"
Const SETTING_SCHEDULE_START_WORK_DATE_COL_CELL As String = "B11"
Const SETTING_SCHEDULE_END_WORK_DATE_COL_CELL As String = "B12"
Const SETTING_SCHEDULE_INPUT_WORK_DAY_CELL As String = "B13"

Const HOLIDAY_LIST_SIZE As Integer = 30

Type HolidayList
    day As Date
    holiday As Boolean
End Type


Sub InputScheduleUpdateDate()
Attribute InputScheduleUpdateDate.VB_ProcData.VB_Invoke_Func = "D\n14"

    InputDate (True)

End Sub

Sub InputSchedule()
Attribute InputSchedule.VB_ProcData.VB_Invoke_Func = "A\n14"

    InputDate (False)

End Sub

Sub InputDate(update_date As Boolean)
Attribute InputDate.VB_ProcData.VB_Invoke_Func = "S\n14"

    Dim r, c, n, i As Integer
    Dim wb As Workbook
    Dim macro_ws, ws As Worksheet
    
    ' �ݒ�p�����[�^
    Dim setting_wb_name As String
    Dim setting_ws_name As String
    Dim setting_schedule_start_row As Integer
    Dim setting_schedule_end_row As Integer
    Dim setting_schedule_start_col As Integer
    Dim setting_schedule_end_col As Integer
    Dim setting_work_days_row As Integer
    Dim setting_date_row As Integer
    Dim setting_worker_col As Integer
    Dim setting_required_for_input_col As Integer
    Dim setting_start_work_date_col As Integer
    Dim setting_end_work_date_col As Integer
    Dim setting_input_work_date As Integer
    
    ' ��Ɨp�ϐ�
    Dim worker_name As String       ' ��ƎҖ�
    Dim work_days_for_week As Double   ' 1�T�Ԃɂ�����H��
    Dim act_row, act_col As Double
    
    Dim total_for_days As Double       ' ���͍ςݍH��(��)
    Dim total_for_work As Double       ' ���͍ςݍH��(���)
    Dim required_for_input As Double ' ���͕K�v�H��
    Dim can_input As Double           ' ���͉\�H��
    Dim input_work_days As Double  ' ���͍H��
    
    Dim start_date_col As Integer   ' ��ƊJ�n�T�̗�
    Dim end_date_col As Integer ' ��ƏI���T�̗�
    Dim start_date As Date
    Dim end_date As Date
    
    Dim hList() As HolidayList

    ' �ݒ�p�����[�^ ������
    Set macro_ws = ThisWorkbook.Worksheets(SCHEDULE_SETTING_SHEET)
    setting_wb_name = macro_ws.Range(SETTING_WB_NAME_CELL).Value
    setting_ws_name = macro_ws.Range(SETTING_WS_NAME_CELL).Value
    setting_schedule_start_row = macro_ws.Range(SETTING_SCHEDULE_START_ROW_CELL).Value
    setting_schedule_end_row = macro_ws.Range(SETTING_SCHEDULE_END_ROW_CELL).Value
    setting_schedule_start_col = macro_ws.Range(SETTING_SCHEDULE_START_COL_CELL).Value
    setting_schedule_end_col = macro_ws.Range(SETTING_SCHEDULE_END_COL_CELL).Value
    setting_work_days_row = macro_ws.Range(SETTING_SCHEDULE_WORK_DAYS_ROW_CELL).Value
    setting_date_row = macro_ws.Range(SETTING_SCHEDULE_DATE_ROW_CELL).Value
    setting_worker_col = macro_ws.Range(SETTING_SCHEDULE_WORKER_COL_CELL).Value
    setting_required_for_input_col = macro_ws.Range(SETTING_SCHEDULE_REQUIRED_FOR_INPUT_COL_CELL).Value
    setting_start_work_date_col = macro_ws.Range(SETTING_SCHEDULE_START_WORK_DATE_COL_CELL).Value
    setting_end_work_date_col = macro_ws.Range(SETTING_SCHEDULE_END_WORK_DATE_COL_CELL).Value
    setting_input_work_date = macro_ws.Range(SETTING_SCHEDULE_INPUT_WORK_DAY_CELL).Value

    ' ��Ɨp�ϐ�������
    Set wb = Workbooks(setting_wb_name)
    Set ws = wb.Worksheets(setting_ws_name)
    act_row = ActiveCell.Row
    act_col = ActiveCell.Column
    worker_name = ws.Cells(act_row, setting_worker_col).Value
    
    ' �x���f�[�^�擾
'    r = 1
'    i = 0
'    n = 0
'    ReDim Preserve hList(HOLIDAY_LIST_SIZE) As HolidayList
'    While macro_ws.Cells(r, 5).Value <> ""
'        hList(i).day = macro_ws.Cells(r, 5).Value
'        If macro_ws.Cells(r, 6).Value <> "" Then
'            hList(i).holiday = True
'        Else
'            hList(i).holiday = False
'        End If
'
'        n = UBound(hList)
'        If n <= i Then
'            ReDim Preserve hList(n + HOLIDAY_LIST_SIZE) As HolidayList
'        End If
'        r = r + 1
'        i = i + 1
'    Wend
'
'    For n = 0 To UBound(hList)
'        If hList(n).day = 0 Then
'            Exit For
'        End If
'        Debug.Print (hList(n).day & " " & hList(n).holiday)
'    Next n

    
    ' ���͊�H��
    work_days_for_week = setting_input_work_date

    ' ���͕K�v�H���擾
    total_for_work = 0
    For c = setting_schedule_start_col To setting_schedule_end_col
        total_for_work = total_for_work + ws.Cells(act_row, c).Value
    Next c
    required_for_input = ws.Cells(act_row, setting_required_for_input_col) - total_for_work

    
    ' �H������ - ���͕K�v�H�����J��Ԃ�
    c = 0
    Do
        ' ���͍ς݂̃Z���̓X�L�b�v
        If ws.Cells(act_row, act_col + c).Value <> "" Then
            GoTo CONTINUE
        End If
    
        ' ���͍ςݍH��(��)
        total_for_days = 0
        For r = setting_schedule_start_row To setting_schedule_end_row
            If ws.Cells(r, setting_worker_col).Value = worker_name Then
                total_for_days = total_for_days + ws.Cells(r, act_col + c).Value
            End If
        Next r
        
        ' ���͉\�H��(��) �c�Ɠ�������͍ςݍH��(��)������
        can_input = ws.Cells(setting_work_days_row, act_col + c).Value - total_for_days
        If can_input < 0 Then
            can_input = 0
        End If
        
        ' ���͍H������͉\�H��(��)�ŕ␳
        If can_input > work_days_for_week Then
            input_work_days = work_days_for_week
        Else
            input_work_days = can_input
        End If
        
        ' ���͍H������͕K�v�H���ŕ␳
        If required_for_input < input_work_days Then
            input_work_days = required_for_input
        End If
                
        ' ����
        If input_work_days > 0 Then
            ws.Cells(act_row, act_col + c).Value = input_work_days
            required_for_input = required_for_input - input_work_days
        End If
        
CONTINUE:
        ' ���̗�(�T)��
        c = c + 1
    
    Loop While required_for_input > 0 And c <= setting_schedule_end_row
    
    
    ' �J�n��/�I��������
    
    If update_date = True Then
        
        ' �J�n�I�����̏T�̗񐔂����߂�
        c = 0
        start_date_col = -1
        end_date_col = -1
        For c = setting_schedule_start_col To setting_schedule_end_col
            If ws.Cells(act_row, c).Value <> 0 Then
                If start_date_col = -1 Then
                    start_date_col = c
                End If
                end_date_col = c
            End If
        Next c
        
        ' �J�n��
        
        ' ��ƊJ�n�T�̑���Ƃ̓��͍ςݍH�����擾
        total_for_days = 0
        For r = setting_schedule_start_row To setting_schedule_end_row
            If ws.Cells(r, setting_worker_col).Value = worker_name Then
                If r <> act_row Then
                    total_for_days = total_for_days + ws.Cells(r, start_date_col).Value
                End If
            End If
        Next r
                
        start_date = CDate(ws.Cells(setting_date_row, start_date_col).Value) + total_for_days
        ws.Cells(act_row, setting_start_work_date_col).Value = start_date
        
        ' �I����
        
        ' ��ƏI���T�̑���Ƃ̓��͍ςݍH�����擾
        total_for_days = 0
        For r = setting_schedule_start_row To setting_schedule_end_row
            If ws.Cells(r, setting_worker_col).Value = worker_name Then
                ' If r <> act_row Then
                    total_for_days = total_for_days + ws.Cells(r, end_date_col).Value
                ' End If
            End If
        Next r
        
        end_date = CDate(ws.Cells(setting_date_row, end_date_col).Value) + total_for_days
        ws.Cells(act_row, setting_end_work_date_col).Value = end_date
        
        ' ws.Cells(act_row, setting_end_work_date_col).Value = ws.Cells(setting_date_row, end_date_col).Value + ws.Cells(act_row, end_date_col).Value - 1
        
    End If
    
End Sub


