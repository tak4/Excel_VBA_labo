Attribute VB_Name = "Module1"
Option Explicit

' Setting
Const SCHEDULE_SETTING_SHEET As String = "schedule_macro"
Const SETTING_WB_NAME_CELL As String = "B1"
Const SETTING_WS_NAME_CELL As String = "B2"


Const SCHEDULE_START_ROW As Integer = 10
Const SCHEDULE_END_ROW As Integer = 20
Const SCHEDULE_START_COL As Integer = 6
Const SCHEDULE_END_COL As Integer = 16
Const WORK_DAYS_ROW = 6     ' �c�Ɠ����L�ڂ����s��
Const DATE_ROW = 8  ' ���t���L�ڂ����s��

Const WORKER_COL As Integer = 2 ' ��ƎҖ��̗�
Const REQUIRED_FOR_INPUT_COL As Integer = 3 ' ���͕K�v�H���̗�
Const START_WORK_DATE_COL As Integer = 4
Const END_WORK_DATE_COL As Integer = 5

Const INPUT_WORK_DAY As Integer = 5 ' ���͊�H��

Sub InputDate()
Attribute InputDate.VB_ProcData.VB_Invoke_Func = "S\n14"

    Dim r, c As Integer
    Dim wb As Workbook
    Dim ws As Worksheet
    
    Dim setting_wb_name As String
    Dim setting_ws_name As String
    Dim setting_schedule_start_row As Integer
    Dim setting_schedule_end_row As Integer
    Dim setting_schedule_start_col As Integer
    Dim setting_schedule_end_col As Integer
    
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
    
    
    ' ���͊�H��
    work_days_for_week = INPUT_WORK_DAY
    
    setting_wb_name = ThisWorkbook.Worksheets(SCHEDULE_SETTING_SHEET).Range(SETTING_WB_NAME_CELL).Value
    setting_ws_name = ThisWorkbook.Worksheets(SCHEDULE_SETTING_SHEET).Range(SETTING_WS_NAME_CELL).Value
    setting_schedule_start_row
    setting_schedule_end_row
    setting_schedule_start_col
    setting_schedule_end_col
    
    Set wb = Workbooks(setting_wb_name)
    Set ws = wb.Worksheets(setting_ws_name)
    act_row = ActiveCell.Row
    act_col = ActiveCell.Column
    worker_name = ws.Cells(act_row, WORKER_COL).Value

    ' ���͕K�v�H��
    total_for_work = 0
    For c = SCHEDULE_START_COL To SCHEDULE_END_COL
        total_for_work = total_for_work + ws.Cells(act_row, c).Value
    Next c
    
    required_for_input = ws.Cells(act_row, REQUIRED_FOR_INPUT_COL) - total_for_work
    
    ' ���͕K�v�H�����J��Ԃ�
    c = 0
    Do
        ' ���͍ς݂̃Z���̓X�L�b�v
        If ws.Cells(act_row, act_col + c).Value <> "" Then
            GoTo CONTINUE
        End If
    
        ' ���͍ςݍH��(��)
        total_for_days = 0
        For r = SCHEDULE_START_ROW To SCHEDULE_END_ROW
            If ws.Cells(r, WORKER_COL).Value = worker_name Then
                total_for_days = total_for_days + ws.Cells(r, act_col + c).Value
            End If
        Next r
        
        ' ���͉\�H��(��) �c�Ɠ�������͍ςݍH��(��)������
        can_input = ws.Cells(WORK_DAYS_ROW, act_col + c).Value - total_for_days
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
    
    Loop While required_for_input > 0 And c <= SCHEDULE_END_ROW
    
    ' �J�n��/�I��������
    ' �J�n�I�����̏T�̗񐔂����߂�
    c = 0
    start_date_col = -1
    end_date_col = -1
    For c = SCHEDULE_START_COL To SCHEDULE_END_COL
        If ws.Cells(act_row, c).Value <> 0 Then
            If start_date_col = -1 Then
                start_date_col = c
            End If
            end_date_col = c
        End If
    Next c
    
    ' ��ƊJ�n�T�̑���Ƃ̓��͍ςݍH�����擾
    total_for_days = 0
    For r = SCHEDULE_START_ROW To SCHEDULE_END_ROW
        If r <> act_row Then
            total_for_days = total_for_days + ws.Cells(r, start_date_col).Value
        End If
    Next r
    
    ws.Cells(act_row, START_WORK_DATE_COL).Value = ws.Cells(DATE_ROW, start_date_col).Value + total_for_days
    ws.Cells(act_row, END_WORK_DATE_COL).Value = ws.Cells(DATE_ROW, end_date_col).Value + ws.Cells(act_row, end_date_col).Value
    

End Sub
