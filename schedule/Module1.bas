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
Dim wb As Workbook
Dim macro_ws, ws As Worksheet

Dim worker_name As String       ' ��ƎҖ�
Dim work_days_for_week As Double   ' 1�T�Ԃɂ�����H��
Dim act_row, act_col As Integer

' �x�����X�g
Const HOLIDAY_LIST_SIZE As Integer = 30

Type HolidayList
    day As Date
    holiday As Boolean
End Type

'
' �}�N���G���g���F���ԁA�J�n���A�I��������͂���
'
Sub InputScheduleUpdateDate()
Attribute InputScheduleUpdateDate.VB_ProcData.VB_Invoke_Func = "D\n14"

    Initial
    InputPeriod
    InputWorkStartEndDate (act_row)

End Sub

'
' �}�N���G���g���F�J�n���A�I��������͂���
'
Sub InputUpdateDate()
Attribute InputUpdateDate.VB_ProcData.VB_Invoke_Func = "S\n14"

    Initial
    InputWorkStartEndDate (act_row)

End Sub


'
' �}�N���G���g���F�J�n���A�I�����␳(1�������̓��t���͂�؂�̂Ă�)
'
Sub CorrectionDate()
Attribute CorrectionDate.VB_ProcData.VB_Invoke_Func = "C\n14"
    Dim r As Integer
    Dim correct_start_date, entered_start_date As Date  ' ��ƊJ�n���t
    Dim correct_end_date, entered_end_date As Date      ' ��ƏI�����t
    
    Initial
    For r = setting_schedule_start_row To setting_schedule_end_row
        
        ' �J�n���␳
        entered_start_date = CDate(ws.Cells(r, setting_start_work_date_col).Value)
        correct_start_date = Int(CDate(ws.Cells(r, setting_start_work_date_col).Value))
        ws.Cells(r, setting_start_work_date_col).Value = correct_start_date

        ' �O�ׂ̈̓��t�̕ύX�`�F�b�N 1��������؂�̂Ă邾���Ȃ̂ŁA�ς��Ȃ��͂�
        If ws.Cells(r, setting_start_work_date_col).Value <> Int(entered_start_date) Then
            ws.Cells(r, setting_start_work_date_col).Font.ColorIndex = 5
        End If
        
        ' �I�����␳
        entered_end_date = CDate(ws.Cells(r, setting_end_work_date_col).Value)
        correct_end_date = Int(CDate(ws.Cells(r, setting_end_work_date_col).Value))
        ws.Cells(r, setting_end_work_date_col).Value = correct_end_date
        
        ' �O�ׂ̈̓��t�̕ύX�`�F�b�N 1��������؂�̂Ă邾���Ȃ̂ŁA�ς��Ȃ��͂�
        If ws.Cells(r, setting_end_work_date_col).Value <> Int(entered_end_date) Then
            ws.Cells(r, setting_end_work_date_col).Font.ColorIndex = 5
        End If
    
    Next r

End Sub


'
' ������
'
Sub Initial()

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

End Sub

'
' ���Ԃ���͂���
'
Sub InputPeriod()
Attribute InputPeriod.VB_ProcData.VB_Invoke_Func = " \n14"

    Dim r, c, n, i As Integer
    
    Dim total_for_days As Double       ' ���͍ςݍH��(��)
    Dim total_for_work As Double       ' ���͍ςݍH��(���)
    Dim required_for_input As Double ' ���͕K�v�H��
    Dim can_input As Double           ' ���͉\�H��
    Dim input_work_days As Double  ' ���͍H��
        
    ' ���͊�H��
    work_days_for_week = setting_input_work_date

    ' �c��̓��͕K�v�H���擾 ���͕K�v�H��������͍ςݍH��������
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
        
        ' ���͉\�H��(��)�F�c�Ɠ�������͍ςݍH��(��)������
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
                
        ' �H������
        If input_work_days > 0 Then
            ' ����
            ws.Cells(act_row, act_col + c).Value = input_work_days
            
            ' ���͕K�v�H�������Z
            required_for_input = required_for_input - input_work_days
        End If
        
CONTINUE:
        ' ���̗�(�T)��
        c = c + 1
    
    Loop While required_for_input > 0 And c <= setting_schedule_end_row

End Sub


'
' ��ƊJ�n��/�I��������
'
Sub InputWorkStartEndDate(ByVal target_row As Integer)
    
    Dim r, c As Integer ' Loop�p
    
    Dim start_date_col As Integer   ' ��ƊJ�n�T�̗�
    Dim end_date_col As Integer     ' ��ƏI���T�̗�
    Dim total_for_days As Double    ' ���͍ςݍH��(��)
    
    Dim start_date, entered_start_date As Date  ' ��ƊJ�n���t
    Dim end_date, entered_end_date As Date      ' ��ƏI�����t

    ' ��ƊJ�n�A��ƏI�����̏T�̗񐔂����߂�
    c = 0
    start_date_col = -1
    end_date_col = -1
    For c = setting_schedule_start_col To setting_schedule_end_col
        If ws.Cells(target_row, c).Value <> 0 Then
            If start_date_col = -1 Then
                start_date_col = c
            End If
            end_date_col = c
        End If
    Next c
        
    If start_date_col = -1 Then
        '�J�n����������Ȃ��ꍇ�͊J�n���A�I�������͂��X�L�b�v
        GoTo SKIP_INPUT_DATE
    End If

    
    ' �J�n�� ����Ƃ̓��͍ςݍH�����l�����ĊJ�n�������߂�
    
    ' ��ƊJ�n�T�̓��͍ςݍH��(�Ώۍ�Ƃ����O)���擾
    total_for_days = 0
    For r = setting_schedule_start_row To setting_schedule_end_row
        If ws.Cells(r, setting_worker_col).Value = worker_name Then
            If r <> target_row Then
                total_for_days = total_for_days + ws.Cells(r, start_date_col).Value
            End If
        End If
    Next r

    ' 1�������̎��Ԃ͐؂�̂Ă����̂�Int�Ŋۂ߂�
    ' �ۂ߂����ʂ����ɓ��͍ς݂ƈقȂ�ꍇ�͍X�V����
    entered_start_date = Int(CDate(ws.Cells(target_row, setting_start_work_date_col).Value))
    start_date = Int(CDate(ws.Cells(setting_date_row, start_date_col).Value) + total_for_days)
    If entered_start_date <> start_date Then
        ws.Cells(target_row, setting_start_work_date_col).Value = start_date
        ws.Cells(target_row, setting_start_work_date_col).Font.ColorIndex = 3
    End If
    
    ' �I���� ���͍ςݍH�����l�����ďI���������߂�
    
    ' ��ƏI���T�̓��͍ςݍH��(�Ώۍ�Ƃ��܂�)���擾
    total_for_days = 0
    For r = setting_schedule_start_row To setting_schedule_end_row
        If ws.Cells(r, setting_worker_col).Value = worker_name Then
            total_for_days = total_for_days + ws.Cells(r, end_date_col).Value
        End If
    Next r
    
    ' 1�������̎��Ԃ͐؂�̂Ă����̂�Int�Ŋۂ߂�
    ' �ۂ߂����ʂ����ɓ��͍ς݂ƈقȂ�ꍇ�͍X�V����
    entered_end_date = Int(CDate(ws.Cells(target_row, setting_end_work_date_col).Value))
    end_date = Int(CDate(ws.Cells(setting_date_row, end_date_col).Value) + total_for_days)
    If entered_end_date <> end_date Then
        ws.Cells(target_row, setting_end_work_date_col).Value = end_date
        ws.Cells(target_row, setting_end_work_date_col).Font.ColorIndex = 3
    End If
    
SKIP_INPUT_DATE:
    
End Sub


Sub GetHoliday()

    'Dim hList() As HolidayList
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

End Sub
