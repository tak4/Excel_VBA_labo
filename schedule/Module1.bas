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
Const WORK_DAYS_ROW = 6     ' 営業日を記載した行数
Const DATE_ROW = 8  ' 日付を記載した行数

Const WORKER_COL As Integer = 2 ' 作業者名の列数
Const REQUIRED_FOR_INPUT_COL As Integer = 3 ' 入力必要工数の列数
Const START_WORK_DATE_COL As Integer = 4
Const END_WORK_DATE_COL As Integer = 5

Const INPUT_WORK_DAY As Integer = 5 ' 入力基準工数

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
    
    Dim worker_name As String       ' 作業者名
    Dim work_days_for_week As Double   ' 1週間にかける工数
    Dim act_row, act_col As Double
    
    Dim total_for_days As Double       ' 入力済み工数(日)
    Dim total_for_work As Double       ' 入力済み工数(作業)
    Dim required_for_input As Double ' 入力必要工数
    Dim can_input As Double           ' 入力可能工数
    Dim input_work_days As Double  ' 入力工数
    
    Dim start_date_col As Integer   ' 作業開始週の列
    Dim end_date_col As Integer ' 作業終了週の列
    Dim start_date As Date
    Dim end_date As Date
    
    
    ' 入力基準工数
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

    ' 入力必要工数
    total_for_work = 0
    For c = SCHEDULE_START_COL To SCHEDULE_END_COL
        total_for_work = total_for_work + ws.Cells(act_row, c).Value
    Next c
    
    required_for_input = ws.Cells(act_row, REQUIRED_FOR_INPUT_COL) - total_for_work
    
    ' 入力必要工数分繰り返し
    c = 0
    Do
        ' 入力済みのセルはスキップ
        If ws.Cells(act_row, act_col + c).Value <> "" Then
            GoTo CONTINUE
        End If
    
        ' 入力済み工数(日)
        total_for_days = 0
        For r = SCHEDULE_START_ROW To SCHEDULE_END_ROW
            If ws.Cells(r, WORKER_COL).Value = worker_name Then
                total_for_days = total_for_days + ws.Cells(r, act_col + c).Value
            End If
        Next r
        
        ' 入力可能工数(日) 営業日から入力済み工数(日)を引く
        can_input = ws.Cells(WORK_DAYS_ROW, act_col + c).Value - total_for_days
        If can_input < 0 Then
            can_input = 0
        End If
        
        ' 入力工数を入力可能工数(日)で補正
        If can_input > work_days_for_week Then
            input_work_days = work_days_for_week
        Else
            input_work_days = can_input
        End If
        
        ' 入力工数を入力必要工数で補正
        If required_for_input < input_work_days Then
            input_work_days = required_for_input
        End If
                
        ' 入力
        If input_work_days > 0 Then
            ws.Cells(act_row, act_col + c).Value = input_work_days
            required_for_input = required_for_input - input_work_days
        End If
        
CONTINUE:
        ' 次の列(週)へ
        c = c + 1
    
    Loop While required_for_input > 0 And c <= SCHEDULE_END_ROW
    
    ' 開始日/終了日入力
    ' 開始終了日の週の列数を求める
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
    
    ' 作業開始週の他作業の入力済み工数を取得
    total_for_days = 0
    For r = SCHEDULE_START_ROW To SCHEDULE_END_ROW
        If r <> act_row Then
            total_for_days = total_for_days + ws.Cells(r, start_date_col).Value
        End If
    Next r
    
    ws.Cells(act_row, START_WORK_DATE_COL).Value = ws.Cells(DATE_ROW, start_date_col).Value + total_for_days
    ws.Cells(act_row, END_WORK_DATE_COL).Value = ws.Cells(DATE_ROW, end_date_col).Value + ws.Cells(act_row, end_date_col).Value
    

End Sub
