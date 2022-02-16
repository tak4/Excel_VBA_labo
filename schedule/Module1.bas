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
    
    ' 設定パラメータ
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
    
    ' 作業用変数
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
    
    Dim hList() As HolidayList

    ' 設定パラメータ 初期化
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

    ' 作業用変数初期化
    Set wb = Workbooks(setting_wb_name)
    Set ws = wb.Worksheets(setting_ws_name)
    act_row = ActiveCell.Row
    act_col = ActiveCell.Column
    worker_name = ws.Cells(act_row, setting_worker_col).Value
    
    ' 休日データ取得
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

    
    ' 入力基準工数
    work_days_for_week = setting_input_work_date

    ' 入力必要工数取得
    total_for_work = 0
    For c = setting_schedule_start_col To setting_schedule_end_col
        total_for_work = total_for_work + ws.Cells(act_row, c).Value
    Next c
    required_for_input = ws.Cells(act_row, setting_required_for_input_col) - total_for_work

    
    ' 工数入力 - 入力必要工数分繰り返し
    c = 0
    Do
        ' 入力済みのセルはスキップ
        If ws.Cells(act_row, act_col + c).Value <> "" Then
            GoTo CONTINUE
        End If
    
        ' 入力済み工数(日)
        total_for_days = 0
        For r = setting_schedule_start_row To setting_schedule_end_row
            If ws.Cells(r, setting_worker_col).Value = worker_name Then
                total_for_days = total_for_days + ws.Cells(r, act_col + c).Value
            End If
        Next r
        
        ' 入力可能工数(日) 営業日から入力済み工数(日)を引く
        can_input = ws.Cells(setting_work_days_row, act_col + c).Value - total_for_days
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
    
    Loop While required_for_input > 0 And c <= setting_schedule_end_row
    
    
    ' 開始日/終了日入力
    
    If update_date = True Then
        
        ' 開始終了日の週の列数を求める
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
        
        ' 開始日
        
        ' 作業開始週の他作業の入力済み工数を取得
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
        
        ' 終了日
        
        ' 作業終了週の他作業の入力済み工数を取得
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


