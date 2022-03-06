Attribute VB_Name = "Module1"
Option Explicit

' Setting
Const SCHEDULE_SETTING_SHEET As String = "schedule_macro"
Const SETTING_WB_NAME_CELL As String = "B1"
Const SETTING_WS_NAME_CELL As String = "B2"
Const SETTING_SCHEDULE_START_ROW_CELL As String = "B3"
Const SETTING_SCHEDULE_END_ROW_CELL As String = "B4"
Const SETTING_SCHEDULE_START_COL_CELL As String = "B5"  ' 計画表開始列
Const SETTING_SCHEDULE_END_COL_CELL As String = "B6"    ' 計画表終了列
Const SETTING_SCHEDULE_WORK_DAYS_ROW_CELL As String = "B7"
Const SETTING_SCHEDULE_DATE_ROW_CELL As String = "B8"
Const SETTING_SCHEDULE_WORKER_COL_CELL As String = "B9"
Const SETTING_SCHEDULE_REQUIRED_FOR_INPUT_COL_CELL As String = "B10"
Const SETTING_SCHEDULE_START_WORK_DATE_COL_CELL As String = "B11"   ' 作業開始日
Const SETTING_SCHEDULE_END_WORK_DATE_COL_CELL As String = "B12" ' 作業終了日
Const SETTING_SCHEDULE_BASE_WORKING_HOURS_PER_DAY_ROW_CELL As String = "B13" ' １日の作業時間(１週間平均)：基準
Const SETTING_SCHEDULE_WORKING_HOURS_PER_DAY_ROW_CELL As String = "B14" ' １日の作業時間(１週間平均)


' 設定パラメータ
Dim setting_wb_name As String
Dim setting_ws_name As String
Dim setting_schedule_start_row As Integer   ' 計画表開始行
Dim setting_schedule_end_row As Integer     ' 計画表開始行
Dim setting_schedule_start_col As Integer   ' 計画表開始列
Dim setting_schedule_end_col As Integer     ' 計画表終了列
Dim setting_work_days_row As Integer
Dim setting_date_row As Integer
Dim setting_worker_col As Integer
Dim setting_required_for_input_col As Integer
Dim setting_start_work_date_col As Integer
Dim setting_end_work_date_col As Integer
Dim setting_base_working_hours_per_day As Double
Dim setting_working_hours_per_day As Double


' 作業用変数
Dim wb As Workbook
Dim macro_ws, ws As Worksheet

Dim worker_name As String           ' 作業者名
Dim act_row, act_col As Long

' Undo用データ
Type UndoData
    saved_row As Integer    ' 保存済み行
    start_work_date As Date ' 作業開始日
    end_work_date As Date   ' 作業終了日
    work_day() As Double    ' 入力工数
End Type

Dim undo_data As UndoData

' 休日リスト
Const HOLIDAY_LIST_SIZE As Integer = 30

Type HolidayList
    day As Date
    holiday As Boolean
End Type

'
' マクロエントリ：期間を入力する
'
Sub EntryInputPeriod()
Attribute EntryInputPeriod.VB_ProcData.VB_Invoke_Func = "S\n14"

    Initial
    InputPeriod
    InputWorkStartEndDate (act_row)

End Sub

'
' マクロエントリ：開始日、終了日を入力する
'
Sub EntryInputDate()
Attribute EntryInputDate.VB_ProcData.VB_Invoke_Func = "D\n14"

    Initial
    InputWorkStartEndDate (act_row)

End Sub

'
' マクロエントリ：Undoデータ読み出し
'
Sub EntryLoadUndoData()
Attribute EntryLoadUndoData.VB_ProcData.VB_Invoke_Func = "Z\n14"

    LoadUndoData

End Sub


'
' マクロエントリ：開始日、終了日補正(1日未満の日付入力を切り捨てる)
'
Sub EntryCorrectionDate()
Attribute EntryCorrectionDate.VB_ProcData.VB_Invoke_Func = "C\n14"
    Dim r As Integer
    Dim correct_start_date, entered_start_date As Date  ' 作業開始日付
    Dim correct_end_date, entered_end_date As Date      ' 作業終了日付
    
    Initial
    For r = setting_schedule_start_row To setting_schedule_end_row
        
        ' 開始日補正
        entered_start_date = CDate(ws.Cells(r, setting_start_work_date_col).Value)
        correct_start_date = Int(entered_start_date)
        
        ' 入力の無い日付はスキップ
        If entered_start_date <> 0 Then
            ws.Cells(r, setting_start_work_date_col).Value = correct_start_date
    
            ' 念の為の日付の変更チェック 1日未満を切り捨てるだけなので、変わらないはず
            If ws.Cells(r, setting_start_work_date_col).Value <> Int(entered_start_date) Then
                ws.Cells(r, setting_start_work_date_col).Font.ColorIndex = 5
            End If
        End If
        
        
        ' 終了日補正
        entered_end_date = CDate(ws.Cells(r, setting_end_work_date_col).Value)
        correct_end_date = Int(entered_end_date)
        
        ' 入力の無い日付はスキップ
        If entered_end_date <> 0 Then
            ws.Cells(r, setting_end_work_date_col).Value = correct_end_date
            
            ' 念の為の日付の変更チェック 1日未満を切り捨てるだけなので、変わらないはず
            If ws.Cells(r, setting_end_work_date_col).Value <> Int(entered_end_date) Then
                ws.Cells(r, setting_end_work_date_col).Font.ColorIndex = 5
            End If
        End If
    
    Next r

End Sub


'
' 初期化
'
Sub Initial()
    
    Dim undo_work_day_array_num As Integer  ' Redo 工数配列数

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
    setting_base_working_hours_per_day = macro_ws.Range(SETTING_SCHEDULE_BASE_WORKING_HOURS_PER_DAY_ROW_CELL).Value
    setting_working_hours_per_day = macro_ws.Range(SETTING_SCHEDULE_WORKING_HOURS_PER_DAY_ROW_CELL).Value

    ' 作業用変数初期化
    Set wb = Workbooks(setting_wb_name)
    Set ws = wb.Worksheets(setting_ws_name)
    act_row = ActiveCell.Row
    act_col = ActiveCell.Column
    worker_name = ws.Cells(act_row, setting_worker_col).Value
    
    undo_work_day_array_num = setting_schedule_end_col - setting_schedule_start_col + 1
    undo_data.saved_row = 0
    undo_data.start_work_date = 0
    undo_data.end_work_date = 0
    ReDim undo_data.work_day(undo_work_day_array_num)

End Sub


'
' 期間を入力する
'
Sub InputPeriod()
Attribute InputPeriod.VB_ProcData.VB_Invoke_Func = " \n14"

    Dim r, c, n, i As Integer
    
    Dim total_for_work As Double        ' 入力済み工数(作業：行)
    Dim total_for_days As Double        ' 入力済み工数(日：列)
    
    Dim required_for_input As Double    ' 入力必要工数
    Dim work_days_for_week As Double    ' １週間にかける工数
    Dim input_work_days As Double       ' 入力工数
    
    ' 期間のUndoDataを保存
    SaveUndoData
        
    ' 入力済み工数(作業：行方向合計)取得
    total_for_work = 0
    For c = setting_schedule_start_col To setting_schedule_end_col
        total_for_work = total_for_work + ws.Cells(act_row, c).Value
    Next c
    
    ' 残りの入力必要工数取得 入力必要工数から入力済み工数(日：行方向合計)を引く
    required_for_input = ws.Cells(act_row, setting_required_for_input_col) - total_for_work

    ' 工数入力 (入力必要工数分繰り返し)
    c = 0
    Do
        ' 入力済みのセルはスキップ
        If ws.Cells(act_row, act_col + c).Value <> "" Then
            GoTo CONTINUE
        End If
    
        ' 入力工数を残りの入力必要工数で初期化
        input_work_days = required_for_input
        
        ' 入力済み工数(日：列合計)取得
        total_for_days = 0
        For r = setting_schedule_start_row To setting_schedule_end_row
            If ws.Cells(r, setting_worker_col).Value = worker_name Then
                ' 切り上げて0になるのを防ぐ
'                total_for_days = total_for_days + WorksheetFunction.RoundUp(ws.Cells(r, act_col + c).Value, 2)
                total_for_days = total_for_days + ws.Cells(r, act_col + c).Value
            End If
        Next r
        
        ' １週間にかける工数を算出 １日にかける工数で補正する
        work_days_for_week = ws.Cells(setting_work_days_row, act_col + c).Value * _
            (setting_working_hours_per_day / setting_base_working_hours_per_day)
        
        ' １週間にかける工数を補正 入力済み工数(日：列合計)を引くことで残り入力可能な工数となる
        work_days_for_week = work_days_for_week - total_for_days
        
        ' 入力工数補正：１週間にかける工数を超えている場合は、１週間にかける工数に丸める
        If input_work_days > work_days_for_week Then
            input_work_days = work_days_for_week
        End If
        
        ' 入力工数 を入力する
        If input_work_days > 0 Then
            ' 入力
            ws.Cells(act_row, act_col + c).Value = input_work_days
            
            ' 入力必要工数を減算
            required_for_input = required_for_input - input_work_days
        End If
        
CONTINUE:
        ' 次の列(週)へ
        c = c + 1
    
    Loop While required_for_input > 0 And c <= setting_schedule_end_row

End Sub


'
' 作業開始日/終了日入力
'
Sub InputWorkStartEndDate(ByVal target_row As Long)
    
    Dim r, c As Integer ' Loop用
    
    Dim start_date_col As Integer   ' 作業開始週の列
    Dim end_date_col As Integer     ' 作業終了週の列
    Dim total_for_days As Double    ' 入力済み工数(日)
    
    Dim correct_date As Date
    Dim start_date, entered_start_date As Date  ' 作業開始日付
    Dim end_date, entered_end_date As Date      ' 作業終了日付

    ' 作業開始、作業終了日の週の列数を求める
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
        '開始日が見つからない場合は開始日、終了日入力をスキップ
        GoTo SKIP_INPUT_DATE
    End If

    
    ' 開始日 他作業の入力済み工数を考慮して開始日を決める
    
    ' 作業開始週の他作業の入力済み工数を取得
    total_for_days = 0
    For r = setting_schedule_start_row To setting_schedule_end_row
        If ws.Cells(r, setting_worker_col).Value = worker_name Then
            If r <> target_row Then
                total_for_days = total_for_days + ws.Cells(r, start_date_col).Value
            End If
        End If
    Next r

    ' 他作業の入力済み工数を１日にかける工数で補正する
    total_for_days = total_for_days * (setting_base_working_hours_per_day / setting_working_hours_per_day)

    ' 1日未満の時間は切り捨てたいのでIntで丸める
    ' 丸めた結果が既に入力済みと異なる場合は更新する
    entered_start_date = Int(CDate(ws.Cells(target_row, setting_start_work_date_col).Value))
    start_date = Int(CDate(ws.Cells(setting_date_row, start_date_col).Value) + total_for_days)
    
    ' 土日を考慮 翌日にする
    While Weekday(start_date) = 1 Or Weekday(start_date) = 7
        start_date = start_date + 1
    Wend
    
    If entered_start_date <> start_date Then
        ws.Cells(target_row, setting_start_work_date_col).Value = start_date
        ws.Cells(target_row, setting_start_work_date_col).Font.ColorIndex = 3
    End If
    
    ' 終了日 入力済み工数を考慮して終了日を決める
    
    ' 作業終了週の他作業の入力済み工数を取得
    total_for_days = 0
    For r = setting_schedule_start_row To setting_schedule_end_row
        If ws.Cells(r, setting_worker_col).Value = worker_name Then
            total_for_days = total_for_days + ws.Cells(r, end_date_col).Value
        End If
    Next r
    
    ' 他作業の入力済み工数を１日にかける工数で補正する
    total_for_days = total_for_days * (setting_base_working_hours_per_day / setting_working_hours_per_day)

    ' 1日未満の時間は切り捨てたいのでIntで丸める
    ' 丸めた結果が既に入力済みと異なる場合は更新する
    entered_end_date = Int(CDate(ws.Cells(target_row, setting_end_work_date_col).Value))
    end_date = Int(CDate(ws.Cells(setting_date_row, end_date_col).Value) + total_for_days)
    
    ' 土日を考慮 翌日にする
    While Weekday(end_date) = 1 Or Weekday(end_date) = 7
        end_date = end_date + 1
    Wend

    If entered_end_date <> end_date Then
        ws.Cells(target_row, setting_end_work_date_col).Value = end_date
        ws.Cells(target_row, setting_end_work_date_col).Font.ColorIndex = 3
    End If
    
SKIP_INPUT_DATE:
    
End Sub

'
' マクロエントリ：UnDoデータ保存
'
Sub SaveUndoData()
    Dim i As Integer
    
    If act_row > setting_schedule_end_row Then
        Exit Sub
    End If

    ' 作業開始日／作業終了日を保存
    undo_data.start_work_date = ws.Cells(act_row, setting_start_work_date_col).Value
    undo_data.end_work_date = ws.Cells(act_row, setting_end_work_date_col).Value

    ' 期間工数を保存
    For i = 0 To UBound(undo_data.work_day) - 1
        undo_data.work_day(i) = ws.Cells(act_row, setting_schedule_start_col + i).Value
    Next i
    
    ' 保存済み行設定
    undo_data.saved_row = act_row

End Sub

'
' UnDoデータ読み出し
'
Sub LoadUndoData()
Attribute LoadUndoData.VB_ProcData.VB_Invoke_Func = "Z\n14"
    Dim i As Integer

    If undo_data.saved_row = act_row Then
    
        ' 作業開始日／作業終了日を保存
        If undo_data.start_work_date <> 0 Then
            ws.Cells(undo_data.saved_row, setting_start_work_date_col).Value = undo_data.start_work_date
        Else
            ws.Cells(undo_data.saved_row, setting_start_work_date_col).Value = ""
        End If
        
        If undo_data.end_work_date <> 0 Then
            ws.Cells(undo_data.saved_row, setting_end_work_date_col).Value = undo_data.end_work_date
        Else
            ws.Cells(undo_data.saved_row, setting_end_work_date_col).Value = ""
        End If
    
        ' 期間工数を保存
        For i = 0 To UBound(undo_data.work_day) - 1
            If undo_data.work_day(i) <> 0 Then
                ws.Cells(undo_data.saved_row, setting_schedule_start_col + i).Value = undo_data.work_day(i)
            Else
                ws.Cells(undo_data.saved_row, setting_schedule_start_col + i).Value = ""
            End If
        Next i
        
        ' 保存済み行クリア
        undo_data.saved_row = 0
    
    End If

End Sub


Sub GetHoliday()

    'Dim hList() As HolidayList
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

End Sub
