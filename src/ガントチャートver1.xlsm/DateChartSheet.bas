Attribute VB_Name = "DateChartSheet"
'*****************************************************************
'ライン作成を関数化（予定ライン）
'*****************************************************************
Function MAKECHARTLINE_SCHEDULED(startTerm As Range, endTerm As Range)
    
    Static preStartTerm As Date
    Static preEndTerm As Date
    
    'If preStartTerm = startTerm And preEndTerm = endTerm Then
    '    Exit Function
    'End If
    
    preStartTerm = startTerm.Value
    preEndTerm = endTerm.Value
    
    
    '処理対象のラインオブジェクトの行番号
    Dim lineRowNo As Integer
    lineRowNo = startTerm.row
    
    'ラインインスタンスの生成
    Dim lineCl As New LineClass
    Set lineCl = New LineClass
    
    'ラインオブジェクト名
    Dim lineName As String
    
    'ラインオブジェクト
    Dim lineObj As Shape
    
    'ライン生成時の一時範囲指定用
    Dim tmpRange As Range
    
    'モード設定
    Dim mode As Integer
    mode = SCHEDULED_MODE
    
    '予定用ラインの存在フラグ
    Dim isExistSL As Boolean
    isExistSL = lineCl.IsExistsOnSheetWithMode(lineRowNo, mode)
    
    
    '予定開始と予定終了セルに入力がなければ、削除処理
    If startTerm.Value = "" And endTerm.Value = "" And isExistSL Then
        lineName = lineCl.GetShapeName(lineRowNo, mode)
        Set lineObj = lineCl.GetLineObj(lineName)
        Call lineCl.deleteLine(lineObj)
        Exit Function
    ElseIf startTerm.Value = "" And endTerm.Value = "" Then
        Exit Function
    End If
    
    'カレンダーインスタンスの生成
    Dim calCl As New CalendarClass
    Set calCl = New CalendarClass
    
    'ラインの左端を設定
    Dim leftCol As Integer
    Set tmpRange = calCl.GetDateRngByDate(startTerm.Value)
    leftCol = tmpRange.Column
    'ラインの右端を設定
    Dim rightCol As Integer
    Set tmpRange = calCl.GetDateRngByDate(endTerm.Value)
    rightCol = tmpRange.Column

    '日付セルに入力された値をもとに、ラインオブジェクトの範囲を設定
    Dim settingRange As Range
    Set settingRange = ActiveSheet.Range(Cells(lineRowNo, leftCol).Address, Cells(lineRowNo, rightCol).Address)
    
    If isExistSL Then
        lineName = lineCl.GetShapeName(lineRowNo, mode)
        Set lineObj = lineCl.GetLineObj(lineName)
        Call lineCl.ModifyLine(lineObj, settingRange, mode)
    Else
        Call lineCl.MakeLineByRange(settingRange, mode)
    End If
    
    
    'MsgBox tmpRange.Column
    'MsgBox "最初列" & startTerm.Column & "最初列" & endTerm.Column
    'MsgBox Year(startTerm)
End Function

'*****************************************************************
'ライン作成を関数化（実績ライン）
'*****************************************************************
Function MAKECHARTLINE_RESULT(startTerm As Range, endTerm As Range, resultRate As Long, planedRate As Long)

    Static preStartTerm As Date
    Static preEndTerm As Date
    
    'If preStartTerm = startTerm And preEndTerm = endTerm Then
    '    Exit Function
    'End If
    
    preStartTerm = startTerm.Value
    preEndTerm = endTerm.Value
    

    '処理対象のラインオブジェクトの行番号
    Dim lineRowNo As Integer
    lineRowNo = startTerm.row
    
    'ラインインスタンスの生成
    Dim lineCl As New LineClass
    Set lineCl = New LineClass
    
    'ラインオブジェクト名
    Dim lineName As String
    
    'ラインオブジェクト
    Dim lineObj As Shape
    
    'ライン生成時の一時範囲指定用
    Dim tmpRange As Range
    
    'モード設定
    Dim mode As Integer
    If Not (IsNumeric(resultRate)) Or Not (IsNumeric(planedRate)) Then
        MsgBox "進捗率 もしくは 予定されていた進捗の値が不正です。"
        Exit Function
    ElseIf resultRate < 0 Or resultRate > 100 Then
        MsgBox "進捗率は 0～100 の数値を入力してください。"
        Exit Function
    ElseIf resultRate < planedRate Then
        mode = MINUS_RESULT_MODE
    ElseIf resultRate >= planedRate Then
        mode = PLUS_RESULT_MODE
    End If
    
    'プラス進捗ラインの存在フラグ
    Dim isExistPLUS As Boolean
    isExistPLUS = lineCl.IsExistsOnSheetWithMode(lineRowNo, PLUS_RESULT_MODE)
    
    'マイナス進捗ラインの存在フラグ
    Dim isExistMINUS As Boolean
    isExistMINUS = lineCl.IsExistsOnSheetWithMode(lineRowNo, MINUS_RESULT_MODE)
    
    
    '予定開始と予定終了セルに入力がなければ、削除処理
    If resultRate = 0 And isExistPLUS Then
        lineName = lineCl.GetShapeName(lineRowNo, PLUS_RESULT_MODE)
        Set lineObj = lineCl.GetLineObj(lineName)
        Call lineCl.deleteLine(lineObj)
        Exit Function
    ElseIf resultRate = 0 And isExistMINUS Then
        lineName = lineCl.GetShapeName(lineRowNo, MINUS_RESULT_MODE)
        Set lineObj = lineCl.GetLineObj(lineName)
        Call lineCl.deleteLine(lineObj)
        Exit Function
    ElseIf resultRate = 0 Then
        Exit Function
    End If
    
    'カレンダーインスタンスの生成
    Dim calCl As New CalendarClass
    Set calCl = New CalendarClass
    
    
    'ラインの左端を設定
    Dim leftCol As Integer
    Set tmpRange = calCl.GetDateRngByDate(startTerm.Value)
    leftCol = tmpRange.Column
    'ラインの右端を設定
    Dim rightCol As Integer
    Set tmpRange = calCl.GetDateRngByDate(endTerm.Value)
    rightCol = tmpRange.Column


    
    Dim settingRange As Range
    Set settingRange = ActiveSheet.Range(Cells(lineRowNo, leftCol).Address, Cells(lineRowNo, rightCol).Address)
    
    If isExistPLUS Then
        lineName = lineCl.GetShapeName(lineRowNo, PLUS_RESULT_MODE)
        Set lineObj = lineCl.GetLineObj(lineName)
        Call lineCl.ModifyLine(lineObj, settingRange, mode)
        Call lineCl.SetLineWidth(lineObj, resultRate, settingRange)
    ElseIf isExistMINUS Then
        lineName = lineCl.GetShapeName(lineRowNo, MINUS_RESULT_MODE)
        Set lineObj = lineCl.GetLineObj(lineName)
        Call lineCl.ModifyLine(lineObj, settingRange, mode)
        Call lineCl.SetLineWidth(lineObj, resultRate, settingRange)
    Else
        Call lineCl.MakeLineByRange(settingRange, mode)
        lineName = lineCl.GetShapeName(lineRowNo, mode)
        Set lineObj = lineCl.GetLineObj(lineName)
        Call lineCl.SetLineWidth(lineObj, resultRate, settingRange)
    End If
    
End Function

'*****************************************************************
'ライン作成を関数化（予定進捗）
'*****************************************************************
Function MAKECHARTLINE_SCHEDULED_STATUS(startTerm As Range, endTerm As Range, planedRate As Long)

    Static preStartTerm As Date
    Static preEndTerm As Date
    
    preStartTerm = startTerm.Value
    preEndTerm = endTerm.Value
    

    '処理対象のラインオブジェクトの行番号
    Dim lineRowNo As Integer
    lineRowNo = startTerm.row
    
    'ラインインスタンスの生成
    Dim lineCl As New LineClass
    Set lineCl = New LineClass
    
    'ラインオブジェクト名
    Dim lineName As String
    
    'ラインオブジェクト
    Dim lineObj As Shape
    
    'ライン生成時の一時範囲指定用
    Dim tmpRange As Range
    
    'モード設定
    Dim mode As Integer
    mode = SCHEDULED_STATUS_MODE
    
    '予定進捗ラインの存在フラグ
    Dim isExistSSL As Boolean
    isExistSSL = lineCl.IsExistsOnSheetWithMode(lineRowNo, mode)
    
    
    If Not (IsNumeric(planedRate)) Then
        MsgBox "予定されていた進捗の値が不正です。"
        Exit Function
    End If
    
    '予定開始と予定終了セルに入力がなければ、削除処理
    If planedRate = 0 And isExistSSL Then
        lineName = lineCl.GetShapeName(lineRowNo, mode)
        Set lineObj = lineCl.GetLineObj(lineName)
        Call lineCl.deleteLine(lineObj)
        Exit Function
    ElseIf planedRate = 0 Then
        Exit Function
    End If
    
    'カレンダーインスタンスの生成
    Dim calCl As New CalendarClass
    Set calCl = New CalendarClass
    
    
    'ラインの左端を設定
    Dim leftCol As Integer
    Set tmpRange = calCl.GetDateRngByDate(startTerm.Value)
    leftCol = tmpRange.Column
    'ラインの右端を設定
    Dim rightCol As Integer
    Set tmpRange = calCl.GetDateRngByDate(endTerm.Value)
    rightCol = tmpRange.Column


    
    Dim settingRange As Range
    Set settingRange = ActiveSheet.Range(Cells(lineRowNo, leftCol).Address, Cells(lineRowNo, rightCol).Address)
    
    If isExistSSL Then
        lineName = lineCl.GetShapeName(lineRowNo, mode)
        Set lineObj = lineCl.GetLineObj(lineName)
        Call lineCl.ModifyLine(lineObj, settingRange, mode)
        Call lineCl.SetLineWidth(lineObj, planedRate, settingRange)
    Else
        Call lineCl.MakeLineByRange(settingRange, mode)
        lineName = lineCl.GetShapeName(lineRowNo, mode)
        Set lineObj = lineCl.GetLineObj(lineName)
        Call lineCl.SetLineWidth(lineObj, planedRate, settingRange)
    End If
    
End Function


'*****************************************************************
'カレンダーの設定を反映する
'*****************************************************************
Sub SettingCalendar()
    'カレンダーインスタンスを生成
    Dim calCl As CalendarClass
    Set calCl = New CalendarClass
    
    '開始年月日を取得
    Dim startYmd As Date
    startYmd = ActiveSheet.Cells(CALENDAR_CONF_START_ROW, CALENDAR_CONF_START_COLUMN)
    
    '終了年月日を取得
    Dim endYmd As Date
    endYmd = ActiveSheet.Cells(CALENDAR_CONF_END_ROW, CALENDAR_CONF_END_COLUMN)
    'カレンダーを一旦削除
    Call calCl.DeleteCalendar
    '入力値に基づいて作成
    Call calCl.SettingCalendar(startYmd, endYmd)
    
End Sub




