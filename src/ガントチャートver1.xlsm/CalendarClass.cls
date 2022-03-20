VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CalendarClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'ワークシートオブジェクト
Dim sheet As Worksheet

'***********年月日シート用******************
'年Dictionaryデータ
'   キー：年
'   値  ：月Dictionary
Dim yearDic As Dictionary

'月Dictionaryデータ
'   キー：月
'   値  ：日Dictionary
Dim monthDic As Dictionary

'日Dictionaryデータ
'   キー：日
'   値  ：Rangeオブジェクト
Dim dayDicForYmd As Dictionary

'日付を記載しているRangeオブジェクト
Dim dateRng As Range

'日付Dictionaryデータ
'   キー：日付(yyyy/MM/dd)
'   値  ：Rangeオブジェクト
Dim dateDic As Dictionary
'******************************************


'***********時間単位シート用******************
'日Dictionaryデータ
'   キー：日
'   値  ：オブジェクト
Dim dayDicForTimely As Dictionary


'********************************************



'*****************************************************************
'コンストラクタ
'シート上のカレンダー設定を取得する
'*****************************************************************
Sub Class_Initialize()
    Set sheet = ActiveSheet
    
    '日の最終列を取得
    Dim dayLastCol As Integer
    dayLastCol = sheet.Cells(CHART_DAY_ROW, Columns.count).End(xlToLeft).Column
    
    '日のRangeオブジェクトを取得する
    Set dateRng = sheet.Range(sheet.Cells(CHART_DAY_ROW, CHART_DAY_COLUMN).Address, sheet.Cells(CHART_DAY_ROW, dayLastCol).Address)
    
    
    Set dateDic = New Dictionary
    '日付とRangeオブジェクトのセットをDictionaryオブジェクトへ格納する
    Dim Var As Variant
    For Each Var In dateRng
        dateDic.Add Var.Value, Var
    Next Var
End Sub

'*****************************************************************
'渡された日付をもとに、Rangeオブジェクトを返す
'*****************************************************************
Function GetDateRngByDate(dt As Date)
    Set GetDateRngByDate = dateDic.Item(dt)
End Function


'*****************************************************************
'渡された年月日をもとに、カレンダーを作成する
'*****************************************************************
Function SettingCalendar(startDt As Date, endDt As Date)
    '開始～終了年月日の差分
    Dim diffDt As Integer
    diffDt = endDt - startDt
    
    'カレンダーへ反映させるための年月日配列
    Dim ymdArray() As Variant
    ReDim ymdArray(diffDt)
    
    '年月日を設定する範囲
    Dim settingRange As Range
    Set settingRange = sheet.Range(sheet.Cells(CHART_DAY_ROW, CHART_DAY_COLUMN).Address, sheet.Cells(CHART_DAY_ROW, CHART_DAY_COLUMN + diffDt).Address)
    
    '雛形の列をコピーする
    Columns(BASE_CALENDAR_COLUMN).Copy
    settingRange.EntireColumn.PasteSpecial (xlPasteAll)
    Application.CutCopyMode = False
    
    
    'ループ用変数
    Dim cntDt As Integer
    cntDt = 0
    
    Dim tmpDt As Date
    tmpDt = startDt
    
    For cntDt = 0 To diffDt
        ymdArray(cntDt) = (tmpDt)
        tmpDt = tmpDt + 1
    Next cntDt
    
    settingRange = ymdArray
    sheet.Range(sheet.Cells(1, CHART_DAY_COLUMN).Address, sheet.Cells(1, CHART_DAY_COLUMN + diffDt).Address).EntireColumn.Hidden = False
    
End Function


'*****************************************************************
'カレンダーを削除する
'*****************************************************************
Function DeleteCalendar()
    '日の最終列を取得
    Dim dayLastCol As Integer
    dayLastCol = sheet.Cells(CHART_DAY_ROW, Columns.count).End(xlToLeft).Column
    
    sheet.Range(Columns(CHART_DAY_COLUMN), Columns(dayLastCol)).Delete
    
End Function



