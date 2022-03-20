Attribute VB_Name = "MainSheetConst"

'==================Common==================
'メインワークシートの名前
Public Const MAIN_SHEET_NAME = "ガントチャート作成フォーム"

'年単位ガントチャートシートの名前
Public Const PER_YEAR_SHEET_NAME = "年単位ガントチャート"

'月単位ガントチャートシートの名前
Public Const PER_MONTH_SHEET_NAME = "月単位ガントチャート"

'日単位ガントチャートシートの名前
Public Const PER_DAY_SHEET_NAME = "日単位ガントチャート"

'時間単位ガントチャートシートの名前
Public Const PER_TIMELY_SHEET_NAME = "時間単位ガントチャート"

'ディクショナリーオブジェクト作成時の文字列
Public Const STR_SCRIPTING_DICTIONARY = "Scripting.Dictionary"

'==========================================


'==================LineShape Name==================
'ラインオブジェクト作成時の名前
Public Const CHART_LINE_NAME = "CHART_LINE"
'予定進捗オブジェクト作成時の名前
Public Const SCHEDULED_STATUS_LINE_NAME = "SCHEDULED_STATUS_LINE"
'===========================================


'==================CTL NAME==================
'時間単位ラジオボタン
Public Const OPTION_TIMELY_NAME = "Option_Timely"
'年月日単位ラジオボタン
Public Const OPTION_DATE_NAME = "Option_Date"
'時間単位ドロップダウン
Public Const DROPDOWN_TIMELY_NAME = "DropDown_Timely"
'年月日単位ドロップダウン
Public Const DROPDOWN_DATE_NAME = "DropDown_Date"
'===========================================


'==================CTL MODE==================
'予定モード
Public Const SCHEDULED_MODE = 1
'実績（良好）モード
Public Const PLUS_RESULT_MODE = 2
'実績（不良）モード
Public Const MINUS_RESULT_MODE = 3
'予定進捗モード
Public Const SCHEDULED_STATUS_MODE = 4


'予定モード_ライン名
Public Const SCHEDULED_MODE_NAME = "S"
'実績モード_ライン名
Public Const RESULT_MODE_NAME = "R"
'==========================================


'==================COLUMN DEFINE==================
'列：予定開始
Public Const START_SCHEDULED_COLUMN = 8
'列：予定終了
Public Const END_SCHEDULED_COLUMN = 9
'列：所要
Public Const COST_STEPS_COLUMN = 10
'列：実績開始
Public Const START_RESULT_COLUMN = 11
'列：実績終了
Public Const END_RESULT_COLUMN = 12
'列：進捗率
Public Const STATUS_COLUMN = 13
'列：進捗予定
Public Const EST_STATUS_COLUMN = 15
'=================================================


'==================CHART SHEET DEFINE==================
'カレンダー単位設定値
Public Const CALENDAR_SETTING_ROW = 1
Public Const CALENDAR_SETTING_COLUMN = 27

'年単位設定文字列
Public Const CALENDAR_SETTING_YEAR_STR = "年単位"
'月単位設定文字列
Public Const CALENDAR_SETTING_MONTH_STR = "月単位"
'日単位設定文字列
Public Const CALENDAR_SETTING_DAY_STR = "日単位"
'時間単位設定文字列
Public Const CALENDAR_SETTING_TIMELY_STR = "時間単位"


'年項目：列、行
Public Const CHART_YEAR_ROW = 2
Public Const CHART_YEAR_COLUMN = 30
'月項目：列、行
Public Const CHART_MONTH_ROW = 3
Public Const CHART_MONTH_COLUMN = 30
'日項目：列、行
Public Const CHART_DAY_ROW = 4
Public Const CHART_DAY_COLUMN = 30

'曜日項目：列、行
Public Const CHART_WEEK_ROW = 5
Public Const CHART_WEEK_COLUMN = 30

'列：雛形
Public Const BASE_CALENDAR_COLUMN = 29

'カレンダー設定用 開始年月日：列、行
Public Const CALENDAR_CONF_START_ROW = 2
Public Const CALENDAR_CONF_START_COLUMN = 4
'カレンダー設定用 終了年月日：列、行
Public Const CALENDAR_CONF_END_ROW = 3
Public Const CALENDAR_CONF_END_COLUMN = 4
'=================================================




