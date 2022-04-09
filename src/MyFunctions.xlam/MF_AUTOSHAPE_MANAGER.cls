VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MF_AUTOSHAPE_MANAGER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***********************************
' 図形をコントロールするクラス
' < 本クラスが提供する機能 >
'   1. ブック内に存在する図形の名前や入力値、ハイパーリンク値などの一覧管理
'   2. 図形に入力されている値の検索・置換
'   3. 既に配置されている図形の検索・置換
'***********************************


' 図形を管理するシートオブジェクト
Private CONTROL_SHEET As Worksheet

' 図形を扱うサービスクラス
Private mf_shp_service As MF_SHAPE_SERVICE

Dim mf_autoshape_map As Object


' 各管理機能に必要な、各情報の列位置や、読み取り開始位置を定義する
Const NAME_COL = 1
Const CONTENT_COL = 2
Const HYPERLINK_COL = 3
Const SUBHYPERLINK_COL = 4
Const TEMPLATE_AUTOSHAPE_COL = 5
Const READ_ROW = 5
Const TITLE_ROW = READ_ROW - 1

' 絞り込み
Const FILTER_BOX_TITLE_COL = 1
Const FILTER_BOX_VALUE_COL = 2
Const FILTER_BOX_ROW = 2

' 除外文字列
Const EXCLUDE_STR_TITLE_COL = 1
Const EXCLUDE_STR_VALUE_COL = 2
Const EXCLUDE_STR_ROW = 3

Const NAME_TITLE = "図形名"
Const CONTENT_TITLE = "入力値"
Const HYPERLINK_TITLE = "他ブックへのハイパーリンク"
Const SUBHYPERLINK_TITLE = "同ブック内のハイパーリンク"
Const TEMPLATE_AUTOSHAPE_TITLE = "雛形図形"

Const FILTER_BOX_TITLE = "絞り込み"
Const EXCLUDE_STR_TITLE = "除外文字"

' デフォルトの管理シート名
Const DEFAULT_CONTROL_SHEETNAME = "_ASCSHT_"




' ==============================
' 初期化処理
' ==============================

' 初期化
Public Sub Class_Initialize()

    If IS_EXIST_CONTROL_SHEET Then
        Set CONTROL_SHEET = Worksheets(DEFAULT_CONTROL_SHEETNAME)
    Else
        CREATE_TEMPLATE_SHEET
    End If
    
    Set mf_autoshape_map = CreateObject("Scripting.Dictionary")
    Set mf_shp_service = New MF_SHAPE_SERVICE
    Call mf_shp_service.CacheAutoShapeMap(mf_autoshape_map)
    
End Sub

' シートをアクティブ化して表示させる
Public Sub ACTIVATE()
    CONTROL_SHEET.ACTIVATE
End Sub

' 管理表の名前を取得する
Public Function GET_SHEET_NAME()
    GET_SHEET_NAME = CONTROL_SHEET.name
End Function

' 管理表が存在しているか
Public Function IS_EXIST_CONTROL_SHEET()
    IS_EXIST_CONTROL_SHEET = False
    Dim ws As Worksheet
    For Each ws In Worksheets
        If ws.name = DEFAULT_CONTROL_SHEETNAME Then
            IS_EXIST_CONTROL_SHEET = True
            Exit Function
        End If
    Next ws
End Function



' ==============================
' 管理シートのテンプレート作成
' ==============================

' 最上位メソッド
Public Sub CREATE_TEMPLATE_SHEET()

    ' シートの追加
    Set CONTROL_SHEET = ActiveWorkbook.Sheets.Add
    CONTROL_SHEET.name = DEFAULT_CONTROL_SHEETNAME
    
    ' シート内の一覧表を作成
    CREATE_CONTROL_LIST
    
    ' シート内にボタンを配置
    ' 　・図形の収集と更新
    ' 　・図形の作成と置換
    CREATE_CONTROL_BUTTONS
End Sub


' 管理表の一覧表を作成する
Private Sub CREATE_CONTROL_LIST()

    ' 各表項目のタイトル名や、幅を設定する
    CONTROL_SHEET.Cells(TITLE_ROW, NAME_COL).Value = NAME_TITLE
    CONTROL_SHEET.Cells(TITLE_ROW, CONTENT_COL).Value = CONTENT_TITLE
    CONTROL_SHEET.Cells(TITLE_ROW, HYPERLINK_COL).Value = HYPERLINK_TITLE
    CONTROL_SHEET.Cells(TITLE_ROW, SUBHYPERLINK_COL).Value = SUBHYPERLINK_TITLE
    CONTROL_SHEET.Cells(TITLE_ROW, TEMPLATE_AUTOSHAPE_COL).Value = TEMPLATE_AUTOSHAPE_TITLE
    CONTROL_SHEET.Cells(FILTER_BOX_ROW, FILTER_BOX_TITLE_COL).Value = FILTER_BOX_TITLE
    CONTROL_SHEET.Cells(EXCLUDE_STR_ROW, EXCLUDE_STR_TITLE_COL).Value = EXCLUDE_STR_TITLE
    
    Columns(NAME_COL).ColumnWidth = 15
    Columns(CONTENT_COL).ColumnWidth = 30
    Columns(HYPERLINK_COL).ColumnWidth = 50
    Columns(SUBHYPERLINK_COL).ColumnWidth = 30
    Columns(TEMPLATE_AUTOSHAPE_COL).ColumnWidth = 20
    Rows.RowHeight = 40
    
    ' 表のスタイル定義範囲を、Rangeで取得しておく
    Dim area As Range
    Dim titleArea As Range
    Set area = CONTROL_SHEET.Range(CONTROL_SHEET.Cells(TITLE_ROW, NAME_COL), CONTROL_SHEET.Cells(READ_ROW + 200, TEMPLATE_AUTOSHAPE_COL))
    Set titleArea = CONTROL_SHEET.Range(CONTROL_SHEET.Cells(TITLE_ROW, NAME_COL), CONTROL_SHEET.Cells(TITLE_ROW, TEMPLATE_AUTOSHAPE_COL))
    
    ' 罫線周りの設定
    area.Borders(xlDiagonalDown).LineStyle = xlNone
    area.Borders(xlDiagonalUp).LineStyle = xlNone
    With area.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With area.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With area.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With area.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With area.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With area.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With titleArea.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    With titleArea
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    titleArea.Font.Bold = True
End Sub

' 管理用のボタンを配置する
Private Sub CREATE_CONTROL_BUTTONS()
    CREATE_UPDATE_BUTTON
    CREATE_REFLECT_BUTTON
End Sub
Private Function CREATE_UPDATE_BUTTON()     ' 一覧情報更新ボタン
    Dim updateBtn As Variant
    Set updateBtn = CONTROL_SHEET.Buttons.Add(0, 0, 80, 25)
    updateBtn.OnAction = "UPDATE_CONTROL_LIST"
    updateBtn.name = "__MFCTRL_UPDATELIST"
    updateBtn.Characters.Text = "情報更新"
    With updateBtn.Characters(Start:=1, Length:=3).Font
        .name = "Meiryo UI"
        .FontStyle = "標準"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 1
    End With
    Set CREATE_UPDATE_BUTTON = updateBtn
End Function
Private Function CREATE_REFLECT_BUTTON()     ' 作成・置換ボタン
    Dim reflectBtn As Variant
    Set reflectBtn = CONTROL_SHEET.Buttons.Add(100, 0, 80, 25)
    reflectBtn.OnAction = "REFLECT_AUTOSHAPE"
    reflectBtn.name = "__MFCTRL_UPDATELIST"
    reflectBtn.Characters.Text = "作成・置換"
    With reflectBtn.Characters(Start:=1, Length:=3).Font
        .name = "Meiryo UI"
        .FontStyle = "標準"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 1
    End With
    Set CREATE_REFLECT_BUTTON = reflectBtn
End Function


' ==============================
' 管理表に配置している各ボタンの処理定義
' ==============================

' 情報収集ボタン
Public Sub UPDATE_CONTROL_LIST()
    Dim answer As Integer
    answer = MsgBox("編集中の情報は破棄されます。続けますか？", vbOKCancel)
    If answer = vbCancel Then
        Exit Sub
    End If
    
    CLEAR_CONTROL_LIST
    ' キャッシュオブジェクトを更新して、図形情報から一覧表に反映させる
    Dim filterWord As String
    filterWord = CONTROL_SHEET.Cells(FILTER_BOX_ROW, FILTER_BOX_VALUE_COL).Value
    Dim excludeWord As String
    excludeWord = CONTROL_SHEET.Cells(EXCLUDE_STR_ROW, EXCLUDE_STR_VALUE_COL).Value
    Call mf_shp_service.CacheAutoShapeMap(mf_autoshape_map, filterWord, excludeWord)
    Call REFLECT_TO_LIST(mf_autoshape_map)
End Sub
Public Sub REFLECT_TO_LIST(mf_autoshape_map As Object)
    Dim maxRow As Long
    maxRow = mf_autoshape_map.Count + READ_ROW

    ' 雛形図形以外の情報を２次元配列で保持する
    Dim CONTROL_LIST_EXCLUDE_SHAPE() As Variant
    ReDim CONTROL_LIST_EXCLUDE_SHAPE(1 To maxRow, NAME_COL To SUBHYPERLINK_COL)
    ' 雛形図形列は、(1,MaxRow)の２次元配列で保持する
    Dim TEMPLATE_SHAPE_LIST() As Variant
    ReDim TEMPLATE_SHAPE_LIST(1 To maxRow, 1 To 1)
        
    ' 各２次元配列を反映させるためのRangeオブジェクト
    Dim CONTROL_LIST_RANGE_EXCLUDE_SHAPE As Range
    Set CONTROL_LIST_RANGE_EXCLUDE_SHAPE = CONTROL_SHEET.Range(CONTROL_SHEET.Cells(READ_ROW, NAME_COL), CONTROL_SHEET.Cells(maxRow, SUBHYPERLINK_COL))
    Dim TAMPLATE_SHAPE_LIST_RANGE As Range
    Set TAMPLATE_SHAPE_LIST_RANGE = CONTROL_SHEET.Range(CONTROL_SHEET.Cells(READ_ROW, TEMPLATE_AUTOSHAPE_COL), CONTROL_SHEET.Cells(maxRow, TEMPLATE_AUTOSHAPE_COL))
    
    ' Shapeオブジェクトを取得して、
    ' そのプロパティは文字列として２次元配列へ
    ' 図形オブジェクト自体は、図形としての２次元配列へ格納する
    Dim mf_shp As MF_SHAPE_DATA
    Dim shp As Shape
    Dim copyShp As Shape
    Dim template_shape_cell As Range
    Dim itemlist() As Variant
    itemlist = mf_autoshape_map.Items
    Dim dicLoop As Long
    Dim readrow As Long
    
    On Error Resume Next
    For dicLoop = 1 To mf_autoshape_map.Count
        
        Set mf_shp = itemlist(dicLoop - 1)(0)
        Set shp = mf_shp.GET_SHAPE
        CONTROL_LIST_EXCLUDE_SHAPE(dicLoop, NAME_COL) = shp.name

        CONTROL_LIST_EXCLUDE_SHAPE(dicLoop, CONTENT_COL) = shp.TextFrame2.TextRange.Characters.Text
        CONTROL_LIST_EXCLUDE_SHAPE(dicLoop, HYPERLINK_COL) = shp.hyperlink.Address
        CONTROL_LIST_EXCLUDE_SHAPE(dicLoop, SUBHYPERLINK_COL) = shp.hyperlink.SubAddress
        
        readrow = READ_ROW + dicLoop - 1
        Set template_shape_cell = CONTROL_SHEET.Cells(readrow, TEMPLATE_AUTOSHAPE_COL)
        shp.Copy
        template_shape_cell.Select
        CONTROL_SHEET.Paste
        With Selection
            .name = shp.name
            .TextFrame2.TextRange.Characters.Text = shp.TextFrame2.TextRange.Characters.Text
            '.Top = template_shape_cell.Top
            '.Left = template_shape_cell.Left
            .Width = 120
            .Height = 38
        End With
        
        TEMPLATE_SHAPE_LIST(dicLoop, 1) = Selection
    Next dicLoop
On Error GoTo 0
    
    ' 一覧表に反映
    CONTROL_LIST_RANGE_EXCLUDE_SHAPE.Value = CONTROL_LIST_EXCLUDE_SHAPE
    
    MsgBox "図形収集処理完了"
End Sub

' 管理表に設定されている図形情報を更新 または 新規作成する。
' 新規作成された図形は、管理表に出力される
Public Sub REFLECT_AUTOSHAPE()
    
    Dim loopCnt As Long
    Dim maxCnt As Long
    maxCnt = MF_ROW_TAILINCOLUMN(CONTROL_SHEET.Range( _
                                CONTROL_SHEET.Cells(READ_ROW, NAME_COL), _
                                CONTROL_SHEET.Cells(READ_ROW, NAME_COL)))

    Dim name As String
    Dim content As String
    Dim hyperlink As String
    Dim subhyperlink As String
    Dim template As Shape
    
    ' 表から各図形情報を取得し、
    ' 名前に合致する図形がなければ作成、
    ' あれば、情報を更新する
    For loopCnt = READ_ROW To maxCnt
        name = CONTROL_SHEET.Cells(loopCnt, NAME_COL).Value
        content = CONTROL_SHEET.Cells(loopCnt, CONTENT_COL).Value
        hyperlink = CONTROL_SHEET.Cells(loopCnt, HYPERLINK_COL).Value
        subhyperlink = CONTROL_SHEET.Cells(loopCnt, SUBHYPERLINK_COL).Value
        Set template = SHAPE_INCELL(CONTROL_SHEET.Cells(loopCnt, TEMPLATE_AUTOSHAPE_COL))
        
        Call mf_shp_service.CacheAutoShapeMap(mf_autoshape_map)
        
        If template Is Nothing Then
            MsgBox "[ " & name & " ] の雛形図形が存在しません。　行：" & loopCnt
            GoTo Continue
        End If
        Call mf_shp_service.CREATE_MFSHAPE(mf_autoshape_map, CONTROL_SHEET, name, content, hyperlink, subhyperlink, template)
Continue:
    Next loopCnt
End Sub

' 指定のセルに含まれる図形を返す
Public Function SHAPE_INCELL(cell As Range) As Shape
    Dim shp As Shape
    
    ' 全ての図形を取得し、その図形が指定のセルに含まれているか確認する
    ' 最初の図形のみを返す
    For Each shp In cell.Worksheet.Shapes
    
        Dim shpContainsCell As Range
        Set shpContainsCell = Range(shp.TopLeftCell, shp.BottomRightCell)
        
        If Not (Intersect(shpContainsCell, cell) Is Nothing) Then
            If Intersect(shpContainsCell, cell).Address = shpContainsCell.Address Then
                Set SHAPE_INCELL = shp
                Exit Function
            End If
        End If
        
    Next
End Function

Public Sub CLEAR_CONTROL_LIST()
    ' 更新対象範囲のセルと、図形を一括でクリアする
    Dim lastRow As Long
    lastRow = MF_ROW_TAILINCOLUMN(CONTROL_SHEET.Cells(1, NAME_COL))
    If lastRow >= READ_ROW Then
        Dim deleteRange As Range
        Set deleteRange = CONTROL_SHEET.Range(CONTROL_SHEET.Cells(READ_ROW, NAME_COL), CONTROL_SHEET.Cells(lastRow, TEMPLATE_AUTOSHAPE_COL))
        deleteRange.ClearContents
        Call DELETE_SHAPE_IN_RANGE(deleteRange)
    End If
End Sub
