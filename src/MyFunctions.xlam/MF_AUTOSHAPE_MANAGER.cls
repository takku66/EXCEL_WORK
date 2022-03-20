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


' 各管理機能に必要な、各情報の列位置や、読み取り開始位置を定義する
Const NAME_COL = 1
Const CONTENT_COL = 2
Const HYPERLINK_COL = 3
Const SUBHYPERLINK_COL = 4
Const TEMPLATE_AUTOSHAPE_COL = 5
Const READ_ROW = 4
Const TITLE_ROW = READ_ROW - 1

Const NAME_TITLE = "図形名"
Const CONTENT_TITLE = "入力値"
Const HYPERLINK_TITLE = "他ブックへのハイパーリンク"
Const SUBHYPERLINK_TITLE = "同ブック内のハイパーリンク"
Const TEMPLATE_AUTOSHAPE_TITLE = "雛形図形"

' デフォルトの管理シート名
Const DEFAULT_CONTROL_SHEETNAME = "_ASCSHT_"



'******************************
' 初期化処理
'
'******************************
Public Sub Class_Initialize()

    If IS_EXIST_CONTROL_SHEET Then
        Set CONTROL_SHEET = Worksheets(DEFAULT_CONTROL_SHEETNAME)
    Else
        CREATE_TEMPLATE_SHEET
    End If
    
    Set mf_shp_service = New MF_SHAPE_SERVICE
    
End Sub
Public Sub ACTIVATE()
    CONTROL_SHEET.ACTIVATE
End Sub
Public Function GET_SHEET_NAME()
    GET_SHEET_NAME = CONTROL_SHEET.name
End Function
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




' 管理シートのテンプレートを作成する
Public Sub CREATE_TEMPLATE_SHEET()

    ' シートの追加
    Set CONTROL_SHEET = ActiveWorkbook.Sheets.Add
    CONTROL_SHEET.name = DEFAULT_CONTROL_SHEETNAME
    
    ' シート内の一覧表を作成
    CREATE_CONTROL_LIST
    
    ' シート内にボタンを配置
    ' 収集ボタン
    ' ボタンの置換
    CREATE_CONTROL_BUTTONS
End Sub

' 管理表を作成
Private Sub CREATE_CONTROL_LIST()
    CONTROL_SHEET.Cells(TITLE_ROW, NAME_COL).Value = NAME_TITLE
    CONTROL_SHEET.Cells(TITLE_ROW, CONTENT_COL).Value = CONTENT_TITLE
    CONTROL_SHEET.Cells(TITLE_ROW, HYPERLINK_COL).Value = HYPERLINK_TITLE
    CONTROL_SHEET.Cells(TITLE_ROW, SUBHYPERLINK_COL).Value = SUBHYPERLINK_TITLE
    CONTROL_SHEET.Cells(TITLE_ROW, TEMPLATE_AUTOSHAPE_COL).Value = TEMPLATE_AUTOSHAPE_TITLE
    Columns(NAME_COL).ColumnWidth = 15
    Columns(CONTENT_COL).ColumnWidth = 30
    Columns(HYPERLINK_COL).ColumnWidth = 50
    Columns(SUBHYPERLINK_COL).ColumnWidth = 30
    Rows.RowHeight = 30
    
    Dim area As Range
    Dim titleArea As Range
    Set area = CONTROL_SHEET.Range(CONTROL_SHEET.Cells(TITLE_ROW, NAME_COL), CONTROL_SHEET.Cells(READ_ROW + 200, TEMPLATE_AUTOSHAPE_COL))
    Set titleArea = CONTROL_SHEET.Range(CONTROL_SHEET.Cells(TITLE_ROW, NAME_COL), CONTROL_SHEET.Cells(TITLE_ROW, TEMPLATE_AUTOSHAPE_COL))
    
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
    CREATE_REPLACE_BUTTON
End Sub
Private Function CREATE_UPDATE_BUTTON()     ' 一覧情報更新ボタン
    Dim updateBtn As Variant
    Set updateBtn = CONTROL_SHEET.Buttons.Add(0, 0, 80, 25)
    updateBtn.OnAction = "UPDATE_CONTROL_LIST"
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
Private Function CREATE_REPLACE_BUTTON()     ' 作成・置換ボタン
    Dim replaceBtn As Variant
    Set replaceBtn = CONTROL_SHEET.Buttons.Add(100, 0, 80, 25)
    replaceBtn.OnAction = "REFLECT_AUTOSHAPE"
    replaceBtn.Characters.Text = "作成・置換"
    With replaceBtn.Characters(Start:=1, Length:=3).Font
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
    Set CREATE_REPLACE_BUTTON = replaceBtn
End Function



' 図形を収集して、管理表に反映させる
Public Sub UPDATE_CONTROL_LIST()
    COLLECT_AUTOSHAPE
End Sub
Public Sub COLLECT_AUTOSHAPE(Optional collectName As String)
    MsgBox "図形収集処理：未実装"
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
        
        Call mf_shp_service.CREATE_MFSHAPE(CONTROL_SHEET, name, content, hyperlink, subhyperlink, template)
        
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

