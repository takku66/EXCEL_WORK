VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LineClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'ワークシートオブジェクト
Dim sheet As Worksheet
'タスク名と行番号のDictionaryデータ
Dim categoryDic As Dictionary

'図形オブジェクト
Dim shObj As Shape


'*****************************************************************
'クラスコンストラクタ
'*****************************************************************
Sub Class_Initialize()
    '作成対象のシートオブジェクトを設定
    '※実際は、ActiveSheetを使っていった方がいい？
    Set sheet = ActiveSheet
        
End Sub


'*****************************************************************
'ガントチャートを作成する
'*****************************************************************
Sub MakeChartSheet()

End Sub


'*****************************************************************
'Rangeオブジェクトをもとに、ガントチャート用のラインを描画する
'*****************************************************************
Sub MakeLineByRange(ByVal rng As Range, mode As Integer)
    
    '作成するラインオブジェクト
    Dim lineObj As Shape
    'ラインを作成する範囲
    Dim lineRange As Range
    Set lineRange = rng
    
    
    Select Case mode
        Case SCHEDULED_STATUS_MODE
            'ラインオブジェクトを設定する
            Set lineObj = ActiveSheet.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1)
        Case Else
            'ラインオブジェクトを設定する
            Set lineObj = ActiveSheet.Shapes.AddLine(1, 1, 1, 1)
    End Select
    
    Call InitLine(lineObj, lineRange, mode)
    
End Sub

'*****************************************************************
'対象のラインオブジェクトを修正する
'*****************************************************************
Sub ModifyLine(targetLine As Shape, ByVal lineRange As Range, mode As Integer)
    
    Call InitLine(targetLine, lineRange, mode)
    
End Sub


'*****************************************************************
'対象のラインオブジェクトを削除する
'*****************************************************************
Sub deleteLine(ByVal lineObj As Shape)
    lineObj.Delete
End Sub


'*****************************************************************
'シート上のラインオブジェクトを全て削除する
'*****************************************************************
Sub deleteAllLine()
    
End Sub

'*****************************************************************
'対象のラインオブジェクトを取得する
'*****************************************************************
Function GetLineObj(lineName As String)
    Dim lineObj As Shape
    
    On Error GoTo Catch
        Set lineObj = ActiveSheet.Shapes(lineName)
        Set GetLineObj = lineObj
        Exit Function
Catch:
    Set GetLineObj = Nothing
    
End Function



'*****************************************************************
'ラインの書式設定（色彩、モード名、高さ）
'*****************************************************************
Function InitLine(ByVal lineObj As Shape, ByVal lineRange As Range, mode As Integer)
    'ラインオブジェクト名
    Dim lineName As String
    
    '始点X
    Dim sx As Single
    '始点Y
    Dim sy As Single
    '終点X
    Dim ex As Single
    '終点Y
    Dim ey As Single
    
    '始点X, 終点X を設定する
    sx = lineRange.Left
    ex = lineRange.Left + lineRange.Width
    
    
    '色彩設定用（赤）
    Dim red As Integer
    '色彩設定用（緑）
    Dim green As Integer
    '色彩設定用（青）
    Dim blue As Integer
    '透明度
    Dim tp As Integer
    
    
    '引数に受け取ったモードをもとに、ラインを設定する
    Select Case mode
        '予定入力用
        Case SCHEDULED_MODE
            sy = lineRange.Top + (lineRange.Height / 4) + 4
            ey = lineRange.Top + (lineRange.Height / 4) + 4
            red = 0
            green = 224
            blue = 96
            lineName = SCHEDULED_MODE_NAME & "_" & lineRange.row
        '実績（前納 / オンスケ）用
        Case PLUS_RESULT_MODE
            sy = lineRange.Top + (lineRange.Height / 4 * 3) - 4
            ey = lineRange.Top + (lineRange.Height / 4 * 3) - 4
            red = 0
            green = 96
            blue = 224
            lineName = RESULT_MODE_NAME & "_" & lineRange.row
        '実績（遅延）用
        Case MINUS_RESULT_MODE
            sy = lineRange.Top + (lineRange.Height / 4 * 3) - 4
            ey = lineRange.Top + (lineRange.Height / 4 * 3) - 4
            red = 255
            green = 0
            blue = 0
            lineName = RESULT_MODE_NAME & "_" & lineRange.row
        '予定進捗用
        Case SCHEDULED_STATUS_MODE
            sy = lineRange.Top + (lineRange.Height / 4 * 3) - 6
            ey = lineRange.Top + (lineRange.Height / 4 * 3) - 6
            red = 0
            green = 0
            blue = 0
            lineName = lineRange.row
            tp = 1
        Case Else
    End Select
    
    If lineObj.Type = MsoShapeType.msoLine Then
        'ラインの縦幅
        lineObj.Line.Weight = 6
        'ラインの横幅
        lineObj.Width = lineRange.Width
        'ラインの色彩
        lineObj.Line.ForeColor.RGB = RGB(red, green, blue)
        'ライン名
        lineObj.name = CHART_LINE_NAME & "_" & lineName
        'ラインの位置
        lineObj.Top = sy
        lineObj.Left = sx
    ElseIf lineObj.Type = MsoShapeType.msoAutoShape Then
        'ラインの縦幅
        lineObj.Height = 4
        'ラインの横幅
        lineObj.Width = lineRange.Width
        'ラインの色彩
        lineObj.Line.ForeColor.RGB = RGB(red, green, blue)
        'ラインの枠線
        lineObj.Line.DashStyle = msoLineDash
        'ライン名
        lineObj.name = SCHEDULED_STATUS_LINE_NAME & "_" & lineName
        'ラインの位置
        lineObj.Top = sy
        lineObj.Left = sx
        'ラインの透明度
        lineObj.Fill.Transparency = tp
        '
        lineObj.ZOrder msoSendToBack
    End If
    
    Set InitLine = lineObj
End Function


'*****************************************************************
'ラインの横幅を設定する
'*****************************************************************
Function SetLineWidth(ByVal lineObj As Shape, rate As Long, ByVal lineRange As Range)
    'ラインオブジェクト名
    Dim lineName As String
    
    '始点X
    Dim sx As Single
    '始点Y
    Dim sy As Single
    '終点X
    Dim ex As Single
    '終点Y
    Dim ey As Single

    'ラインの横幅
    lineObj.Width = lineRange.Width * (rate / 100)
    
    Set SetLineWidth = lineObj
End Function


'*****************************************************************
'シート上に記載しているタスク名を全て取得する
'*****************************************************************
Function GetTaskCategory()
    Dim sheet As Worksheet
    Set sheet = Worksheets(MAIN_SHEET_NAME)
    
    Dim taskCol As Integer
    taskCol = 2
    
    Dim startRow As Integer
    startRow = 6
    
    Dim lastRow As Integer
    lastRow = sheet.Cells(Rows.count, taskCol).End(xlUp).row
    
    Set categoryDic = New Dictionary
    
    Dim count As Integer
    
    For count = startRow To lastRow
        categoryDic.Add sheet.Cells(count, taskCol), count
    Next count
    
    Set GetTaskCategory = categoryDic
End Function


'*****************************************************************
'タスク名から、モードを無視したShapeオブジェクト名を取得する
'*****************************************************************
Function GetBothShapeNameByTask(lineRowNo As Integer)
    GetBothShapeNameByTask = CHART_LINE_NAME & "_*_" & lineRowNo
End Function

'*****************************************************************
'検索用のライン名を作成し、返す
'*****************************************************************
Function GetShapeName(lineRowNo As Integer, mode As Integer)
    '戻り値用文字列
    Dim rtnName As String
    
    'モードによって、返却文字列を変更する
    Select Case mode
        Case SCHEDULED_MODE
            rtnName = CHART_LINE_NAME & "_" & SCHEDULED_MODE_NAME & "_" & lineRowNo
        Case PLUS_RESULT_MODE
            rtnName = CHART_LINE_NAME & "_" & RESULT_MODE_NAME & "_" & lineRowNo
        Case MINUS_RESULT_MODE
            rtnName = CHART_LINE_NAME & "_" & RESULT_MODE_NAME & "_" & lineRowNo
        Case SCHEDULED_STATUS_MODE
            rtnName = SCHEDULED_STATUS_LINE_NAME & "_" & lineRowNo
        Case Else
    End Select
    GetShapeName = rtnName
End Function


'*****************************************************************
'タスク名からシートの情報を参照して、ラインのモードを取得する
'*****************************************************************
Function GetMode(lineRowNo As Integer)
    
End Function

'*****************************************************************
'シート上にラインオブジェクトが存在するか判定する
'※ライン番号のみ条件抽出
'*****************************************************************
Function IsExistsOnSheet(lineRowNo As String)
    
    Dim obj As Object
    
    Dim lineName As String
    lineName = GetBothShapeNameByTask(lineRowNo)
    
    'シート上の全Shapeオブジェクトをループ処理
    For Each obj In ActiveSheet.Shapes
        '検索用のライン名に合致した場合
        If obj.name Like lineName Then
            IsExistsOnSheet = True
            Exit Function
        End If
    Next obj
    IsExistsOnSheet = False
End Function

'*****************************************************************
'シート上にラインオブジェクトが存在するか判定する
'※ライン番号とモード指定条件抽出
'*****************************************************************
Function IsExistsOnSheetWithMode(lineRowNo As Integer, mode As Integer)
    
    Dim lineName As String
    lineName = GetShapeName(lineRowNo, mode)
    
    Dim sh As Shape
    
    
    On Error GoTo Catch
        Set sh = ActiveSheet.Shapes(lineName)
        
        IsExistsOnSheetWithMode = True
        Exit Function
Catch:
    
    IsExistsOnSheetWithMode = False
    
End Function


'*****************************************************************
'指定した範囲上に対象のラインオブジェクトが存在するか判定する
'※ライン番号のみ条件指定
'*****************************************************************
Function IsExistsOnRange(lineRowNo As Integer, ByVal rng As Range)
    
    'ループ用オブジェクト
    Dim obj As Object
    
    'ラインNoから検索用のライン名を取得
    Dim lineName As String
    lineName = GetBothShapeNameByTask(lineRowNo)
    
    '対象のラインオブジェクト範囲格納用
    Dim tmprng As Range
    
    'シート上の全Shapeオブジェクトをループ処理
    For Each obj In ActiveSheet.Shapes
        '検索用のライン名に合致した場合
        If obj.name Like lineName Then
            'ラインの範囲を取得
            Set tmprng = ActiveSheet.Range(ActiveSheet.Shapes(obj.name).TopLeftCell, ActiveSheet.Shapes(obj.name).BottomRightCell)
            '範囲の重複チェック
            If Not (Intersect(rng, tmprng) Is Nothing) Then
                '重複していた場合
                IsExistsOnRange = True
                Exit For
            Else
                '重複していない場合
                IsExistsOnRange = False
                Exit For
            End If
        End If
    Next obj

End Function

'*****************************************************************
'指定した範囲上に対象のラインオブジェクトが存在するか判定する
'※ライン番号とモードを条件指定
'*****************************************************************
Function IsExistsOnRangeWithMode(lineRowNo As Integer, ByVal rng As Range, mode As Integer)
    
    'ラインNoから検索用のライン名を取得
    Dim lineName As String
    lineName = GetShapeName(lineRowNo, mode)
    
    '対象のラインオブジェクト範囲格納用
    Dim tmprng As Range
        
    Dim sh As Shape
    
    On Error GoTo Catch
        
        Set sh = ActiveSheet.Shapes(linaName)
        
            'ラインの範囲を取得
            Set tmprng = ActiveSheet.Range(sh.TopLeftCell, sh.BottomRightCell)
            '範囲の重複チェック
            If Not (Intersect(rng, tmprng) Is Nothing) Then
                '重複していた場合
                IsExistsOnRange = True
                Exit For
            Else
                '重複していない場合
                IsExistsOnRange = False
                Exit For
            End If

        Exit Function
Catch:
    IsExistsOnRangeWithMode = False
End Function

