Attribute VB_Name = "MyFunctions"
Option Explicit

Dim MF_AS_MNG As MF_AUTOSHAPE_MANAGER

'=============================================
' 関数群

' 行内で最後に値がある列数を返す
Public Function MF_COLUMN_TAILINROW(cell As Range)
    MF_COLUMN_TAILINROW = Worksheets(cell.Parent.name).Cells(cell.Row, Columns.Count).End(xlToLeft).Column
End Function

' 列内で最後に値がある行数を返す
Public Function MF_ROW_TAILINCOLUMN(cell As Range)
    MF_ROW_TAILINCOLUMN = Worksheets(cell.Parent.name).Cells(Rows.Count, cell.Column).End(xlUp).Row
End Function

' 行内で最初に値がある列数を返す
Public Function MF_COLUMN_HEADINROW(cell As Range)
    MF_COLUMN_HEADINROW = Worksheets(cell.Parent.name).Cells(cell.Row, 1).End(xlToRight).Column
End Function

' 列内で最初に値がある行数を返す
Public Function MF_ROW_HEADINCOLUMN(cell As Range)
    MF_ROW_HEADINCOLUMN = Worksheets(cell.Parent.name).Cells(1, cell.Column).End(xlDown).Row
End Function

' 分割文字列で分割し、指定の位置の値を返す
Public Function MF_VALUE_FROMSPLIT(cell As Range, splitStr As String, idx As Long)
    Dim ary() As String
    ary = split(cell.Value, splitStr)
    MF_VALUE_FROMSPLIT = ary(idx)
End Function
'=============================================


'=============================================
' 図形管理系

' 図形管理シートを作成する
Public Sub CREATE_AUTOSHAPE_MANAGER()
    INIT_AUTOSHAPE_MANAGER
    If Not (MF_AS_MNG.IS_EXIST_CONTROL_SHEET) Then
        MF_AS_MNG.CREATE_TEMPLATE_SHEET
    End If
    MF_AS_MNG.ACTIVATE
End Sub

' 図形管理オブジェクトを初期化する
Private Function INIT_AUTOSHAPE_MANAGER()
    If MF_AS_MNG Is Nothing Then
        Set MF_AS_MNG = New MF_AUTOSHAPE_MANAGER
    End If
    Set INIT_AUTOSHAPE_MANAGER = MF_AS_MNG
End Function

' 図形情報の一覧を更新する
Public Sub UPDATE_CONTROL_LIST()
    ' 途中でマクロが失敗した場合は、管理用オブジェクトが消える
    ' 暫定対応として、保険の初期化処理
    INIT_AUTOSHAPE_MANAGER
    MF_AS_MNG.UPDATE_CONTROL_LIST
End Sub

' 一覧にある図形情報を、実物の図形に反映させる
' 図形がなければ、作成される
Public Sub REFLECT_AUTOSHAPE()
    ' 途中でマクロが失敗した場合は、管理用オブジェクトが消える
    ' 暫定対応として、保険の初期化処理
    INIT_AUTOSHAPE_MANAGER
    MF_AS_MNG.REFLECT_AUTOSHAPE
End Sub


'=============================================



'=============================================
' 共通

Public Sub SHOW_MF_HELP()
    MF_HELP.Show
End Sub

Public Sub DELETE_SHAPE_IN_RANGE(Optional targetRange As Range)
    Dim shp As Shape
    Dim shprng As Range
    Dim trng As Range
    Set trng = targetRange
    If targetRange Is Nothing Then
        Set trng = Selection
    End If
    
    'シート内の図形をループ
    For Each shp In ActiveSheet.Shapes
        '図形の左上と右下のセル範囲を格納
        Set shprng = Range(shp.TopLeftCell, shp.BottomRightCell)
        '図形のセル範囲と、選択したセルの範囲が、重なっているかを判定
        If Not Intersect(shprng, trng) Is Nothing Then
            shp.Delete '図形を削除
        End If
    Next
End Sub


'=============================================


Public Sub execCache()
    Dim obj As Object
    Set obj = CreateObject("Scripting.Dictionary")
    
    Dim svc As MF_SHAPE_SERVICE
    Set svc = New MF_SHAPE_SERVICE
    Call svc.CacheAutoShapeMap(obj)
    
    Dim Key As Variant
    For Each Key In obj.Keys
        Debug.Print Key
        Debug.Print IsArray(obj.Item(Key))
        
        Dim tmp As Variant
        For Each tmp In obj.Item(Key)
            Debug.Print tmp.GET_SHAPE.name
        Next tmp
    Next Key
End Sub
