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

'=============================================


