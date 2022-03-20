VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MF_SHAPE_SERVICE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Sub CREATE_MFSHAPE(reflectSht As Worksheet, name As String, content As String, hyperlink As String, subhyperlink As String, template As Shape)
    
    Dim mf_shp As MF_SHAPE_DATA
    
    ' 名前で図形を検索し、
    ' 存在していれば、プロパティを更新
    ' 存在していなければ、引数のテンプレートを元に、新しく図形を作成する
    Set mf_shp = FIND_SHAPE_OF_ALL(name)
    
    If mf_shp Is Nothing Then
        Set mf_shp = New MF_SHAPE_DATA
        Call mf_shp.SET_SHEET(reflectSht)
    End If
    Call mf_shp.SET_NAME(name)
    Call mf_shp.SET_CONTENT(content)
    Call mf_shp.SET_HYPERLINK(hyperlink)
    Call mf_shp.SET_SUBHYPERLINK(subhyperlink)
    
    Set CREATE_MFSHAPE = REFLECT_SHAPE(mf_shp, template)

End Sub

Public Function CHANGE_CONTENT(mf_shp As MF_SHAPE_DATA, content As String)
    Call mf_shp.SET_CONTENT(content)
    Set CHANGE_CONTENT = mf_shp
End Function

Public Function CHANGE_HYPERLINK(mf_shp As MF_SHAPE_DATA, hyperlink As String)
    Call mf_shp.SET_HYPERLINK(hyperlink)
    Set CHANGE_HYPERLINK = mf_shp
End Function


' 指定されたシート・名前の図形を取得する
Public Function FIND_SHAPE_OF_SHEET(sht As Worksheet, name As String)
On Error GoTo NOT_FOUND
    Set FIND_SHAPE_OF_SHEET = sht.Shapes(name)
    Exit Function
NOT_FOUND:
    Set FIND_SHAPE_OF_SHEET = Nothing
    Exit Function
End Function

Public Function FIND_SHAPE_OF_ALL(name As String) As MF_SHAPE_DATA
    Dim sht As Worksheet
    Dim shp As Shape
    Dim mf_shp As MF_SHAPE_DATA
    Set mf_shp = New MF_SHAPE_DATA
    For Each sht In ActiveWorkbook.Sheets
        Set shp = FIND_SHAPE_OF_SHEET(sht, name)
        If shp Is Nothing Then
            GoTo Continue
        Else
            Call mf_shp.SET_SHAPE(shp)
            Call mf_shp.SET_SHEET(sht)
            Set FIND_SHAPE_OF_ALL = mf_shp
            Exit Function
        End If
Continue:
    Next sht
End Function


' 図形情報を反映させる
Private Function REFLECT_SHAPE(mf_shp As MF_SHAPE_DATA, template As Shape)

    Dim copyShp As Shape
    If mf_shp.GET_SHAPE Is Nothing Then
        ' 新規作成
        Set copyShp = template.Duplicate
        With copyShp
            .name = mf_shp.GET_NAME
            .TextFrame2.TextRange.Characters.Text = mf_shp.GET_CONTENT
            .Top = 0
            .Left = 400
            .Width = 100
            .Height = 40
        End With
    Else
        ' 更新
        Set copyShp = mf_shp.GET_SHAPE
        With copyShp
            .TextFrame2.TextRange.Characters.Text = mf_shp.GET_CONTENT
        End With
    End If
    
    mf_shp.GET_SHEET.Hyperlinks.Add _
                                Anchor:=copyShp, _
                                Address:=mf_shp.GET_HYPERLINK, _
                                SubAddress:=mf_shp.GET_SUBHYPERLINK
                                
    Call mf_shp.SET_SHAPE(copyShp)
    Set REFLECT_SHAPE = mf_shp
End Function


