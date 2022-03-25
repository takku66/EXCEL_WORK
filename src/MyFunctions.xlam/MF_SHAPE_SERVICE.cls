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


Public Sub CREATE_MFSHAPE(autoshapeMap As Object, reflectSht As Worksheet, name As String, content As String, hyperlink As String, subhyperlink As String, template As Shape)
    
    Dim mf_shp As Variant
    Dim mf_shplist() As MF_SHAPE_DATA
    
    ' 名前で図形を検索し、
    ' 存在していれば、プロパティを更新
    ' 存在していなければ、引数のテンプレートを元に、新しく図形を作成する
    If autoshapeMap.Exists(name) Then
        ReDim mf_shplist(0 To UBound(autoshapeMap.Item(name)))
        mf_shplist = autoshapeMap.Item(name)
    Else
        ReDim mf_shplist(0)
        Set mf_shplist(0) = New MF_SHAPE_DATA
    End If
    
    For Each mf_shp In mf_shplist
        If mf_shp.GET_SHEET Is Nothing Then
            Call mf_shp.SET_SHEET(reflectSht)
        End If
        Call mf_shp.SET_NAME(name)
        Call mf_shp.SET_CONTENT(content)
        Call mf_shp.SET_HYPERLINK(hyperlink)
        Call mf_shp.SET_SUBHYPERLINK(subhyperlink)
        
        Set mf_shp = REFLECT_SHAPE(mf_shp, template)
    Next mf_shp

End Sub

Public Function CHANGE_CONTENT(mf_shp As MF_SHAPE_DATA, content As String)
    Call mf_shp.SET_CONTENT(content)
    Set CHANGE_CONTENT = mf_shp
End Function

Public Function CHANGE_HYPERLINK(mf_shp As MF_SHAPE_DATA, hyperlink As String)
    Call mf_shp.SET_HYPERLINK(hyperlink)
    Set CHANGE_HYPERLINK = mf_shp
End Function



Public Sub CacheAutoShapeMap(mf_autoshape_map As Object)
    Dim sht As Worksheet
    Dim shp As Shape
    Dim mf_shp As MF_SHAPE_DATA
    Dim mf_shplist() As MF_SHAPE_DATA
    
    Set mf_autoshape_map = CreateObject("Scripting.Dictionary")
    
    Dim name As String
    For Each sht In ActiveWorkbook.Sheets
        For Each shp In sht.Shapes
        
            name = shp.name
            If name Like "__MFCTRL_*" Then
                GoTo Continue
            End If
            
            Set mf_shp = New MF_SHAPE_DATA
            Call mf_shp.SET_SHAPE(shp)
            Call mf_shp.SET_SHEET(sht)
            
            If mf_autoshape_map.Exists(name) Then
                mf_shplist = mf_autoshape_map.Item(name)
                ReDim Preserve mf_shplist(UBound(mf_shplist) + 1)
                Set mf_shplist(UBound(mf_shplist)) = mf_shp
                mf_autoshape_map.Remove name
                mf_autoshape_map.Add name, mf_shplist
            Else
                ReDim mf_shplist(0)
                Set mf_shplist(0) = mf_shp
                mf_autoshape_map.Add name, mf_shplist
            End If
Continue:
        Next shp
    Next sht
End Sub

' 図形情報を反映させる
Private Function REFLECT_SHAPE(mf_shp As Variant, template As Shape)

    Dim copyShp As Shape
    If mf_shp.GET_SHAPE Is Nothing Then
        ' 新規作成
        Set copyShp = template.Duplicate
        With copyShp
            .name = mf_shp.GET_NAME
            .TEXTFRAME2.TextRange.Characters.Text = mf_shp.GET_CONTENT
            .Top = 0
            .Left = 400
            .Width = 100
            .Height = 40
        End With
    Else
        ' 更新
        Set copyShp = mf_shp.GET_SHAPE
        With copyShp
            .TEXTFRAME2.TextRange.Characters.Text = mf_shp.GET_CONTENT
        End With
    End If
    
    mf_shp.GET_SHEET.Hyperlinks.Add _
                                Anchor:=copyShp, _
                                Address:=mf_shp.GET_HYPERLINK, _
                                SubAddress:=mf_shp.GET_SUBHYPERLINK
                                
    If mf_shp.GET_HYPERLINK = "" And mf_shp.GET_SUBHYPERLINK = "" Then
        copyShp.hyperlink.Delete
    End If
                                
    Call mf_shp.SET_SHAPE(copyShp)
    Set REFLECT_SHAPE = mf_shp
End Function


