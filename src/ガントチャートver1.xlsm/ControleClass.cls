VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ControleClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************************
'シート上のコントロールオブジェクトを管理するクラス
'*****************************************************************
Option Explicit

'一時オブジェクト
Dim obj As Object

'シート上のコントロールオブジェクトを取得する
Dim objDic As Dictionary


'*****************************************************************
'コンストラクタ
'シート上のオブジェクトを全て取得する
'*****************************************************************
Sub Class_Initialize()
    
    Set objDic = New Dictionary
    
        For Each obj In ActiveSheet.Shapes
            obj.Select
            If TypeName(Selection) Like "Option*" Then
                objDic.Add obj.name, ActiveSheet.OptionButtons(obj.name)
            ElseIf TypeName(Selection) Like "DropDown*" Then
                objDic.Add obj.name, ActiveSheet.DropDowns(obj.name)
            ElseIf TypeName(Selection) Like "CheckBox*" Then
                objDic.Add obj.name, ActiveSheet.CheckBoxes(obj.name)
            ElseIf TypeName(Selection) Like "TextBox*" Then
                objDic.Add obj.name, ActiveSheet.TextBoxes(obj.name)
            End If
            
        Next obj
    ActiveSheet.Cells(1, 1).Select
End Sub

'*****************************************************************
'コレクションオブジェクトのゲッター
'*****************************************************************
Function GetControleDic()
    Set GetControleDic = objDic
End Function



