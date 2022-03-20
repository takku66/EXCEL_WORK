Attribute VB_Name = "MainSheetModule"
Option Explicit


Dim obj As Object

'シート上のコントロールを格納するクラス
Dim sheetCtl As ControleClass

'コントロールを格納するコレクションオブジェクト
Dim ctlDic As Dictionary

'*********************************************************
'オプションボタンのクリック
'*********************************************************
Sub Option_Button_Click()
    Call BtnAbleToggleInit
End Sub

'*********************************************************
'シート上のボタンの使用状態の切り替えを行う（初期化処理）
'*********************************************************
Sub BtnAbleToggleInit()

    'シート上のコントロールオブジェクトクラスを取得する
    Set sheetCtl = New ControleClass
    'コレクションオブジェクトを取得する
    Set ctlDic = New Dictionary
    'シート上のコントロールを格納したコレクションオブジェクトを取得する
    Set ctlDic = sheetCtl.GetControleDic
    
    Dim optBtn As OptionButton
    Dim dropDown As dropDown
    
    Dim key As Variant
    
    
    For Each key In ctlDic
    
    If key Like "Option_*" Then
    Set optBtn = ctlDic.Item(key)
        Select Case key
            Case OPTION_TIMELY_NAME
                Set dropDown = ctlDic.Item(DROPDOWN_TIMELY_NAME)
            Case OPTION_DATE_NAME
                Set dropDown = ctlDic.Item(DROPDOWN_DATE_NAME)
            Case Else
                '例外処理
        End Select
        
        If optBtn.Value = xlOn Then
            dropDown.Enabled = True
        ElseIf optBtn.Value = xlOff Then
            dropDown.Enabled = False
        End If
        
    End If
        
    Next key
    
End Sub


'*********************************************************
'シート上のボタン動作による使用状態の切り替えを行う（ボタン押下時処理）
'*********************************************************
Sub BtnAbleToggleByClick(targetName As String)

    'シート上のコントロールオブジェクトクラスを取得する
    Set sheetCtl = New ControleClass
    'コレクションオブジェクトを取得する
    Set ctlDic = New Dictionary
    'シート上のコントロールを格納したコレクションオブジェクトを取得する
    Set ctlDic = sheetCtl.GetControleDic

    Dim optBtn As OptionButton
    
    

    
    
End Sub

