Attribute VB_Name = "TestModule"
Dim categoryList As Dictionary
Dim lineData As LineClass

'*****************************************************************
'
'*****************************************************************
Sub main()
    Dim sheet As Worksheet
    Set sheet = Worksheets(MAIN_SHEET_NAME)
    
    Set lineData = New LineClass
    lineData.MakeChartSheet
    
End Sub

'*****************************************************************
'存在判定がうまくいっているかテスト用
'*****************************************************************
Sub submain()
    Dim sheet As Worksheet
    Set sheet = Worksheets(MAIN_SHEET_NAME)
    
    Dim Var As Variant
    
    Set categoryList = GetTaskCategory
    
    For Each Var In categoryList
    
        If Var Is Nothing Or IsEmpty(Var) Then
            Exit For
        End If
        
        Debug.Print categoryList(Var)
        MsgBox IsExistsOnSheet(categoryList(Var))
    Next Var
End Sub


Sub test()
    Dim sh As Shape
    
    
    On Error GoTo Catch
        Set sh = ActiveSheet.Shapes("CHART_LINE_*_7")
        
        MsgBox "あります"
        Exit Sub
Catch:
    MsgBox "ありません"


    
    
End Sub

