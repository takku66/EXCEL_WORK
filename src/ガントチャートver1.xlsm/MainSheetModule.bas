Attribute VB_Name = "MainSheetModule"
Option Explicit


Dim obj As Object

'�V�[�g��̃R���g���[�����i�[����N���X
Dim sheetCtl As ControleClass

'�R���g���[�����i�[����R���N�V�����I�u�W�F�N�g
Dim ctlDic As Dictionary

'*********************************************************
'�I�v�V�����{�^���̃N���b�N
'*********************************************************
Sub Option_Button_Click()
    Call BtnAbleToggleInit
End Sub

'*********************************************************
'�V�[�g��̃{�^���̎g�p��Ԃ̐؂�ւ����s���i�����������j
'*********************************************************
Sub BtnAbleToggleInit()

    '�V�[�g��̃R���g���[���I�u�W�F�N�g�N���X���擾����
    Set sheetCtl = New ControleClass
    '�R���N�V�����I�u�W�F�N�g���擾����
    Set ctlDic = New Dictionary
    '�V�[�g��̃R���g���[�����i�[�����R���N�V�����I�u�W�F�N�g���擾����
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
                '��O����
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
'�V�[�g��̃{�^������ɂ��g�p��Ԃ̐؂�ւ����s���i�{�^�������������j
'*********************************************************
Sub BtnAbleToggleByClick(targetName As String)

    '�V�[�g��̃R���g���[���I�u�W�F�N�g�N���X���擾����
    Set sheetCtl = New ControleClass
    '�R���N�V�����I�u�W�F�N�g���擾����
    Set ctlDic = New Dictionary
    '�V�[�g��̃R���g���[�����i�[�����R���N�V�����I�u�W�F�N�g���擾����
    Set ctlDic = sheetCtl.GetControleDic

    Dim optBtn As OptionButton
    
    

    
    
End Sub

