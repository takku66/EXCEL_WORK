Attribute VB_Name = "DateChartSheet"
'*****************************************************************
'���C���쐬���֐����i�\�胉�C���j
'*****************************************************************
Function MAKECHARTLINE_SCHEDULED(startTerm As Range, endTerm As Range)
    
    Static preStartTerm As Date
    Static preEndTerm As Date
    
    'If preStartTerm = startTerm And preEndTerm = endTerm Then
    '    Exit Function
    'End If
    
    preStartTerm = startTerm.Value
    preEndTerm = endTerm.Value
    
    
    '�����Ώۂ̃��C���I�u�W�F�N�g�̍s�ԍ�
    Dim lineRowNo As Integer
    lineRowNo = startTerm.row
    
    '���C���C���X�^���X�̐���
    Dim lineCl As New LineClass
    Set lineCl = New LineClass
    
    '���C���I�u�W�F�N�g��
    Dim lineName As String
    
    '���C���I�u�W�F�N�g
    Dim lineObj As Shape
    
    '���C���������̈ꎞ�͈͎w��p
    Dim tmpRange As Range
    
    '���[�h�ݒ�
    Dim mode As Integer
    mode = SCHEDULED_MODE
    
    '�\��p���C���̑��݃t���O
    Dim isExistSL As Boolean
    isExistSL = lineCl.IsExistsOnSheetWithMode(lineRowNo, mode)
    
    
    '�\��J�n�Ɨ\��I���Z���ɓ��͂��Ȃ���΁A�폜����
    If startTerm.Value = "" And endTerm.Value = "" And isExistSL Then
        lineName = lineCl.GetShapeName(lineRowNo, mode)
        Set lineObj = lineCl.GetLineObj(lineName)
        Call lineCl.deleteLine(lineObj)
        Exit Function
    ElseIf startTerm.Value = "" And endTerm.Value = "" Then
        Exit Function
    End If
    
    '�J�����_�[�C���X�^���X�̐���
    Dim calCl As New CalendarClass
    Set calCl = New CalendarClass
    
    '���C���̍��[��ݒ�
    Dim leftCol As Integer
    Set tmpRange = calCl.GetDateRngByDate(startTerm.Value)
    leftCol = tmpRange.Column
    '���C���̉E�[��ݒ�
    Dim rightCol As Integer
    Set tmpRange = calCl.GetDateRngByDate(endTerm.Value)
    rightCol = tmpRange.Column

    '���t�Z���ɓ��͂��ꂽ�l�����ƂɁA���C���I�u�W�F�N�g�͈̔͂�ݒ�
    Dim settingRange As Range
    Set settingRange = ActiveSheet.Range(Cells(lineRowNo, leftCol).Address, Cells(lineRowNo, rightCol).Address)
    
    If isExistSL Then
        lineName = lineCl.GetShapeName(lineRowNo, mode)
        Set lineObj = lineCl.GetLineObj(lineName)
        Call lineCl.ModifyLine(lineObj, settingRange, mode)
    Else
        Call lineCl.MakeLineByRange(settingRange, mode)
    End If
    
    
    'MsgBox tmpRange.Column
    'MsgBox "�ŏ���" & startTerm.Column & "�ŏ���" & endTerm.Column
    'MsgBox Year(startTerm)
End Function

'*****************************************************************
'���C���쐬���֐����i���у��C���j
'*****************************************************************
Function MAKECHARTLINE_RESULT(startTerm As Range, endTerm As Range, resultRate As Long, planedRate As Long)

    Static preStartTerm As Date
    Static preEndTerm As Date
    
    'If preStartTerm = startTerm And preEndTerm = endTerm Then
    '    Exit Function
    'End If
    
    preStartTerm = startTerm.Value
    preEndTerm = endTerm.Value
    

    '�����Ώۂ̃��C���I�u�W�F�N�g�̍s�ԍ�
    Dim lineRowNo As Integer
    lineRowNo = startTerm.row
    
    '���C���C���X�^���X�̐���
    Dim lineCl As New LineClass
    Set lineCl = New LineClass
    
    '���C���I�u�W�F�N�g��
    Dim lineName As String
    
    '���C���I�u�W�F�N�g
    Dim lineObj As Shape
    
    '���C���������̈ꎞ�͈͎w��p
    Dim tmpRange As Range
    
    '���[�h�ݒ�
    Dim mode As Integer
    If Not (IsNumeric(resultRate)) Or Not (IsNumeric(planedRate)) Then
        MsgBox "�i���� �������� �\�肳��Ă����i���̒l���s���ł��B"
        Exit Function
    ElseIf resultRate < 0 Or resultRate > 100 Then
        MsgBox "�i������ 0�`100 �̐��l����͂��Ă��������B"
        Exit Function
    ElseIf resultRate < planedRate Then
        mode = MINUS_RESULT_MODE
    ElseIf resultRate >= planedRate Then
        mode = PLUS_RESULT_MODE
    End If
    
    '�v���X�i�����C���̑��݃t���O
    Dim isExistPLUS As Boolean
    isExistPLUS = lineCl.IsExistsOnSheetWithMode(lineRowNo, PLUS_RESULT_MODE)
    
    '�}�C�i�X�i�����C���̑��݃t���O
    Dim isExistMINUS As Boolean
    isExistMINUS = lineCl.IsExistsOnSheetWithMode(lineRowNo, MINUS_RESULT_MODE)
    
    
    '�\��J�n�Ɨ\��I���Z���ɓ��͂��Ȃ���΁A�폜����
    If resultRate = 0 And isExistPLUS Then
        lineName = lineCl.GetShapeName(lineRowNo, PLUS_RESULT_MODE)
        Set lineObj = lineCl.GetLineObj(lineName)
        Call lineCl.deleteLine(lineObj)
        Exit Function
    ElseIf resultRate = 0 And isExistMINUS Then
        lineName = lineCl.GetShapeName(lineRowNo, MINUS_RESULT_MODE)
        Set lineObj = lineCl.GetLineObj(lineName)
        Call lineCl.deleteLine(lineObj)
        Exit Function
    ElseIf resultRate = 0 Then
        Exit Function
    End If
    
    '�J�����_�[�C���X�^���X�̐���
    Dim calCl As New CalendarClass
    Set calCl = New CalendarClass
    
    
    '���C���̍��[��ݒ�
    Dim leftCol As Integer
    Set tmpRange = calCl.GetDateRngByDate(startTerm.Value)
    leftCol = tmpRange.Column
    '���C���̉E�[��ݒ�
    Dim rightCol As Integer
    Set tmpRange = calCl.GetDateRngByDate(endTerm.Value)
    rightCol = tmpRange.Column


    
    Dim settingRange As Range
    Set settingRange = ActiveSheet.Range(Cells(lineRowNo, leftCol).Address, Cells(lineRowNo, rightCol).Address)
    
    If isExistPLUS Then
        lineName = lineCl.GetShapeName(lineRowNo, PLUS_RESULT_MODE)
        Set lineObj = lineCl.GetLineObj(lineName)
        Call lineCl.ModifyLine(lineObj, settingRange, mode)
        Call lineCl.SetLineWidth(lineObj, resultRate, settingRange)
    ElseIf isExistMINUS Then
        lineName = lineCl.GetShapeName(lineRowNo, MINUS_RESULT_MODE)
        Set lineObj = lineCl.GetLineObj(lineName)
        Call lineCl.ModifyLine(lineObj, settingRange, mode)
        Call lineCl.SetLineWidth(lineObj, resultRate, settingRange)
    Else
        Call lineCl.MakeLineByRange(settingRange, mode)
        lineName = lineCl.GetShapeName(lineRowNo, mode)
        Set lineObj = lineCl.GetLineObj(lineName)
        Call lineCl.SetLineWidth(lineObj, resultRate, settingRange)
    End If
    
End Function

'*****************************************************************
'���C���쐬���֐����i�\��i���j
'*****************************************************************
Function MAKECHARTLINE_SCHEDULED_STATUS(startTerm As Range, endTerm As Range, planedRate As Long)

    Static preStartTerm As Date
    Static preEndTerm As Date
    
    preStartTerm = startTerm.Value
    preEndTerm = endTerm.Value
    

    '�����Ώۂ̃��C���I�u�W�F�N�g�̍s�ԍ�
    Dim lineRowNo As Integer
    lineRowNo = startTerm.row
    
    '���C���C���X�^���X�̐���
    Dim lineCl As New LineClass
    Set lineCl = New LineClass
    
    '���C���I�u�W�F�N�g��
    Dim lineName As String
    
    '���C���I�u�W�F�N�g
    Dim lineObj As Shape
    
    '���C���������̈ꎞ�͈͎w��p
    Dim tmpRange As Range
    
    '���[�h�ݒ�
    Dim mode As Integer
    mode = SCHEDULED_STATUS_MODE
    
    '�\��i�����C���̑��݃t���O
    Dim isExistSSL As Boolean
    isExistSSL = lineCl.IsExistsOnSheetWithMode(lineRowNo, mode)
    
    
    If Not (IsNumeric(planedRate)) Then
        MsgBox "�\�肳��Ă����i���̒l���s���ł��B"
        Exit Function
    End If
    
    '�\��J�n�Ɨ\��I���Z���ɓ��͂��Ȃ���΁A�폜����
    If planedRate = 0 And isExistSSL Then
        lineName = lineCl.GetShapeName(lineRowNo, mode)
        Set lineObj = lineCl.GetLineObj(lineName)
        Call lineCl.deleteLine(lineObj)
        Exit Function
    ElseIf planedRate = 0 Then
        Exit Function
    End If
    
    '�J�����_�[�C���X�^���X�̐���
    Dim calCl As New CalendarClass
    Set calCl = New CalendarClass
    
    
    '���C���̍��[��ݒ�
    Dim leftCol As Integer
    Set tmpRange = calCl.GetDateRngByDate(startTerm.Value)
    leftCol = tmpRange.Column
    '���C���̉E�[��ݒ�
    Dim rightCol As Integer
    Set tmpRange = calCl.GetDateRngByDate(endTerm.Value)
    rightCol = tmpRange.Column


    
    Dim settingRange As Range
    Set settingRange = ActiveSheet.Range(Cells(lineRowNo, leftCol).Address, Cells(lineRowNo, rightCol).Address)
    
    If isExistSSL Then
        lineName = lineCl.GetShapeName(lineRowNo, mode)
        Set lineObj = lineCl.GetLineObj(lineName)
        Call lineCl.ModifyLine(lineObj, settingRange, mode)
        Call lineCl.SetLineWidth(lineObj, planedRate, settingRange)
    Else
        Call lineCl.MakeLineByRange(settingRange, mode)
        lineName = lineCl.GetShapeName(lineRowNo, mode)
        Set lineObj = lineCl.GetLineObj(lineName)
        Call lineCl.SetLineWidth(lineObj, planedRate, settingRange)
    End If
    
End Function


'*****************************************************************
'�J�����_�[�̐ݒ�𔽉f����
'*****************************************************************
Sub SettingCalendar()
    '�J�����_�[�C���X�^���X�𐶐�
    Dim calCl As CalendarClass
    Set calCl = New CalendarClass
    
    '�J�n�N�������擾
    Dim startYmd As Date
    startYmd = ActiveSheet.Cells(CALENDAR_CONF_START_ROW, CALENDAR_CONF_START_COLUMN)
    
    '�I���N�������擾
    Dim endYmd As Date
    endYmd = ActiveSheet.Cells(CALENDAR_CONF_END_ROW, CALENDAR_CONF_END_COLUMN)
    '�J�����_�[����U�폜
    Call calCl.DeleteCalendar
    '���͒l�Ɋ�Â��č쐬
    Call calCl.SettingCalendar(startYmd, endYmd)
    
End Sub




