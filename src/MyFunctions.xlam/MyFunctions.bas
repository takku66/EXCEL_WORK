Attribute VB_Name = "MyFunctions"
Option Explicit

Dim MF_AS_MNG As MF_AUTOSHAPE_MANAGER

'=============================================
' �֐��Q

' �s���ōŌ�ɒl������񐔂�Ԃ�
Public Function MF_COLUMN_TAILINROW(cell As Range)
    MF_COLUMN_TAILINROW = Worksheets(cell.Parent.name).Cells(cell.Row, Columns.Count).End(xlToLeft).Column
End Function

' ����ōŌ�ɒl������s����Ԃ�
Public Function MF_ROW_TAILINCOLUMN(cell As Range)
    MF_ROW_TAILINCOLUMN = Worksheets(cell.Parent.name).Cells(Rows.Count, cell.Column).End(xlUp).Row
End Function

' �s���ōŏ��ɒl������񐔂�Ԃ�
Public Function MF_COLUMN_HEADINROW(cell As Range)
    MF_COLUMN_HEADINROW = Worksheets(cell.Parent.name).Cells(cell.Row, 1).End(xlToRight).Column
End Function

' ����ōŏ��ɒl������s����Ԃ�
Public Function MF_ROW_HEADINCOLUMN(cell As Range)
    MF_ROW_HEADINCOLUMN = Worksheets(cell.Parent.name).Cells(1, cell.Column).End(xlDown).Row
End Function

' ����������ŕ������A�w��̈ʒu�̒l��Ԃ�
Public Function MF_VALUE_FROMSPLIT(cell As Range, splitStr As String, idx As Long)
    Dim ary() As String
    ary = split(cell.Value, splitStr)
    MF_VALUE_FROMSPLIT = ary(idx)
End Function
'=============================================


'=============================================
' �}�`�Ǘ��n

' �}�`�Ǘ��V�[�g���쐬����
Public Sub CREATE_AUTOSHAPE_MANAGER()
    INIT_AUTOSHAPE_MANAGER
    If Not (MF_AS_MNG.IS_EXIST_CONTROL_SHEET) Then
        MF_AS_MNG.CREATE_TEMPLATE_SHEET
    End If
    MF_AS_MNG.ACTIVATE
End Sub

' �}�`�Ǘ��I�u�W�F�N�g������������
Private Function INIT_AUTOSHAPE_MANAGER()
    If MF_AS_MNG Is Nothing Then
        Set MF_AS_MNG = New MF_AUTOSHAPE_MANAGER
    End If
    Set INIT_AUTOSHAPE_MANAGER = MF_AS_MNG
End Function

' �}�`���̈ꗗ���X�V����
Public Sub UPDATE_CONTROL_LIST()
    ' �r���Ń}�N�������s�����ꍇ�́A�Ǘ��p�I�u�W�F�N�g��������
    ' �b��Ή��Ƃ��āA�ی��̏���������
    INIT_AUTOSHAPE_MANAGER
    MF_AS_MNG.UPDATE_CONTROL_LIST
End Sub

' �ꗗ�ɂ���}�`�����A�����̐}�`�ɔ��f������
' �}�`���Ȃ���΁A�쐬�����
Public Sub REFLECT_AUTOSHAPE()
    ' �r���Ń}�N�������s�����ꍇ�́A�Ǘ��p�I�u�W�F�N�g��������
    ' �b��Ή��Ƃ��āA�ی��̏���������
    INIT_AUTOSHAPE_MANAGER
    MF_AS_MNG.REFLECT_AUTOSHAPE
End Sub


'=============================================



'=============================================
' ����

Public Sub SHOW_MF_HELP()
    MF_HELP.Show
End Sub

'=============================================


