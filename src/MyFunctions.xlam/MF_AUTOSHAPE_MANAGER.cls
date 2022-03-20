VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MF_AUTOSHAPE_MANAGER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***********************************
' �}�`���R���g���[������N���X
' < �{�N���X���񋟂���@�\ >
'   1. �u�b�N���ɑ��݂���}�`�̖��O����͒l�A�n�C�p�[�����N�l�Ȃǂ̈ꗗ�Ǘ�
'   2. �}�`�ɓ��͂���Ă���l�̌����E�u��
'   3. ���ɔz�u����Ă���}�`�̌����E�u��
'***********************************


' �}�`���Ǘ�����V�[�g�I�u�W�F�N�g
Private CONTROL_SHEET As Worksheet

' �}�`�������T�[�r�X�N���X
Private mf_shp_service As MF_SHAPE_SERVICE


' �e�Ǘ��@�\�ɕK�v�ȁA�e���̗�ʒu��A�ǂݎ��J�n�ʒu���`����
Const NAME_COL = 1
Const CONTENT_COL = 2
Const HYPERLINK_COL = 3
Const SUBHYPERLINK_COL = 4
Const TEMPLATE_AUTOSHAPE_COL = 5
Const READ_ROW = 4
Const TITLE_ROW = READ_ROW - 1

Const NAME_TITLE = "�}�`��"
Const CONTENT_TITLE = "���͒l"
Const HYPERLINK_TITLE = "���u�b�N�ւ̃n�C�p�[�����N"
Const SUBHYPERLINK_TITLE = "���u�b�N���̃n�C�p�[�����N"
Const TEMPLATE_AUTOSHAPE_TITLE = "���`�}�`"

' �f�t�H���g�̊Ǘ��V�[�g��
Const DEFAULT_CONTROL_SHEETNAME = "_ASCSHT_"



'******************************
' ����������
'
'******************************
Public Sub Class_Initialize()

    If IS_EXIST_CONTROL_SHEET Then
        Set CONTROL_SHEET = Worksheets(DEFAULT_CONTROL_SHEETNAME)
    Else
        CREATE_TEMPLATE_SHEET
    End If
    
    Set mf_shp_service = New MF_SHAPE_SERVICE
    
End Sub
Public Sub ACTIVATE()
    CONTROL_SHEET.ACTIVATE
End Sub
Public Function GET_SHEET_NAME()
    GET_SHEET_NAME = CONTROL_SHEET.name
End Function
Public Function IS_EXIST_CONTROL_SHEET()
    IS_EXIST_CONTROL_SHEET = False
    Dim ws As Worksheet
    For Each ws In Worksheets
        If ws.name = DEFAULT_CONTROL_SHEETNAME Then
            IS_EXIST_CONTROL_SHEET = True
            Exit Function
        End If
    Next ws
End Function




' �Ǘ��V�[�g�̃e���v���[�g���쐬����
Public Sub CREATE_TEMPLATE_SHEET()

    ' �V�[�g�̒ǉ�
    Set CONTROL_SHEET = ActiveWorkbook.Sheets.Add
    CONTROL_SHEET.name = DEFAULT_CONTROL_SHEETNAME
    
    ' �V�[�g���̈ꗗ�\���쐬
    CREATE_CONTROL_LIST
    
    ' �V�[�g���Ƀ{�^����z�u
    ' ���W�{�^��
    ' �{�^���̒u��
    CREATE_CONTROL_BUTTONS
End Sub

' �Ǘ��\���쐬
Private Sub CREATE_CONTROL_LIST()
    CONTROL_SHEET.Cells(TITLE_ROW, NAME_COL).Value = NAME_TITLE
    CONTROL_SHEET.Cells(TITLE_ROW, CONTENT_COL).Value = CONTENT_TITLE
    CONTROL_SHEET.Cells(TITLE_ROW, HYPERLINK_COL).Value = HYPERLINK_TITLE
    CONTROL_SHEET.Cells(TITLE_ROW, SUBHYPERLINK_COL).Value = SUBHYPERLINK_TITLE
    CONTROL_SHEET.Cells(TITLE_ROW, TEMPLATE_AUTOSHAPE_COL).Value = TEMPLATE_AUTOSHAPE_TITLE
    Columns(NAME_COL).ColumnWidth = 15
    Columns(CONTENT_COL).ColumnWidth = 30
    Columns(HYPERLINK_COL).ColumnWidth = 50
    Columns(SUBHYPERLINK_COL).ColumnWidth = 30
    Rows.RowHeight = 30
    
    Dim area As Range
    Dim titleArea As Range
    Set area = CONTROL_SHEET.Range(CONTROL_SHEET.Cells(TITLE_ROW, NAME_COL), CONTROL_SHEET.Cells(READ_ROW + 200, TEMPLATE_AUTOSHAPE_COL))
    Set titleArea = CONTROL_SHEET.Range(CONTROL_SHEET.Cells(TITLE_ROW, NAME_COL), CONTROL_SHEET.Cells(TITLE_ROW, TEMPLATE_AUTOSHAPE_COL))
    
    area.Borders(xlDiagonalDown).LineStyle = xlNone
    area.Borders(xlDiagonalUp).LineStyle = xlNone
    With area.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With area.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With area.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With area.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With area.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With area.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With titleArea.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    With titleArea
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    titleArea.Font.Bold = True
End Sub

' �Ǘ��p�̃{�^����z�u����
Private Sub CREATE_CONTROL_BUTTONS()
    CREATE_UPDATE_BUTTON
    CREATE_REPLACE_BUTTON
End Sub
Private Function CREATE_UPDATE_BUTTON()     ' �ꗗ���X�V�{�^��
    Dim updateBtn As Variant
    Set updateBtn = CONTROL_SHEET.Buttons.Add(0, 0, 80, 25)
    updateBtn.OnAction = "UPDATE_CONTROL_LIST"
    updateBtn.Characters.Text = "���X�V"
    With updateBtn.Characters(Start:=1, Length:=3).Font
        .name = "Meiryo UI"
        .FontStyle = "�W��"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 1
    End With
    Set CREATE_UPDATE_BUTTON = updateBtn
End Function
Private Function CREATE_REPLACE_BUTTON()     ' �쐬�E�u���{�^��
    Dim replaceBtn As Variant
    Set replaceBtn = CONTROL_SHEET.Buttons.Add(100, 0, 80, 25)
    replaceBtn.OnAction = "REFLECT_AUTOSHAPE"
    replaceBtn.Characters.Text = "�쐬�E�u��"
    With replaceBtn.Characters(Start:=1, Length:=3).Font
        .name = "Meiryo UI"
        .FontStyle = "�W��"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 1
    End With
    Set CREATE_REPLACE_BUTTON = replaceBtn
End Function



' �}�`�����W���āA�Ǘ��\�ɔ��f������
Public Sub UPDATE_CONTROL_LIST()
    COLLECT_AUTOSHAPE
End Sub
Public Sub COLLECT_AUTOSHAPE(Optional collectName As String)
    MsgBox "�}�`���W�����F������"
End Sub

' �Ǘ��\�ɐݒ肳��Ă���}�`�����X�V �܂��� �V�K�쐬����B
' �V�K�쐬���ꂽ�}�`�́A�Ǘ��\�ɏo�͂����
Public Sub REFLECT_AUTOSHAPE()
    
    Dim loopCnt As Long
    Dim maxCnt As Long
    maxCnt = MF_ROW_TAILINCOLUMN(CONTROL_SHEET.Range( _
                                CONTROL_SHEET.Cells(READ_ROW, NAME_COL), _
                                CONTROL_SHEET.Cells(READ_ROW, NAME_COL)))

    Dim name As String
    Dim content As String
    Dim hyperlink As String
    Dim subhyperlink As String
    Dim template As Shape
    
    ' �\����e�}�`�����擾���A
    ' ���O�ɍ��v����}�`���Ȃ���΍쐬�A
    ' ����΁A�����X�V����
    For loopCnt = READ_ROW To maxCnt
        name = CONTROL_SHEET.Cells(loopCnt, NAME_COL).Value
        content = CONTROL_SHEET.Cells(loopCnt, CONTENT_COL).Value
        hyperlink = CONTROL_SHEET.Cells(loopCnt, HYPERLINK_COL).Value
        subhyperlink = CONTROL_SHEET.Cells(loopCnt, SUBHYPERLINK_COL).Value
        Set template = SHAPE_INCELL(CONTROL_SHEET.Cells(loopCnt, TEMPLATE_AUTOSHAPE_COL))
        
        Call mf_shp_service.CREATE_MFSHAPE(CONTROL_SHEET, name, content, hyperlink, subhyperlink, template)
        
    Next loopCnt
End Sub

' �w��̃Z���Ɋ܂܂��}�`��Ԃ�
Public Function SHAPE_INCELL(cell As Range) As Shape
    Dim shp As Shape
    
    ' �S�Ă̐}�`���擾���A���̐}�`���w��̃Z���Ɋ܂܂�Ă��邩�m�F����
    ' �ŏ��̐}�`�݂̂�Ԃ�
    For Each shp In cell.Worksheet.Shapes
    
        Dim shpContainsCell As Range
        Set shpContainsCell = Range(shp.TopLeftCell, shp.BottomRightCell)
        
        If Not (Intersect(shpContainsCell, cell) Is Nothing) Then
            If Intersect(shpContainsCell, cell).Address = shpContainsCell.Address Then
                Set SHAPE_INCELL = shp
                Exit Function
            End If
        End If
        
    Next
End Function
