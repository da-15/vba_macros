Attribute VB_Name = "Module2"
'--------------------------------------------------
' �ԕi���グ�邽�߂̃A�b�v���[�h�t�@�C���𐶐����܂��B
'
' 2014.11.18
'--------------------------------------------------

'--------------------------------------------------
' �萔
'--------------------------------------------------
Private Const SHEET_NAME_SETUP_RETURN As String = "�y�ԕi�zCSV����" '�ݒ�V�[�g��
Private Const START_DATA_ROW_RETURN As Integer = 3 '�f�[�^�J�n�s
Private Const MSG_1_R As String = ""
Private Const MSG_2_R As String = "�Ƀt�@�C�����o�͂��܂����B"
Private Const MSG_3_R As String = "�f�[�^�̓��͂�����܂���B"
Private Const MSG_4_R As String = ""


'#NetDepot �ԕi#
Private Const FILE_NAME_NETDEPOT_RETURN As String = "�netdepot_return.csv"
Private Const HEADER_NETDEPOT_RETURN = ""
Private Const LINE_REPLACE_NETDEPOT_RETURN As String = "[ID],[DATE],0:00,�ԕi,�ԕi,�ԕi,�ԕi,�ԕi,�ԕi,�ԕi,�ԕi,�ԕi,-,[JAN],[ITEM_NAME],[STOCK],0,0,,,,0,,0,,,,[MESSAGE]"


Sub createNetDepotCSVReturn()
    Dim strHeader As String
    Dim strData As String
    Dim strId As String
    Dim strDate As String
    
    strDate = Year(Date) & "/" & Month(Date) & "/" & Day(Date)
    strId = "H" & Year(Date) & Month(Date) & Day(Date)
    
    'ID�Ɠ��t��u��������
    strData = Replace(LINE_REPLACE_NETDEPOT_RETURN, "[ID]", strId)
    strData = Replace(strData, "[DATE]", strDate)
    
    Call createCSVReturnItems(FILE_NAME_NETDEPOT_RETURN, HEADER_NETDEPOT_RETURN, strData)
End Sub


' --------------------------------------------------
' CSV�o�� ���C��
'
' strFileName �c �o�͂���t�@�C����
' strHeader �c �o�͂���w�b�_�s�i�w�b�_�s���Ȃ��ꍇ�͋�Ƃ���j
' strLineReplace �c �f�[�^�s�̃t�H�[�}�b�g [JAN][ITEM_NAME][STOCK][MESSAGE]�̕����񂪒u�������
' --------------------------------------------------
Private Sub createCSVReturnItems(strFileName As String, strHeader As String, strLineReplace As String)
    Dim intEndRow As Integer
    Dim strFilePath As String
    Dim i As Integer
    Dim strLine As String 'CSV�o�͗p
    Dim strJAN As String 'JAN
    Dim strItemName As String '���i��
    Dim strStock As String  '�ԕi��
    Dim strMessage As String '�q�ɂւ̃��b�Z�[�W
    Dim IntFlNo As Integer '�t�@�C���I�[�v���p
    
    '�o�̓t�@�C����
    strFilePath = ActiveWorkbook.Path & strFileName
    
    '�f�[�^�s�̗L���̊m�F
    intEndRow = Worksheets(SHEET_NAME_SETUP_RETURN).Cells(Rows.Count, 2).End(xlUp).Row
    If (intEndRow < START_DATA_ROW_RETURN) Then
        Call MsgBox(MSG_3_R, vbExclamation)
        Exit Sub
    End If
    
    '�t�@�C���I�[�v��
    IntFlNo = FreeFile
    Open strFilePath For Output As #IntFlNo
    
    '�w�b�_�s
    If (strHeader <> "") Then
        Print #IntFlNo, strHeader
    End If
    
    '�t�@�C���o��
    For i = START_DATA_ROW_RETURN To intEndRow
        strLine = ""
        strJAN = ""
        strItemName = ""
        strStock = ""
        strMessage = ""
        
        '�f�[�^�擾
        strJAN = Worksheets(SHEET_NAME_SETUP_RETURN).Cells(i, 2).Value
        strItemName = Worksheets(SHEET_NAME_SETUP_RETURN).Cells(i, 3).Value
        strStock = Worksheets(SHEET_NAME_SETUP_RETURN).Cells(i, 4).Value
        strMessage = Worksheets(SHEET_NAME_SETUP_RETURN).Cells(i, 5).Value
        
        
        If (strJAN <> "") Then
            strLine = Replace(strLineReplace, "[JAN]", strJAN)
            strLine = Replace(strLine, "[ITEM_NAME]", strItemName)
            strLine = Replace(strLine, "[STOCK]", strStock)
            strLine = Replace(strLine, "[MESSAGE]", strMessage)
            Print #IntFlNo, strLine
        End If
    Next
    
    '�t�@�C���N���[�Y
    Close #IntFlNo
    
    
    '�t�@�C���o�͌チ�b�Z�[�W
    Call MsgBox(strFilePath & vbCrLf & MSG_2_R, vbInformation)
End Sub
