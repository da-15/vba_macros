Attribute VB_Name = "modShowBarCode2"
'*******************************************************************************
'   �o�[�R�[�h��\�� ��UA_BARCD.DLL�K�{
'
'   �쐬��:��㎡  URL:http://www.ne.jp/asahi/excel/inoue/ [Excel�ł��d��!]
'*******************************************************************************
Option Explicit


Private Const START_ROW As Long = 4 ' JAN�R�[�h��\������s
Private Const JAN_COL As Long = 2 ' JAN�R�[�h��\�������


'*******************************************************************************
' �o�[�R�[�h��\��
'*******************************************************************************
Sub DispBarCode()
    Dim xlAPP As Application
    Dim lngCurrRow As Long ' �o�[�R�[�h��\������s���w��
    Dim objRange As Range
    
    
    Set xlAPP = Application
    xlAPP.ScreenUpdating = False
    xlAPP.Cursor = xlWait
    xlAPP.Calculation = xlCalculationManual
    lngCurrRow = START_ROW
    ' A��Ƀo�[�R�[�h�l���������J��Ԃ�
    Do While Cells(lngCurrRow, JAN_COL).Value <> ""
        ' �����Z�����̏ꍇ�͂����Œ���
        Cells(lngCurrRow, JAN_COL).Select
        Set objRange = Selection
        
        ' �J�����̏����ݒ� -----
        ' objRange.RowHeight = JAN_ROW_HEIGHT
        ' objRange.Font.Size = JAN_FONT_SIZE
        ' objRange.VerticalAlignment = xlVAlignBottom
        ' objRange.HorizontalAlignment = xlHAlignCenter
        ' --------------
        
        ' ���o�[�R�[�h��\��t����
        ' (�����̌W���A���A�]���͓K���ɒ����̂���,
        '  �`�F�b�N�f�W�b�g�̓��W���[�����ōČv�Z����܂�)
        Call ShowBarCode("JAN", Left$(Cells(lngCurrRow, JAN_COL).Value, 12), _
            objRange, Cells(lngCurrRow, JAN_COL).height * 0.65, 1, 5, 2)
        
        lngCurrRow = lngCurrRow + 1
    Loop
    Cells(1, 1).Select
    xlAPP.Calculation = xlCalculationAutomatic
    xlAPP.Cursor = xlDefault
    xlAPP.ScreenUpdating = True
End Sub

'*******************************************************************************
' �o�[�R�[�h������
'*******************************************************************************
Sub EraseBarCode()
    Dim xlAPP As Application
    Dim objPicture As Object
    
    Set xlAPP = Application
    xlAPP.ScreenUpdating = False
    xlAPP.Cursor = xlWait
    xlAPP.Calculation = xlCalculationManual
    For Each objPicture In ActiveSheet.DrawingObjects
        If Left$(objPicture.Name, 7) = "Picture" Then
            objPicture.Delete
        End If
    Next objPicture
    xlAPP.Calculation = xlCalculationAutomatic
    xlAPP.Cursor = xlDefault
    xlAPP.ScreenUpdating = True
End Sub


