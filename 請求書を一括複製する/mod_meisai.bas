Attribute VB_Name = "mod_meisai"
'--------------------------------------------------
' m o d _ m e i s a i
'
' ���׈ꗗ�̌ڋq�f�[�^���ɁA���ו\�𕡐�����B
'
' 2011.03.10 
'
Option Explicit

Private Const START_DATA_ROW As Integer = 13 '�f�[�^�J�n�s

'�^�C�g���A����
Private Const SHEET_NAME_SETUP As String = "�ݒ�" '�ݒ�V�[�g��
Private Const SHEET_NAME_TEMPLATE As String = "�e���v���[�g" '�e���v���[�g�V�[�g��
Private Const SHEET_NAME_PREFIX As String = "�x���˗���" '���������V�[�g�̖��O

'���b�Z�[�W
Private Const MSG_1 As String = "�Â��x���˗������S�č폜����܂����A��낵���ł����H"
Private Const MSG_2 As String = "����"
Private Const MSG_3 As String = "�f�[�^�̓��͂�����܂���B"

'���v���C�X�Ώە���
Private Const REPLACE_WORD As String = "[month]"


' ���ׂ𕡐�����
Public Sub createMeisai()
    Dim sheetName As String
    Dim i As Integer
    Dim intEndRow As Integer
    
    '�f�[�^�s�̗L���̊m�F
    intEndRow = Worksheets(SHEET_NAME_SETUP).Cells(Rows.Count, 2).End(xlUp).Row
    If (intEndRow < START_DATA_ROW) Then
        Call MsgBox(MSG_3, vbExclamation)
        Exit Sub
    End If
    
    ' ���׍폜�����b�Z�[�W
    If (Worksheets.Count > 2) Then
        If (MsgBox(MSG_1, vbOKCancel + vbQuestion) = vbCancel) Then
            Exit Sub
        Else
            Call deleteSheets
        End If
    End If
    
    ' ��ʂ�����h�~
    Application.ScreenUpdating = False
    
    
    For i = START_DATA_ROW To intEndRow
        '�V�[�g�i�e���v���[�g�j�̃R�s�[
        Sheets(SHEET_NAME_TEMPLATE).Copy After:=Worksheets(Worksheets.Count)
        sheetName = SHEET_NAME_PREFIX & " (" & Worksheets.Count - 2 & ")"
        Worksheets(Worksheets.Count).Name = sheetName '�V�[�g���ύX
        Worksheets(Worksheets.Count).Unprotect '�V�[�g�ی�̉���
        
        '���ד��e�̃Z�b�g
        Call setMeisai(i, sheetName)
    Next
    
    Sheets(SHEET_NAME_SETUP).Activate
    
    ' ��ʂ�����h�~����
    Application.ScreenUpdating = True

End Sub

'�e���v���[�g���R�s�[�����V�[�g�ɕK�v���𖄂߂�
Private Sub setMeisai(index As Integer, sheetName As String)
    Dim intRowNum As Integer
    Dim bufString As String
    
    '�f�[�^�̊Y������s
    intRowNum = index '+ START_DATA_ROW - 1
    

    With Worksheets(SHEET_NAME_SETUP)
        '�\����
        Worksheets(sheetName).Cells(6, 3).Value = .Cells(8, 2).Value
        Worksheets(sheetName).Cells(6, 8).Value = .Cells(8, 2).Value
        
        '�x����
        Worksheets(sheetName).Cells(13, 15).Value = .Cells(intRowNum, 2).Value
        
        '���z
        Worksheets(sheetName).Cells(13, 24).Value = .Cells(intRowNum, 4).Value
        
        '�������e
        Worksheets(sheetName).Cells(17, 7).Value = .Cells(intRowNum, 3).Value
        
        
        
    End With
    
End Sub

' �ݒ�ƃe���v���[�g�V�[�g���������s�v�ȃV�[�g���ꊇ�폜����
Private Sub deleteSheets()
    Dim i As Integer
    Dim sheetName As String
    For i = Worksheets.Count To 1 Step -1
        sheetName = Worksheets(i).Name
        If (sheetName <> SHEET_NAME_SETUP And sheetName <> SHEET_NAME_TEMPLATE) Then
            Application.DisplayAlerts = False
            Sheets(sheetName).Delete
            Application.DisplayAlerts = True
        End If
    Next
End Sub

