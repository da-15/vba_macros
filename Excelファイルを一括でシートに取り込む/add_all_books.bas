Attribute VB_Name = "add_all_books"
Option Explicit

Sub addAllReports()
    Dim sFile As String
    Dim sWB As Workbook, dWB As Workbook
    Dim dSheetCount As Long
    Dim i, j As Long
    Dim cntSheetCopied As Integer
    Dim sheetName As String
    
    Dim sourceDir As String
    Const DEFAULT_FILE_NAME As String = "AllReports.xlsx"
    
    
    ' �t�H���_���w�肳��Ă��邩�`�F�b�N
    If ActiveSheet.Cells(3, 2).Value = "" Then
        MsgBox "�܂Ƃ߂����t�@�C���̏ꏊ���w�肵�Ă��������B", vbExclamation
        Exit Sub
    Else
        sourceDir = ActiveSheet.Cells(3, 2).Value
    End If
    
    '��ʂ�����h�~
    Application.ScreenUpdating = False
    
    '�w�肵���t�H���_���ɂ���u�b�N�̃t�@�C�������擾
    sFile = Dir(sourceDir & "*.xls")
    
    '�t�H���_���Ƀu�b�N���Ȃ���ΏI��
    If sFile = "" Then Exit Sub
    
    '�W��p�u�b�N���쐬
    Set dWB = Workbooks.Add
    
    '�W��O�̃V�[�g�����擾�i��ō폜���邽�߁j
    dSheetCount = dWB.Worksheets.Count
    
    '�V�[�g���̃J�E���^������
    cntSheetCopied = 0
    
    Do
        '�R�s�[���̃u�b�N���J��
        Set sWB = Workbooks.Open(Filename:=sourceDir & sFile)
        
    '�Ώۂ̃u�b�N�Ɋ܂܂��V�[�g��S�ăR�s�[
    For j = sWB.Worksheets.Count To 1 Step -1
        sheetName = sWB.Worksheets(j).Name
        sWB.Worksheets(sheetName).Copy After:=dWB.Worksheets(cntSheetCopied + dSheetCount)
        
        '�R�s�[�����V�[�g�̃J�E���g
        cntSheetCopied = cntSheetCopied + 1
        
        
        '�V�[�g�������ɃR�s�[No��ǉ�
        'ActiveSheet.Name = ActiveSheet.Name & "_" & cntSheetCopied
    Next
                
        '�R�s�[���t�@�C�������
        sWB.Close SaveChanges:=False
        
        '���̃u�b�N�̃t�@�C�������擾
        sFile = Dir()
        
        ' ���|�[�g�t�@�C���Ɠ������������ꍇ�̓X�L�b�v
        If sFile = DEFAULT_FILE_NAME Then
            sFile = Dir()
        End If
    Loop While sFile <> ""
        
    
    '�W��p�u�b�N�쐬���ɂ������V�[�g���폜
    Application.DisplayAlerts = False
    For i = dSheetCount To 1 Step -1
        dWB.Worksheets(i).Delete
    Next i
    
    Application.DisplayAlerts = True
    
On Error Resume Next
    '�W��p�u�b�N��ۑ����ĕ���
    'dWB.SaveAs Filename:=sourceDir & DEFAULT_FILE_NAME
    'dWB.Close SaveChanges:=False
    
    ' ��ʂ�����h�~ �����܂�
    Application.ScreenUpdating = False
End Sub



