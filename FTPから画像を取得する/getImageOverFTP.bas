Attribute VB_Name = "getImagesOverFTP"
Option Explicit

' �Q�l BASP21 �̗��p
' http://officetanaka.net/excel/vba/tips/tips47.htm

' �萔
Private Const glbImageSubFolder = "�images" ' �摜�t�@�C�����i�[����T�u�t�H���_��
Private Const glbFTP_IP As String = "192.168.1.1" '  FTP�T�[�oIP
Private Const glbFTP_User As String = "xxxxx" ' FTP UserID
Private Const glbFTP_Pass As String = "xxxxx" ' FTP Pass
Private Const glbIMG_PostFix As String = "_1.jpg" '

' EXCEL���[�N�V�[�g���ݒ�
Private Const glbSheetName As String = "CFOS�摜�擾" ' �V�[�g��
Private Const glbSKU_col As Integer = 2 ' JAN�R�[�h�̃J�����ԍ�
Private Const glbSKU_start_row As Integer = 3 ' JAN�R�[�h�̃J�����ԍ�

' ���b�Z�[�W
Private Const glbMsg01 As String = ""
Private Const glbMsg02 As String = ""
Private Const glbMsg03 As String = "FTP�ɐڑ��ł��܂���ł����B"


' ���C���֐�
Sub mainGetImages()

    ' �摜�擾�p�̃t�H���_���쐬
    Call createImageSubFolder
    
    ' �f�[�^�̃N���A
    Call clearData
    
    ' �摜�t�@�C���̎擾
    Call getFilesOverFTP
End Sub

' --------------------------------------------------
' �摜�擾�p�̃t�H���_�𐶐�����
' --------------------------------------------------
Sub createImageSubFolder()
    Dim objFSO As Object
    Dim myFolder As String
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    myFolder = ThisWorkbook.Path & glbImageSubFolder
    
    If objFSO.folderexists(folderspec:=myFolder) = False Then
        ' �t�H���_�����݂��Ȃ��ꍇ�A�t�H���_�𐶐�����
        objFSO.createfolder myFolder
    End If
    Set objFSO = Nothing
End Sub

' --------------------------------------------------
' �����O�ɕs�v�f�[�^���N���A���� �iOK/NG�t���O�̃N���A�j
' --------------------------------------------------
Sub clearData()
    Dim intEndRow
    
    ' �Y���̃��[�N�V�[�g��Activate����
    Worksheets(glbSheetName).Activate
    
    '�f�[�^�̖��s���擾
    intEndRow = Worksheets(glbSheetName).Cells(Rows.Count, glbSKU_col + 1).End(xlUp).Row
    
    ' �f�[�^���N���A
    Worksheets(glbSheetName).Range(Cells(glbSKU_start_row, glbSKU_col + 1), Cells(intEndRow, glbSKU_col + 1)).Clear

End Sub

' --------------------------------------------------
' �摜�t�@�C�����擾����
' --------------------------------------------------
Sub getFilesOverFTP()
    Dim FTP, rc As Long, Server As String, User As String, Pass As String
    Dim strFolder As String
    Dim intEndRow As Integer
    Dim strJAN As String
    Dim i As Integer
    
    'FTP�I�u�W�F�N�g ���s����Basp21���K�v
    Set FTP = CreateObject("basp21.FTP")
    
    '�ۑ���t�H���_�̎w��
    strFolder = ThisWorkbook.Path & glbImageSubFolder
    
    'FTP�ڑ�
    rc = FTP.Connect(glbFTP_IP, glbFTP_User, glbFTP_Pass)
    If rc <> 0 Then
        '�ڑ��G���[
        Call MsgBox(glbMsg03, vbCritical)
        FTP.Close
        Exit Sub
    End If
    
    
    '�f�[�^�̖��s���擾�iJAN�R�[�h�̍s�j
    intEndRow = Worksheets(glbSheetName).Cells(Rows.Count, glbSKU_col).End(xlUp).Row
    
    '�f�[�^���擾
    If (intEndRow >= glbSKU_start_row) Then
        For i = glbSKU_start_row To intEndRow
            strJAN = Trim(Worksheets(glbSheetName).Cells(i, glbSKU_col).Value)
            If (strJAN <> "") Then
                'FTP�擾
                rc = FTP.GetFile(strJAN & glbIMG_PostFix, strFolder)
                If rc <> 1 Then
                    ' �X�e�[�^�X�X�V�i�t�@�C���擾NG�j
                    Worksheets(glbSheetName).Cells(i, glbSKU_col + 1).Value = "NG"
                    Worksheets(glbSheetName).Cells(i, glbSKU_col + 1).Font.ColorIndex = 3 ' 3=��
                Else
                    ' �X�e�[�^�X�X�V�i�t�@�C���擾OK�j
                    Worksheets(glbSheetName).Cells(i, glbSKU_col + 1).Value = "OK"
                End If
            Else
                ' �X�e�[�^�X�X�V�iJAN���w��j
                Worksheets(glbSheetName).Cells(i, glbSKU_col + 1).Value = "-"
            End If
        Next i
    End If
    
    'FTP�ڑ� �N���[�Y
    FTP.Close
End Sub
