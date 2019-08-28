Attribute VB_Name = "check_http_status"
'--------------------------------------------------
' �w�肵���y�[�W�����݂��Ă��邩�`�F�b�N���s���B
'
' HTTP�X�e�[�^�X 404���Ԃ�������NG��Ԃ��B
' �� �I�v�V�����Ƃ��āA�������y�[�W�����݂��Ȃ��ꍇ�ł��Ӑ}�I��404�X�e�[�^�X��Ԃ��Ȃ��y�[�W�����邽��
' �@�@�����I�Ƀy�[�W�ɋL�ڂ���Ă��镶����ǂݍ��݃G���[�y�[�W�Ƃ��ĈӐ}�I�ɔ��肳����B
' ��j
'�u�G���[�y�[�W����L�[���[�h�v��"���T���̏��i�݂͂���܂���ł����B" �Ƃ����L�[���[�h���w�肷���
' ���̃��[�h���܂ރy�[�W���Ԃ��ꂽ���ɃG���[�y�[�W�Ƃ���NG��Ԃ��܂��B
'
'
' 2014.10.27 
' v6
'
' 2015.03.23 
' v7 ��� JP�ǉ�
'
' 2016.10.27 
' v8 �ėp�I�ȃc�[���Ƃ��ĉ��C
'
'--------------------------------------------------

'--------------------------------------------------
' �萔
'--------------------------------------------------
Private Const RANGE_STATUS_RESULT As String = "A6:A65536" '�X�e�[�^�X���ʃJ������Range�i�N���A���Ɏw��j

Private Const COL_RESULT As String = "A"          ' ���ʏo�͗�
Private Const COL_URL As String = "B"             ' URL�o�͗�
Private Const START_DATA_ROW As Integer = 6        '�f�[�^�J�n�s


Private Const ERROR_PAGE_KEYWORD As String = "B3" '�Y�t�����̂��p�ӂ��������܂���B"  '�G���[�y�[�W���f���[�h1
Private Const MSG_1 As String = "OK"        '���ʃJ�����\���p
Private Const MSG_2 As String = "NG"        '���ʃJ�����\���p
Private Const MSG_3 As String = "URL�s��"   '���ʃJ�����\���p


Sub checkPageStatus(strSheetName As String)

    Dim i As Long
    Dim bottom As Long
    
    '�X�e�[�^�X���ʃJ�����̃N���A
    Worksheets(strSheetName).Range(RANGE_STATUS_RESULT).ClearContents
    
    
    '�ŏI�s�̎擾
    bottom = Worksheets(strSheetName).Range(COL_URL + "65536").End(xlUp).Row
    For i = START_DATA_ROW To bottom
        '�X�e�[�^�X�`�F�b�N
        Worksheets(strSheetName).Range(COL_RESULT & i).Value = GetWebStatus(Worksheets(strSheetName).Range(COL_URL & i).Value, Worksheets(strSheetName).Range(ERROR_PAGE_KEYWORD).Value)
        
    Next i
    
  
End Sub

Function GetWebStatus(strURL As String, strErrorKeyword As String) As String
    
    Dim objWinHttp As Object
    
    'Set objWinHttp = CreateObject("MSXML2.XMLHTTP")
    Set objWinHttp = CreateObject("WinHttp.WinHttpRequest.5.1")

On Error GoTo INVALID
    
    If strURL = "" Then
        ' �󕶎��`�F�b�N
        GetWebStatus = ""

    Else
        
        objWinHttp.Open "GET", strURL, False
        objWinHttp.send
        
        
        If objWinHttp.Status <> "200" Then
            'HTTP�X�e�[�^�X200�ȊO���Ԃ����ꍇ�͑S��NG
            GetWebStatus = MSG_2
        
              
        ElseIf isErrorPage(objWinHttp.ResponseText, strErrorKeyword) Then
            '�G���[�ł����Ă�200���Ԃ�ꍇ������
            '�T�C�g�R���e���c����G���[�y�[�W������s��
            '�G���[�y�[�W���\������Ă��� NG
            GetWebStatus = MSG_2
        
        Else
            '���i�ڍ׃y�[�W���\������Ă��� OK
            GetWebStatus = MSG_1
        End If
        
        ' Wait��������i��肭�������ł��Ȃ��ꍇ�ɁB�B�B�j
        ' Application.Wait (Now() + TimeValue("00:00:01"))
    End If
    
    Set objWinHttp = Nothing
    Exit Function
  
INVALID:
    '�G���[������
    GetWebStatus = MSG_3
    Set objWinHttp = Nothing
  
  
End Function


Function isErrorPage(strHttpResponse As String, strErrorKeyword As String) As Boolean
' �\�����ꂽ�y�[�W���G���[�y�[�W�ł��邩���肷��
' �y�V�A�{�X�̏ꍇ�X�e�[�^�X�R�[�h200�ł������Ă���̂�
' �G���[���Ԃ�P�[�X���`�F�b�N���邽��
    
    
    If strErrorKeyword = "" Then
        isErrorPage = False
    ElseIf InStr(strHttpResponse, strErrorKeyword) > 0 Then
        isErrorPage = True
    Else
        isErrorPage = False
    End If
End Function

