VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'--------------------------------------------------
'
'
'
'
'--------------------------------------------------
Option Explicit

Private Const START_ROW As Integer = 10 ' ���o���������̏����o�����J�n����s
Private Const BUTTON_EXTRACT As String = "cmdButton1"
Private Const BUTTON_CLEAR As String = "cmdButton2"
Private Const CHECKBOX_SORT As String = "cbSort"



' �폜�{�^��
Private Sub cmdButton2_Click()
    Call deleteAllShapes
End Sub

' ���o�{�^��
Private Sub cmdButton1_Click()
    Call extractText
End Sub

' �}�`����e�L�X�g�𒊏o����
Private Sub extractText()
    Dim j As Integer
    Dim buf As Integer ' DEBUG�p
    Dim objShape As Shape
    
    
    '���[�̃e�L�X�g�\���s�̃N���A�[
    Range("A:A").ClearContents
    
    
    ' �O���[�v������
    Do While True
        If (ungroupObj() = 0) Then
            Exit Do
        End If
    Loop
    
    
    '10�s�ڂ��珑���o��
    j = START_ROW
    
    '�}�`����e�L�X�g�𒊏o
    For Each objShape In Shapes
        
        '�f�o�b�O�p�Fbuf = Shapes(i).Type
        If (objShape.Type = msoAutoShape Or objShape.Type = msoTextBox) Then
            If objShape.TextEffect.Text <> "" Then
                Cells(j, 1).Value = objShape.TextEffect.Text
                j = j + 1
            End If
        End If
    Next objShape
    
    If cbSort = True Then
        ' �`�F�b�N�{�b�N�X�Ƀ`�F�b�N������ꍇ
        ' ���בւ����s��
        Call sortText
    End If
End Sub

' �O���[�v���� -- �O���[�v�����������������񐔂�Ԃ�
Private Function ungroupObj() As Integer
    Dim cntUnGroup As Integer '������
    Dim objShape As Shape
    
    cntUnGroup = 0
    For Each objShape In Shapes
      If (objShape.Type = msoGroup Or objShape.Type = msoPicture) Then
        ' �摜�t�@�C������쐬�����}�̂Ȃ��ɁA�O���[�v������
        ' �I�u�W�F�N�g���܂܂�邱�Ƃ�����B
            '--------------------------------------------------
            'msoAutoShape         1   �I�[�g�V�F�C�v
            'msoCallout           2   �����o��
            'msoChart             3   ���ߍ��݃O���t
            'msoComment           4   �R�����g
            'msoFreeform          5   �t���[�t�H�[��
            'msoGroup             6   �O���[�v�����ꂽ�}�`
            'msoEmbeddedOLEObject 7   ���ߍ���OLE�I�u�W�F�N�g
            'msoFormControl       8   �t�H�[���R���g���[��
            'msoLine              9   �����E���
            'msoLinkedOLEObject  10  �����NOLE�I�u�W�F�N�g
            'msoLinkedPicture    11  ���摜�Ƀ����N���Ă���}
            'msoOLEControlObject 12  ActiveX�R���g���[��
            'msoPicture          13  �摜�t�@�C������쐬�����}
            'msoPlaceholder      14  �iEXCEL�ł͎g�p���Ȃ��j
            'msoTextEffect       15  ���[�h�A�[�g
            'msoMedia            16  �iEXCEL�ł͎g�p���Ȃ��j
            'msoTextBox          17  �e�L�X�g�{�b�N�X
            'msoScriptAnchor     18  �X�N���v�g�A���J�[
            'msoTable            19  �e�[�u��
            '--------------------------------------------------
      
            On Error GoTo SkipCount
                      objShape.Ungroup
                      cntUnGroup = cntUnGroup + 1
SkipCount:
        On Error GoTo 0
      End If
    Next objShape
    
    ungroupObj = cntUnGroup
    
End Function


' �}�`���i�{�^���ȊO�j�S�č폜����
Private Sub deleteAllShapes()
    Dim objShape As Shape
    
    For Each objShape In Shapes
      If (objShape.Name <> BUTTON_EXTRACT And _
          objShape.Name <> BUTTON_CLEAR And _
          objShape.Name <> CHECKBOX_SORT) Then
          
          ' �}�`�̍폜
          objShape.Delete
      End If
    Next objShape
End Sub

' ���o�����e�L�X�g�̕��בւ����s��
Private Sub sortText()
    Dim endRow As Integer
    Dim startRow As Integer
    
    startRow = START_ROW
    endRow = Cells(Rows.count, 1).End(xlUp).Row
    
    If startRow < endRow Then
        '�\�[�g
        Range("A10", "A" & endRow).Sort _
            Key1:=Range("A10", "A" & endRow), _
            Order1:=xlAscending, _
            Header:=xlNo, _
            MatchCase:=False, _
            Orientation:=xlTopToBottom, _
            DataOption1:=xlSortNormal
    End If
    
End Sub

