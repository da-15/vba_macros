Attribute VB_Name = "modShowBarCode"
'*******************************************************************************
'   �o�[�R�[�h��\�� ��UA_BARCD.DLL�K�{
'
'   �쐬��:��㎡  URL:http://www.ne.jp/asahi/excel/inoue/ [Excel�ł��d��!]
'   (ShowBarCode�̈�����intTop��ǉ� 2010.09.01)
'*******************************************************************************
Option Explicit

Private Const LOGPIXELSX = 88       ' �|�C���g���s�N�Z���ϊ��w��(��)
' �o�[�R�[�h���N���b�v�{�[�h�ɃR�s�[(�v�uUA_BARCD.DLL�v)
Private Declare Function BarcodeCopy Lib "UA_BARCD.DLL" _
    (ByVal bar_buf As String, _
     ByVal bar_margin As Integer, ByVal height As Integer, _
     ByVal haba As Integer, ByVal hdc As Long) As Integer
' �E�B���h�D�n���h����Ԃ�
Private Declare Function FindWindow Lib "USER32.dll" _
    Alias "FindWindowA" (ByVal lpClassName As Any, _
    ByVal lpWindowName As Any) As Long
' DeskTopWindow�擾
Private Declare Function GetDesktopWindow Lib "USER32.dll" _
    () As Long
' �f�o�C�X�R���e�L�X�g�擾
Private Declare Function GetDC Lib "USER32.dll" _
    (ByVal hWnd As Long) As Long
' �f�o�C�X�R���e�L�X�g���
Private Declare Function ReleaseDC Lib "USER32.dll" _
    (ByVal hWnd As Long, ByVal hdc As Long) As Long
' �|�C���g���s�N�Z���ϊ��W���擾API
Private Declare Function GetDeviceCaps Lib "GDI32.dll" _
    (ByVal hdc As Long, ByVal nIndex As Long) As Long

'*******************************************************************************
' �o�[�R�[�h�̕\�� (intTop��ǉ� 2010.09.01)
'*******************************************************************************
Public Sub ShowBarCode(strType As String, _
                       strCode As String, _
                       objRange As Range, _
                       intHeight As Integer, _
                       intPoint As Integer, _
                       intMargin As Integer, _
                       intTop As Integer)
    Dim xlAPP As Application
    Dim hWnd As Long
    Dim lngDC As Long
    Dim intRET As Integer
    Dim intHeight2 As Integer
    Dim objPicture As Object
    Dim lngLeft As Long
    
    If intPoint = 0 Then intPoint = 1
    If intHeight = 0 Then intHeight = objRange.height
    ' Height�̓|�C���g���s�N�Z���ϊ��W�����悶��
    intHeight2 = intHeight * GetLogPixelsXY / 72
    Set xlAPP = Application
    ' Exel�E�B���h�E�n���h���l��hDC�𓾂�
    hWnd = FindWindow("XLMAIN", xlAPP.Caption)
    lngDC = GetDC(hWnd)
    ' �o�[�R�[�h���N���b�v�{�[�h�ɃR�s�[
    intRET = BarcodeCopy(strType & "-" & strCode, intMargin, _
        intHeight2, intPoint, lngDC)
    ReleaseDC hWnd, lngDC
    ' �����_�ł͕s�����͒E�o
    If intRET <> 0 Then Exit Sub
    ' �N���b�v�{�[�h�̓��e���V�[�g�ɃR�s�[
    ActiveSheet.Paste
    ' �\��t���I�u�W�F�N�g���擾����Top��Left�𒲐�
    Set objPicture = Selection
    lngLeft = objRange.Left
    If objRange.Width > objPicture.Width Then
        lngLeft = lngLeft + (objRange.Width - objPicture.Width) / 2
    End If
    objPicture.Left = lngLeft
    objPicture.Top = objRange.Top + intTop
    
    
    
End Sub

'*******************************************************************************
' ��ʐ��דx�̃|�C���g���s�N�Z���ϊ��W���Z�o
'*******************************************************************************
Private Function GetLogPixelsXY() As Long
     Dim lnghwnd As Long
     Dim lngDC As Long

     lnghwnd = GetDesktopWindow()
     lngDC = GetDC(lnghwnd)
     GetLogPixelsXY = GetDeviceCaps(lngDC, LOGPIXELSX)
     ReleaseDC lnghwnd, lngDC
End Function

