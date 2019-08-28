Attribute VB_Name = "modShowBarCode"
'*******************************************************************************
'   バーコードを表示 ※UA_BARCD.DLL必須
'
'   作成者:井上治  URL:http://www.ne.jp/asahi/excel/inoue/ [Excelでお仕事!]
'   (ShowBarCodeの引数にintTopを追加 2010.09.01)
'*******************************************************************************
Option Explicit

Private Const LOGPIXELSX = 88       ' ポイント→ピクセル変換指定(横)
' バーコードをクリップボードにコピー(要「UA_BARCD.DLL」)
Private Declare Function BarcodeCopy Lib "UA_BARCD.DLL" _
    (ByVal bar_buf As String, _
     ByVal bar_margin As Integer, ByVal height As Integer, _
     ByVal haba As Integer, ByVal hdc As Long) As Integer
' ウィンドゥハンドルを返す
Private Declare Function FindWindow Lib "USER32.dll" _
    Alias "FindWindowA" (ByVal lpClassName As Any, _
    ByVal lpWindowName As Any) As Long
' DeskTopWindow取得
Private Declare Function GetDesktopWindow Lib "USER32.dll" _
    () As Long
' デバイスコンテキスト取得
Private Declare Function GetDC Lib "USER32.dll" _
    (ByVal hWnd As Long) As Long
' デバイスコンテキスト解放
Private Declare Function ReleaseDC Lib "USER32.dll" _
    (ByVal hWnd As Long, ByVal hdc As Long) As Long
' ポイント→ピクセル変換係数取得API
Private Declare Function GetDeviceCaps Lib "GDI32.dll" _
    (ByVal hdc As Long, ByVal nIndex As Long) As Long

'*******************************************************************************
' バーコードの表示 (intTopを追加 2010.09.01)
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
    ' Heightはポイント→ピクセル変換係数を乗じる
    intHeight2 = intHeight * GetLogPixelsXY / 72
    Set xlAPP = Application
    ' Exelウィンドウハンドル値とhDCを得る
    hWnd = FindWindow("XLMAIN", xlAPP.Caption)
    lngDC = GetDC(hWnd)
    ' バーコードをクリップボードにコピー
    intRET = BarcodeCopy(strType & "-" & strCode, intMargin, _
        intHeight2, intPoint, lngDC)
    ReleaseDC hWnd, lngDC
    ' 現時点では不成功は脱出
    If intRET <> 0 Then Exit Sub
    ' クリップボードの内容をシートにコピー
    ActiveSheet.Paste
    ' 貼り付けオブジェクトを取得してTopとLeftを調整
    Set objPicture = Selection
    lngLeft = objRange.Left
    If objRange.Width > objPicture.Width Then
        lngLeft = lngLeft + (objRange.Width - objPicture.Width) / 2
    End If
    objPicture.Left = lngLeft
    objPicture.Top = objRange.Top + intTop
    
    
    
End Sub

'*******************************************************************************
' 画面精細度のポイント→ピクセル変換係数算出
'*******************************************************************************
Private Function GetLogPixelsXY() As Long
     Dim lnghwnd As Long
     Dim lngDC As Long

     lnghwnd = GetDesktopWindow()
     lngDC = GetDC(lnghwnd)
     GetLogPixelsXY = GetDeviceCaps(lngDC, LOGPIXELSX)
     ReleaseDC lnghwnd, lngDC
End Function

