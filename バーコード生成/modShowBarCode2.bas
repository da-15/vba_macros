Attribute VB_Name = "modShowBarCode2"
'*******************************************************************************
'   バーコードを表示 ※UA_BARCD.DLL必須
'
'   作成者:井上治  URL:http://www.ne.jp/asahi/excel/inoue/ [Excelでお仕事!]
'*******************************************************************************
Option Explicit


Private Const START_ROW As Long = 4 ' JANコードを表示する行
Private Const JAN_COL As Long = 2 ' JANコードを表示する列


'*******************************************************************************
' バーコードを表示
'*******************************************************************************
Sub DispBarCode()
    Dim xlAPP As Application
    Dim lngCurrRow As Long ' バーコードを表示する行を指定
    Dim objRange As Range
    
    
    Set xlAPP = Application
    xlAPP.ScreenUpdating = False
    xlAPP.Cursor = xlWait
    xlAPP.Calculation = xlCalculationManual
    lngCurrRow = START_ROW
    ' A列にバーコード値がある限り繰り返す
    Do While Cells(lngCurrRow, JAN_COL).Value <> ""
        ' 結合セル等の場合はここで調整
        Cells(lngCurrRow, JAN_COL).Select
        Set objRange = Selection
        
        ' カラムの書式設定 -----
        ' objRange.RowHeight = JAN_ROW_HEIGHT
        ' objRange.Font.Size = JAN_FONT_SIZE
        ' objRange.VerticalAlignment = xlVAlignBottom
        ' objRange.HorizontalAlignment = xlHAlignCenter
        ' --------------
        
        ' ■バーコードを貼り付ける
        ' (高さの係数、幅、余白は適当に調整のこと,
        '  チェックデジットはモジュール側で再計算されます)
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
' バーコードを消去
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


