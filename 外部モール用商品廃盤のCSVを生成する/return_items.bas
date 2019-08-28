Attribute VB_Name = "Module2"
'--------------------------------------------------
' 返品を上げるためのアップロードファイルを生成します。
'
' 2014.11.18
'--------------------------------------------------

'--------------------------------------------------
' 定数
'--------------------------------------------------
Private Const SHEET_NAME_SETUP_RETURN As String = "【返品】CSV生成" '設定シート名
Private Const START_DATA_ROW_RETURN As Integer = 3 'データ開始行
Private Const MSG_1_R As String = ""
Private Const MSG_2_R As String = "にファイルを出力しました。"
Private Const MSG_3_R As String = "データの入力がありません。"
Private Const MSG_4_R As String = ""


'#NetDepot 返品#
Private Const FILE_NAME_NETDEPOT_RETURN As String = "netdepot_return.csv"
Private Const HEADER_NETDEPOT_RETURN = ""
Private Const LINE_REPLACE_NETDEPOT_RETURN As String = "[ID],[DATE],0:00,返品,返品,返品,返品,返品,返品,返品,返品,返品,-,[JAN],[ITEM_NAME],[STOCK],0,0,,,,0,,0,,,,[MESSAGE]"


Sub createNetDepotCSVReturn()
    Dim strHeader As String
    Dim strData As String
    Dim strId As String
    Dim strDate As String
    
    strDate = Year(Date) & "/" & Month(Date) & "/" & Day(Date)
    strId = "H" & Year(Date) & Month(Date) & Day(Date)
    
    'IDと日付を置き換える
    strData = Replace(LINE_REPLACE_NETDEPOT_RETURN, "[ID]", strId)
    strData = Replace(strData, "[DATE]", strDate)
    
    Call createCSVReturnItems(FILE_NAME_NETDEPOT_RETURN, HEADER_NETDEPOT_RETURN, strData)
End Sub


' --------------------------------------------------
' CSV出力 メイン
'
' strFileName … 出力するファイル名
' strHeader … 出力するヘッダ行（ヘッダ行がない場合は空とする）
' strLineReplace … データ行のフォーマット [JAN][ITEM_NAME][STOCK][MESSAGE]の文字列が置き換わる
' --------------------------------------------------
Private Sub createCSVReturnItems(strFileName As String, strHeader As String, strLineReplace As String)
    Dim intEndRow As Integer
    Dim strFilePath As String
    Dim i As Integer
    Dim strLine As String 'CSV出力用
    Dim strJAN As String 'JAN
    Dim strItemName As String '商品名
    Dim strStock As String  '返品数
    Dim strMessage As String '倉庫へのメッセージ
    Dim IntFlNo As Integer 'ファイルオープン用
    
    '出力ファイル名
    strFilePath = ActiveWorkbook.Path & strFileName
    
    'データ行の有無の確認
    intEndRow = Worksheets(SHEET_NAME_SETUP_RETURN).Cells(Rows.Count, 2).End(xlUp).Row
    If (intEndRow < START_DATA_ROW_RETURN) Then
        Call MsgBox(MSG_3_R, vbExclamation)
        Exit Sub
    End If
    
    'ファイルオープン
    IntFlNo = FreeFile
    Open strFilePath For Output As #IntFlNo
    
    'ヘッダ行
    If (strHeader <> "") Then
        Print #IntFlNo, strHeader
    End If
    
    'ファイル出力
    For i = START_DATA_ROW_RETURN To intEndRow
        strLine = ""
        strJAN = ""
        strItemName = ""
        strStock = ""
        strMessage = ""
        
        'データ取得
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
    
    'ファイルクローズ
    Close #IntFlNo
    
    
    'ファイル出力後メッセージ
    Call MsgBox(strFilePath & vbCrLf & MSG_2_R, vbInformation)
End Sub
