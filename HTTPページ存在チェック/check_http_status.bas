Attribute VB_Name = "check_http_status"
'--------------------------------------------------
' 指定したページが存在しているかチェックを行う。
'
' HTTPステータス 404が返った時にNGを返す。
' ※ オプションとして、ただしページが存在しない場合でも意図的に404ステータスを返さないページもあるため
' 　　明示的にページに記載されている文字を読み込みエラーページとして意図的に判定させる。
' 例）
'「エラーページ判定キーワード」に"お探しの商品はみつかりませんでした。" というキーワードを指定すると
' このワードを含むページが返された時にエラーページとしてNGを返します。
'
'
' 2014.10.27 
' v6
'
' 2015.03.23 
' v7 ･･･ JP追加
'
' 2016.10.27 
' v8 汎用的なツールとして改修
'
'--------------------------------------------------

'--------------------------------------------------
' 定数
'--------------------------------------------------
Private Const RANGE_STATUS_RESULT As String = "A6:A65536" 'ステータス結果カラムのRange（クリア時に指定）

Private Const COL_RESULT As String = "A"          ' 結果出力列
Private Const COL_URL As String = "B"             ' URL出力列
Private Const START_DATA_ROW As Integer = 6        'データ開始行


Private Const ERROR_PAGE_KEYWORD As String = "B3" '添付文書のご用意がございません。"  'エラーページ判断ワード1
Private Const MSG_1 As String = "OK"        '結果カラム表示用
Private Const MSG_2 As String = "NG"        '結果カラム表示用
Private Const MSG_3 As String = "URL不正"   '結果カラム表示用


Sub checkPageStatus(strSheetName As String)

    Dim i As Long
    Dim bottom As Long
    
    'ステータス結果カラムのクリア
    Worksheets(strSheetName).Range(RANGE_STATUS_RESULT).ClearContents
    
    
    '最終行の取得
    bottom = Worksheets(strSheetName).Range(COL_URL + "65536").End(xlUp).Row
    For i = START_DATA_ROW To bottom
        'ステータスチェック
        Worksheets(strSheetName).Range(COL_RESULT & i).Value = GetWebStatus(Worksheets(strSheetName).Range(COL_URL & i).Value, Worksheets(strSheetName).Range(ERROR_PAGE_KEYWORD).Value)
        
    Next i
    
  
End Sub

Function GetWebStatus(strURL As String, strErrorKeyword As String) As String
    
    Dim objWinHttp As Object
    
    'Set objWinHttp = CreateObject("MSXML2.XMLHTTP")
    Set objWinHttp = CreateObject("WinHttp.WinHttpRequest.5.1")

On Error GoTo INVALID
    
    If strURL = "" Then
        ' 空文字チェック
        GetWebStatus = ""

    Else
        
        objWinHttp.Open "GET", strURL, False
        objWinHttp.send
        
        
        If objWinHttp.Status <> "200" Then
            'HTTPステータス200以外が返った場合は全てNG
            GetWebStatus = MSG_2
        
              
        ElseIf isErrorPage(objWinHttp.ResponseText, strErrorKeyword) Then
            'エラーであっても200が返る場合がある
            'サイトコンテンツからエラーページ判定を行う
            'エラーページが表示されている NG
            GetWebStatus = MSG_2
        
        Else
            '商品詳細ページが表示されている OK
            GetWebStatus = MSG_1
        End If
        
        ' Waitをかける（上手く処理ができない場合に。。。）
        ' Application.Wait (Now() + TimeValue("00:00:01"))
    End If
    
    Set objWinHttp = Nothing
    Exit Function
  
INVALID:
    'エラー発生時
    GetWebStatus = MSG_3
    Set objWinHttp = Nothing
  
  
End Function


Function isErrorPage(strHttpResponse As String, strErrorKeyword As String) As Boolean
' 表示されたページがエラーページであるか判定する
' 楽天、本店の場合ステータスコード200であがっているのに
' エラーが返るケースをチェックするため
    
    
    If strErrorKeyword = "" Then
        isErrorPage = False
    ElseIf InStr(strHttpResponse, strErrorKeyword) > 0 Then
        isErrorPage = True
    Else
        isErrorPage = False
    End If
End Function

