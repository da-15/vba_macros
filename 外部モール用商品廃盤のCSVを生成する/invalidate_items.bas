Attribute VB_Name = "invalidate_items"
'--------------------------------------------------
' 廃番品を上げるためのアップロードファイルを生成します。
'
' 2013.05.27
'--------------------------------------------------

'--------------------------------------------------
' 定数
'--------------------------------------------------
Private Const SHEET_NAME_SETUP As String = "【廃番】CSV生成" '設定シート名
Private Const START_DATA_ROW As Integer = 3 'データ開始行(Janコードのリスト) 
Private Const MSG_1 As String = "Amazon在庫を0にしても在庫連動により復活する可能性があります。" & vbCrLf & "Crossmallの在庫連動をOFFにすることも忘れないようご注意ください！"
Private Const MSG_2 As String = "にファイルを出力しました。"
Private Const MSG_3 As String = "データの入力がありません。"
Private Const MSG_4 As String = "該当JANの商品マスタが以下のステータスに置き換わります。" & vbCrLf & vbCrLf & "非表示" & vbCrLf & "発注点 = 0" & vbCrLf & "売切（発注不可）フラグ = ON"



'--------------------------------------------------
' ファイル出力フォーマット設定
'--------------------------------------------------
'#ECS非表示#
Private Const FILE_NAME_ECS_STATUS As String = "status.csv"
Private Const HEADER_ECS_INACTIVATE = ""
Private Const LINE_REPLACE_ECS_INACTIVATE As String = "[JAN],2"

'#ECS画像UP#
Private Const FILE_NAME_ECS_IMAGES As String = "images.csv"
Private Const HEADER_ECS_IMAGES = ""
Private Const LINE_REPLACE_ECS_IMAGES As String = "[JAN]"

'#ECS非表示＆発注点#
Private Const FILE_NAME_ECS_THRESHOLD As String = "items_ecs.csv"
Private Const HEADER_ECS_THRESHOLD = "product_code,status,sellout_flg,delive_order_threshold"
Private Const LINE_REPLACE_ECS_THRESHOLD As String = "[JAN],2,1,0"

'#楽天 倉庫#
Private Const FILE_NAME_RAKUTEN As String = "item.csv"
Private Const HEADER_RAKUTEN = "コントロールカラム,商品管理番号(商品URL),倉庫指定"
Private Const LINE_REPLACE_RAKUTEN As String = "u,[JAN],1"

'#楽天 在庫無#
Private Const FILE_NAME_RAKUTEN_STOCK As String = "item.csv"
Private Const HEADER_RAKUTEN_STOCK = "コントロールカラム,商品管理番号（商品URL）,倉庫指定,在庫タイプ,在庫数,在庫数表示,在庫戻しフラグ,在庫切れ時の注文受付"
Private Const LINE_REPLACE_RAKUTEN_STOCK As String = "u,[JAN],0,1,0,0,0,0"

'#楽天 削除#
Private Const FILE_NAME_RAKUTEN_DELETE As String = "item.csv"
Private Const HEADER_RAKUTEN_DELETE = "コントロールカラム,商品管理番号（商品URL）"
Private Const LINE_REPLACE_RAKUTEN_DELETE As String = "d,[JAN]"

'#Yahoo 在庫無#
Private Const FILE_NAME_YAHOO_STOCK As String = "yahoo_stock.csv"
Private Const HEADER_YAHOO_STOCK = "code,sub-code,quantity,mode"
Private Const LINE_REPLACE_YAHOO_STOCK As String = "[JAN],,0,"

'#Yahoo 削除#
Private Const FILE_NAME_YAHOO_DELETE As String = "yahoo_delete.csv"
Private Const HEADER_YAHOO_DELETE = "path,name,code,price"
Private Const LINE_REPLACE_YAHOO_DELETE As String = "a,a,[JAN],1"

'#Amazon 削除#
Private Const FILE_NAME_AMAZON_DELETE As String = "amazon_delete.txt"
Private Const HEADER_AMAZON_DELETE = "TemplateType=Health[tab]Version=2012.1130[tab]この行はAmazonが使用しますので変更や削除しないでください。" & _
    "[tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab]" & _
    "[tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab]" & _
    "[tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab]" & _
    "[tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab]" & _
    "[tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab]" & _
    "[tab][tab][tab][tab][tab][tab]" & vbCrLf
Private Const HEADER_AMAZON_DELETE2 = "商品管理番号[tab]商品名[tab]商品コード(JANコード等)[tab]商品コードのタイプ[tab]ブランド名[tab]" & _
    "メーカー名[tab]商品タイプ[tab]パッケージ商品数[tab]メーカー型番[tab]商品説明の箇条書き1[tab]商品説明の箇条書き2[tab]" & _
    "商品説明の箇条書き3[tab]商品説明の箇条書き4[tab]商品説明の箇条書き5[tab]商品説明文[tab]推奨ブラウズノード1[tab]" & _
    "検索キーワード1[tab]検索キーワード2[tab]検索キーワード3[tab]検索キーワード4[tab]検索キーワード5[tab]商品メイン画像URL[tab]" & _
    "在庫数[tab]リードタイム(出荷までにかかる作業日数)[tab]商品の販売価格[tab]通貨コード[tab]商品のコンディション[tab]" & _
    "商品のコンディション説明[tab]出品者カタログ番号[tab]原材料・成分1[tab]原材料・成分2[tab]原材料・成分3[tab]特別成分[tab]" & _
    "使用上の注意[tab]商品の利用(調理)方法[tab]警告[tab]法規上の免責条項[tab]アダルト商品[tab]スタイル名[tab]" & _
    "推奨ブラウズノード2[tab]親子関係の指定[tab]親商品のSKU(商品管理番号)[tab]親子関係のタイプ[tab]バリエーションテーマ[tab]" & _
    "フレーバー[tab]サイズ[tab]カラー[tab]カラーマップ[tab]香り[tab]商品の形状[tab]特殊機能1[tab]特殊機能2[tab]特殊機能3[tab]" & _
    "特定用途キーワード1[tab]特定用途キーワード2[tab]対象[tab]カラーサンプル画像URL[tab]商品のサブ画像URL1[tab]商品のサブ画像URL2[tab]" & _
    "商品のサブ画像URL3[tab]商品のサブ画像URL4[tab]商品のサブ画像URL5[tab]商品のサブ画像URL6[tab]商品のサブ画像URL7[tab]" & _
    "商品のサブ画像URL8[tab]推奨最小重量[tab]推奨最大重量[tab]推奨重量の単位[tab]商品の公開日[tab]商品の重量の単位[tab]" & _
    "商品の重量[tab]商品の長さの単位[tab]商品の長さ[tab]商品の幅[tab]商品の高さ[tab]配送重量の単位[tab]配送重量[tab]" & _
    "予約商品の販売開始日[tab]メーカー希望小売価格[tab]使用しない支払い方法[tab]配送日時指定SKUリスト[tab]セール価格[tab]" & _
    "フルフィルメントセンターID[tab]セール開始日[tab]セール終了日[tab]最大注文個数[tab]商品の入荷予定日[tab]最大同梱可能個数[tab]" & _
    "ギフトメッセージ[tab]ギフト包装[tab]メーカー製造中止[tab]商品コードなしの理由[tab]プラチナキーワード1[tab]プラチナキーワード2[tab]" & _
    "プラチナキーワード3[tab]プラチナキーワード4[tab]プラチナキーワード5[tab]アップデート・削除[tab][tab][tab][tab]" & _
    "[tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab]" & vbCrLf
Private Const HEADER_AMAZON_DELETE3 = "sku[tab]title[tab]standard-product-id[tab]product-id-type[tab]brand[tab]manufacturer[tab]" & _
    "product_type[tab]item-package-quantity[tab]mfr-part-number[tab]bullet-point1[tab]bullet-point2[tab]bullet-point3[tab]" & _
    "bullet-point4[tab]bullet-point5[tab]description[tab]recommended-browse-node1[tab]search-terms1[tab]search-terms2[tab]" & _
    "search-terms3[tab]search-terms4[tab]search-terms5[tab]main-image-url[tab]quantity[tab]leadtime-to-ship[tab]" & _
    "item-price[tab]currency[tab]condition-type[tab]condition-note[tab]merchant-catalog-number[tab]ingredients1[tab]" & _
    "ingredients2[tab]ingredients3[tab]special-ingredients[tab]indications[tab]directions[tab]warnings[tab]" & _
    "legal-disclaimer[tab]is-adult-product[tab]style-name[tab]recommended-browse-node2[tab]parentage[tab]parent-sku[tab]" & _
    "relationship-type[tab]variation-theme[tab]flavor[tab]size[tab]color[tab]color-map[tab]scent[tab]item-form[tab]" & _
    "special-features1[tab]special-features2[tab]special-features3[tab]specific-uses-keywords1[tab]specific-uses-keywords2[tab]" & _
    "target-audience[tab]swatch-image-url[tab]other-image-url1[tab]other-image-url2[tab]other-image-url3[tab]" & _
    "other-image-url4[tab]other-image-url5[tab]other-image-url6[tab]other-image-url7[tab]other-image-url8[tab]" & _
    "minimum-weight-recommendation[tab]maximum-weight-recommendation[tab]weight-recommendation-unit-of-measure[tab]" & _
    "launch-date[tab]item-weight-unit-of-measure[tab]item-weight[tab]item-length-unit-of-measure[tab]item-length[tab]" & _
    "item-width[tab]item-height[tab]shipping-weight-unit-of-measure[tab]shipping-weight[tab]release-date[tab]msrp[tab]" & _
    "optional-payment-type-exclusion[tab]delivery-schedule-group-id[tab]sale-price[tab]fulfillment-center-id[tab]" & _
    "sale-from-date[tab]sale-through-date[tab]max-order-quantity[tab]restock-date[tab]max-aggregate-ship-quantity[tab]" & _
    "is-gift-message-available[tab]is-giftwrap-available[tab]is-discontinued-by-manufacturer[tab]registered-parameter[tab]" & _
    "platinum-keywords1[tab]platinum-keywords2[tab]platinum-keywords3[tab]platinum-keywords4[tab]platinum-keywords5[tab]" & _
    "update-delete[tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab]"
Private Const LINE_REPLACE_AMAZON_DELETE As String = "[JAN][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab]" & _
    "[tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab]" & _
    "[tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab]" & _
    "[tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab]" & _
    "[tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab]delete[tab][tab][tab][tab][tab][tab][tab][tab][tab][tab]" & _
    "[tab][tab][tab][tab][tab][tab]"


'#Amazon 在庫0#
Private Const FILE_NAME_AMAZON_STOCK As String = "amazon_stock.txt"
Private Const HEADER_AMAZON_STOCK = "TemplateType=Health[tab]Version=1.7/1.2.11[tab]This row for Amazon.com use only.  " & _
    "Do not modify or delete.[tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab]" & _
    "[tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab]" & _
    "[tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab]" & _
    "[tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab]" & _
    "[tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab]" & vbCrLf
Private Const HEADER_AMAZON_STOCK2 = "sku[tab]title[tab]standard-product-id[tab]product-id-type[tab]brand[tab]manufacturer[tab]" & _
    "product_type[tab]mfr-part-number[tab]bullet-point1[tab]bullet-point2[tab]bullet-point3[tab]" & _
    "bullet-point4[tab]bullet-point5[tab]description[tab]recommended-browse-node1[tab]search-terms1[tab]" & _
    "search-terms2[tab]search-terms3[tab]search-terms4[tab]search-terms5[tab]main-image-url[tab]" & _
    "quantity[tab]leadtime-to-ship[tab]item-price[tab]currency[tab]merchant-catalog-number[tab]" & _
    "ingredients1[tab]ingredients2[tab]ingredients3[tab]indications[tab]directions[tab]warnings[tab]" & _
    "legal-disclaimer[tab]is-adult-product[tab]recommended-browse-node2[tab]parentage[tab]parent-sku[tab]" & _
    "relationship-type[tab]variation-theme[tab]flavor[tab]count[tab]size[tab]color[tab]scent[tab]" & _
    "other-image-url1[tab]other-image-url2[tab]other-image-url3[tab]other-image-url4[tab]" & _
    "other-image-url5[tab]other-image-url6[tab]other-image-url7[tab]other-image-url8[tab]launch-date[tab]" & _
    "item-weight-unit-of-measure[tab]item-weight[tab]item-length-unit-of-measure[tab]item-length[tab]" & _
    "item-width[tab]item-height[tab]shipping-weight-unit-of-measure[tab]shipping-weight[tab]" & _
    "release-date[tab]msrp[tab]sale-price[tab]fulfillment-center-id[tab]sale-from-date[tab]" & _
    "sale-through-date[tab]max-order-quantity[tab]restock-date[tab]max-aggregate-ship-quantity[tab]" & _
    "is-gift-message-available[tab]is-giftwrap-available[tab]is-discontinued-by-manufacturer[tab]" & _
    "registered-parameter[tab]platinum-keywords1[tab]platinum-keywords2[tab]platinum-keywords3[tab]" & _
    "platinum-keywords4[tab]platinum-keywords5[tab]update-delete"
Private Const LINE_REPLACE_AMAZON_STOCK As String = "[JAN][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab]" & _
    "[tab][tab][tab]0[tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab]" & _
    "[tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab]" & _
    "[tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab]" & _
    "[tab][tab][tab][tab]"

'--------------------------------------------------
' CSV出力関数の呼び出し
'--------------------------------------------------
'#ECS非表示#
Sub createEcsCSVInactivate()
    Call createCSV(FILE_NAME_ECS_STATUS, HEADER_ECS_INACTIVATE, LINE_REPLACE_ECS_INACTIVATE)
End Sub

'#ECS画像UP#
Sub createEcsCSVImages()
    Call createCSV(FILE_NAME_ECS_IMAGES, HEADER_ECS_IMAGES, LINE_REPLACE_ECS_IMAGES)
End Sub

'#ECS非表示＆発注点#
Sub createEcsCSVInactivateOrderThreshold()
    '注意勧告
    If (MsgBox(MSG_4, vbInformation Or vbOKCancel) = vbCancel) Then
        Exit Sub
    End If
    
    Call createCSV(FILE_NAME_ECS_THRESHOLD, HEADER_ECS_THRESHOLD, LINE_REPLACE_ECS_THRESHOLD)
End Sub


'#楽天#
Sub createRakutenCSV()
    Call createCSV(FILE_NAME_RAKUTEN, HEADER_RAKUTEN, LINE_REPLACE_RAKUTEN)
End Sub

'#楽天 在庫無#
Sub createRakutenCSVStock()
    Call createCSV(FILE_NAME_RAKUTEN_STOCK, HEADER_RAKUTEN_STOCK, LINE_REPLACE_RAKUTEN_STOCK)
End Sub

'#楽天 削除#
Sub createRakutenCSVDelete()
    Call createCSV(FILE_NAME_RAKUTEN_DELETE, HEADER_RAKUTEN_DELETE, LINE_REPLACE_RAKUTEN_DELETE)
End Sub

'#Yahoo! 在庫#
Sub createYahooCSVStock()
    Call createCSV(FILE_NAME_YAHOO_STOCK, HEADER_YAHOO_STOCK, LINE_REPLACE_YAHOO_STOCK)
End Sub

'#Yahoo! 削除#
Sub createYahooCSVDelete()
    Call createCSV(FILE_NAME_YAHOO_DELETE, HEADER_YAHOO_DELETE, LINE_REPLACE_YAHOO_DELETE)
End Sub

'#Amazon 削除#
Sub createAmazonCSVDelete()
    Dim strHeader As String
    Dim strData As String
    
    '[tab]をタブコードに置き換える
    strHeader = Replace(HEADER_AMAZON_DELETE, "[tab]", vbTab) & _
                Replace(HEADER_AMAZON_DELETE2, "[tab]", vbTab) & _
                Replace(HEADER_AMAZON_DELETE3, "[tab]", vbTab)
    strData = Replace(LINE_REPLACE_AMAZON_DELETE, "[tab]", vbTab)
    
    Call createCSV(FILE_NAME_AMAZON_DELETE, strHeader, strData)
End Sub

'#Amazon 在庫#
Sub createAmazonCSVStock()
    Dim strHeader As String
    Dim strData As String
    
    'Amazon在庫についての注意勧告
    If (MsgBox(MSG_1, vbInformation Or vbOKCancel) = vbCancel) Then
        Exit Sub
    End If
    
    '[tab]をタブコードに置き換える
    strHeader = Replace(HEADER_AMAZON_STOCK, "[tab]", vbTab) & _
                Replace(HEADER_AMAZON_STOCK2, "[tab]", vbTab)
    strData = Replace(LINE_REPLACE_AMAZON_STOCK, "[tab]", vbTab)
    
    Call createCSV(FILE_NAME_AMAZON_STOCK, strHeader, strData)
End Sub


' --------------------------------------------------
' CSV出力 メイン
'
' strFileName … 出力するファイル名
' strHeader … 出力するヘッダ行（ヘッダ行がない場合は空とする）
' strLineReplace … データ行のフォーマット [JAN]の文字列がJANコードに置き換わる
' --------------------------------------------------
Private Sub createCSV(strFileName As String, strHeader As String, strLineReplace As String)
    Dim intEndRow As Integer
    Dim strFilePath As String
    Dim i As Integer
    Dim strLine As String '出力行生成用
    Dim IntFlNo As Integer 'ファイルオープン用
    
    '出力ファイル名
    strFilePath = ActiveWorkbook.Path & strFileName
    
    'データ行の有無の確認
    intEndRow = Worksheets(SHEET_NAME_SETUP).Cells(Rows.Count, 2).End(xlUp).Row
    If (intEndRow < START_DATA_ROW) Then
        Call MsgBox(MSG_3, vbExclamation)
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
    For i = START_DATA_ROW To intEndRow
        strLine = ""
        strLine = Worksheets(SHEET_NAME_SETUP).Cells(i, 2).Value
        
        If (strLine <> "") Then
            strLine = Replace(strLineReplace, "[JAN]", strLine)
            Print #IntFlNo, strLine
        End If
    Next
    
    'ファイルクローズ
    Close #IntFlNo
    
    
    'ファイル出力後メッセージ
    Call MsgBox(strFilePath & vbCrLf & MSG_2, vbInformation)
End Sub

