Attribute VB_Name = "getImagesOverFTP"
Option Explicit

' 参考 BASP21 の利用
' http://officetanaka.net/excel/vba/tips/tips47.htm

' 定数
Private Const glbImageSubFolder = "images" ' 画像ファイルを格納するサブフォルダ名
Private Const glbFTP_IP As String = "192.168.1.1" '  FTPサーバIP
Private Const glbFTP_User As String = "xxxxx" ' FTP UserID
Private Const glbFTP_Pass As String = "xxxxx" ' FTP Pass
Private Const glbIMG_PostFix As String = "_1.jpg" '

' EXCELワークシート内設定
Private Const glbSheetName As String = "CFOS画像取得" ' シート名
Private Const glbSKU_col As Integer = 2 ' JANコードのカラム番号
Private Const glbSKU_start_row As Integer = 3 ' JANコードのカラム番号

' メッセージ
Private Const glbMsg01 As String = ""
Private Const glbMsg02 As String = ""
Private Const glbMsg03 As String = "FTPに接続できませんでした。"


' メイン関数
Sub mainGetImages()

    ' 画像取得用のフォルダを作成
    Call createImageSubFolder
    
    ' データのクリア
    Call clearData
    
    ' 画像ファイルの取得
    Call getFilesOverFTP
End Sub

' --------------------------------------------------
' 画像取得用のフォルダを生成する
' --------------------------------------------------
Sub createImageSubFolder()
    Dim objFSO As Object
    Dim myFolder As String
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    myFolder = ThisWorkbook.Path & glbImageSubFolder
    
    If objFSO.folderexists(folderspec:=myFolder) = False Then
        ' フォルダが存在しない場合、フォルダを生成する
        objFSO.createfolder myFolder
    End If
    Set objFSO = Nothing
End Sub

' --------------------------------------------------
' 処理前に不要データをクリアする （OK/NGフラグのクリア）
' --------------------------------------------------
Sub clearData()
    Dim intEndRow
    
    ' 該当のワークシートをActivateする
    Worksheets(glbSheetName).Activate
    
    'データの末行を取得
    intEndRow = Worksheets(glbSheetName).Cells(Rows.Count, glbSKU_col + 1).End(xlUp).Row
    
    ' データをクリア
    Worksheets(glbSheetName).Range(Cells(glbSKU_start_row, glbSKU_col + 1), Cells(intEndRow, glbSKU_col + 1)).Clear

End Sub

' --------------------------------------------------
' 画像ファイルを取得する
' --------------------------------------------------
Sub getFilesOverFTP()
    Dim FTP, rc As Long, Server As String, User As String, Pass As String
    Dim strFolder As String
    Dim intEndRow As Integer
    Dim strJAN As String
    Dim i As Integer
    
    'FTPオブジェクト 実行環境にBasp21が必要
    Set FTP = CreateObject("basp21.FTP")
    
    '保存先フォルダの指定
    strFolder = ThisWorkbook.Path & glbImageSubFolder
    
    'FTP接続
    rc = FTP.Connect(glbFTP_IP, glbFTP_User, glbFTP_Pass)
    If rc <> 0 Then
        '接続エラー
        Call MsgBox(glbMsg03, vbCritical)
        FTP.Close
        Exit Sub
    End If
    
    
    'データの末行を取得（JANコードの行）
    intEndRow = Worksheets(glbSheetName).Cells(Rows.Count, glbSKU_col).End(xlUp).Row
    
    'データを取得
    If (intEndRow >= glbSKU_start_row) Then
        For i = glbSKU_start_row To intEndRow
            strJAN = Trim(Worksheets(glbSheetName).Cells(i, glbSKU_col).Value)
            If (strJAN <> "") Then
                'FTP取得
                rc = FTP.GetFile(strJAN & glbIMG_PostFix, strFolder)
                If rc <> 1 Then
                    ' ステータス更新（ファイル取得NG）
                    Worksheets(glbSheetName).Cells(i, glbSKU_col + 1).Value = "NG"
                    Worksheets(glbSheetName).Cells(i, glbSKU_col + 1).Font.ColorIndex = 3 ' 3=赤
                Else
                    ' ステータス更新（ファイル取得OK）
                    Worksheets(glbSheetName).Cells(i, glbSKU_col + 1).Value = "OK"
                End If
            Else
                ' ステータス更新（JAN未指定）
                Worksheets(glbSheetName).Cells(i, glbSKU_col + 1).Value = "-"
            End If
        Next i
    End If
    
    'FTP接続 クローズ
    FTP.Close
End Sub
