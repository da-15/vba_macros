Attribute VB_Name = "mod_meisai"
'--------------------------------------------------
' m o d _ m e i s a i
'
' 明細一覧の顧客データ元に、明細表を複製する。
'
' 2011.03.10 
'
Option Explicit

Private Const START_DATA_ROW As Integer = 13 'データ開始行

'タイトル、名称
Private Const SHEET_NAME_SETUP As String = "設定" '設定シート名
Private Const SHEET_NAME_TEMPLATE As String = "テンプレート" 'テンプレートシート名
Private Const SHEET_NAME_PREFIX As String = "支払依頼書" '複製されるシートの名前

'メッセージ
Private Const MSG_1 As String = "古い支払依頼書が全て削除されますが、よろしいですか？"
Private Const MSG_2 As String = "月代"
Private Const MSG_3 As String = "データの入力がありません。"

'リプレイス対象文字
Private Const REPLACE_WORD As String = "[month]"


' 明細を複製する
Public Sub createMeisai()
    Dim sheetName As String
    Dim i As Integer
    Dim intEndRow As Integer
    
    'データ行の有無の確認
    intEndRow = Worksheets(SHEET_NAME_SETUP).Cells(Rows.Count, 2).End(xlUp).Row
    If (intEndRow < START_DATA_ROW) Then
        Call MsgBox(MSG_3, vbExclamation)
        Exit Sub
    End If
    
    ' 明細削除許可メッセージ
    If (Worksheets.Count > 2) Then
        If (MsgBox(MSG_1, vbOKCancel + vbQuestion) = vbCancel) Then
            Exit Sub
        Else
            Call deleteSheets
        End If
    End If
    
    ' 画面ちらつき防止
    Application.ScreenUpdating = False
    
    
    For i = START_DATA_ROW To intEndRow
        'シート（テンプレート）のコピー
        Sheets(SHEET_NAME_TEMPLATE).Copy After:=Worksheets(Worksheets.Count)
        sheetName = SHEET_NAME_PREFIX & " (" & Worksheets.Count - 2 & ")"
        Worksheets(Worksheets.Count).Name = sheetName 'シート名変更
        Worksheets(Worksheets.Count).Unprotect 'シート保護の解除
        
        '明細内容のセット
        Call setMeisai(i, sheetName)
    Next
    
    Sheets(SHEET_NAME_SETUP).Activate
    
    ' 画面ちらつき防止解除
    Application.ScreenUpdating = True

End Sub

'テンプレートをコピーしたシートに必要情報を埋める
Private Sub setMeisai(index As Integer, sheetName As String)
    Dim intRowNum As Integer
    Dim bufString As String
    
    'データの該当する行
    intRowNum = index '+ START_DATA_ROW - 1
    

    With Worksheets(SHEET_NAME_SETUP)
        '申請日
        Worksheets(sheetName).Cells(6, 3).Value = .Cells(8, 2).Value
        Worksheets(sheetName).Cells(6, 8).Value = .Cells(8, 2).Value
        
        '支払先
        Worksheets(sheetName).Cells(13, 15).Value = .Cells(intRowNum, 2).Value
        
        '金額
        Worksheets(sheetName).Cells(13, 24).Value = .Cells(intRowNum, 4).Value
        
        '発生内容
        Worksheets(sheetName).Cells(17, 7).Value = .Cells(intRowNum, 3).Value
        
        
        
    End With
    
End Sub

' 設定とテンプレートシートを除いた不要なシートを一括削除する
Private Sub deleteSheets()
    Dim i As Integer
    Dim sheetName As String
    For i = Worksheets.Count To 1 Step -1
        sheetName = Worksheets(i).Name
        If (sheetName <> SHEET_NAME_SETUP And sheetName <> SHEET_NAME_TEMPLATE) Then
            Application.DisplayAlerts = False
            Sheets(sheetName).Delete
            Application.DisplayAlerts = True
        End If
    Next
End Sub

