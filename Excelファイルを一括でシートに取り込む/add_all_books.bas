Attribute VB_Name = "add_all_books"
Option Explicit

Sub addAllReports()
    Dim sFile As String
    Dim sWB As Workbook, dWB As Workbook
    Dim dSheetCount As Long
    Dim i, j As Long
    Dim cntSheetCopied As Integer
    Dim sheetName As String
    
    Dim sourceDir As String
    Const DEFAULT_FILE_NAME As String = "AllReports.xlsx"
    
    
    ' フォルダが指定されているかチェック
    If ActiveSheet.Cells(3, 2).Value = "" Then
        MsgBox "まとめたいファイルの場所を指定してください。", vbExclamation
        Exit Sub
    Else
        sourceDir = ActiveSheet.Cells(3, 2).Value
    End If
    
    '画面ちらつき防止
    Application.ScreenUpdating = False
    
    '指定したフォルダ内にあるブックのファイル名を取得
    sFile = Dir(sourceDir & "*.xls")
    
    'フォルダ内にブックがなければ終了
    If sFile = "" Then Exit Sub
    
    '集約用ブックを作成
    Set dWB = Workbooks.Add
    
    '集約前のシート数を取得（後で削除するため）
    dSheetCount = dWB.Worksheets.Count
    
    'シート数のカウンタ初期化
    cntSheetCopied = 0
    
    Do
        'コピー元のブックを開く
        Set sWB = Workbooks.Open(Filename:=sourceDir & sFile)
        
    '対象のブックに含まれるシートを全てコピー
    For j = sWB.Worksheets.Count To 1 Step -1
        sheetName = sWB.Worksheets(j).Name
        sWB.Worksheets(sheetName).Copy After:=dWB.Worksheets(cntSheetCopied + dSheetCount)
        
        'コピーしたシートのカウント
        cntSheetCopied = cntSheetCopied + 1
        
        
        'シート名末尾にコピーNoを追加
        'ActiveSheet.Name = ActiveSheet.Name & "_" & cntSheetCopied
    Next
                
        'コピー元ファイルを閉じる
        sWB.Close SaveChanges:=False
        
        '次のブックのファイル名を取得
        sFile = Dir()
        
        ' レポートファイルと同名があった場合はスキップ
        If sFile = DEFAULT_FILE_NAME Then
            sFile = Dir()
        End If
    Loop While sFile <> ""
        
    
    '集約用ブック作成時にあったシートを削除
    Application.DisplayAlerts = False
    For i = dSheetCount To 1 Step -1
        dWB.Worksheets(i).Delete
    Next i
    
    Application.DisplayAlerts = True
    
On Error Resume Next
    '集約用ブックを保存して閉じる
    'dWB.SaveAs Filename:=sourceDir & DEFAULT_FILE_NAME
    'dWB.Close SaveChanges:=False
    
    ' 画面ちらつき防止 ここまで
    Application.ScreenUpdating = False
End Sub



