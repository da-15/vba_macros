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

Private Const START_ROW As Integer = 10 ' 抽出した文字の書き出しを開始する行
Private Const BUTTON_EXTRACT As String = "cmdButton1"
Private Const BUTTON_CLEAR As String = "cmdButton2"
Private Const CHECKBOX_SORT As String = "cbSort"



' 削除ボタン
Private Sub cmdButton2_Click()
    Call deleteAllShapes
End Sub

' 抽出ボタン
Private Sub cmdButton1_Click()
    Call extractText
End Sub

' 図形からテキストを抽出する
Private Sub extractText()
    Dim j As Integer
    Dim buf As Integer ' DEBUG用
    Dim objShape As Shape
    
    
    '左端のテキスト表示行のクリアー
    Range("A:A").ClearContents
    
    
    ' グループ化解除
    Do While True
        If (ungroupObj() = 0) Then
            Exit Do
        End If
    Loop
    
    
    '10行目から書き出し
    j = START_ROW
    
    '図形からテキストを抽出
    For Each objShape In Shapes
        
        'デバッグ用：buf = Shapes(i).Type
        If (objShape.Type = msoAutoShape Or objShape.Type = msoTextBox) Then
            If objShape.TextEffect.Text <> "" Then
                Cells(j, 1).Value = objShape.TextEffect.Text
                j = j + 1
            End If
        End If
    Next objShape
    
    If cbSort = True Then
        ' チェックボックスにチェックがある場合
        ' 並べ替えを行う
        Call sortText
    End If
End Sub

' グループ解除 -- グループ化を解除した処理回数を返す
Private Function ungroupObj() As Integer
    Dim cntUnGroup As Integer '処理回数
    Dim objShape As Shape
    
    cntUnGroup = 0
    For Each objShape In Shapes
      If (objShape.Type = msoGroup Or objShape.Type = msoPicture) Then
        ' 画像ファイルから作成した図のなかに、グループ化した
        ' オブジェクトが含まれることがある。
            '--------------------------------------------------
            'msoAutoShape         1   オートシェイプ
            'msoCallout           2   吹き出し
            'msoChart             3   埋め込みグラフ
            'msoComment           4   コメント
            'msoFreeform          5   フリーフォーム
            'msoGroup             6   グループ化された図形
            'msoEmbeddedOLEObject 7   埋め込みOLEオブジェクト
            'msoFormControl       8   フォームコントロール
            'msoLine              9   直線・矢印
            'msoLinkedOLEObject  10  リンクOLEオブジェクト
            'msoLinkedPicture    11  元画像にリンクしている図
            'msoOLEControlObject 12  ActiveXコントロール
            'msoPicture          13  画像ファイルから作成した図
            'msoPlaceholder      14  （EXCELでは使用しない）
            'msoTextEffect       15  ワードアート
            'msoMedia            16  （EXCELでは使用しない）
            'msoTextBox          17  テキストボックス
            'msoScriptAnchor     18  スクリプトアンカー
            'msoTable            19  テーブル
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


' 図形を（ボタン以外）全て削除する
Private Sub deleteAllShapes()
    Dim objShape As Shape
    
    For Each objShape In Shapes
      If (objShape.Name <> BUTTON_EXTRACT And _
          objShape.Name <> BUTTON_CLEAR And _
          objShape.Name <> CHECKBOX_SORT) Then
          
          ' 図形の削除
          objShape.Delete
      End If
    Next objShape
End Sub

' 抽出したテキストの並べ替えを行う
Private Sub sortText()
    Dim endRow As Integer
    Dim startRow As Integer
    
    startRow = START_ROW
    endRow = Cells(Rows.count, 1).End(xlUp).Row
    
    If startRow < endRow Then
        'ソート
        Range("A10", "A" & endRow).Sort _
            Key1:=Range("A10", "A" & endRow), _
            Order1:=xlAscending, _
            Header:=xlNo, _
            MatchCase:=False, _
            Orientation:=xlTopToBottom, _
            DataOption1:=xlSortNormal
    End If
    
End Sub

