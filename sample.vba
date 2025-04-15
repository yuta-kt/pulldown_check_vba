Sub 勘定科目ごとにプルダウン設定()
    Dim ws As Worksheet
    Dim プルダウンシート As Worksheet
    Dim コメントシート As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim 勘定科目 As String
    Dim プルダウン範囲名 As String
    
    ' シートの設定
    Set コメントシート = ThisWorkbook.Sheets("コメント")
    Set プルダウンシート = ThisWorkbook.Sheets("プルダウン")
    
    ' コメントシートの最終行を取得（B列の最終行）
    lastRow = コメントシート.Cells(コメントシート.Rows.Count, "B").End(xlUp).Row
    
    ' B列の各勘定科目に対して処理
    For i = 1 To lastRow
        勘定科目 = コメントシート.Cells(i, "B").Value
        
        ' 勘定科目が空でない場合のみ処理
        If 勘定科目 <> "" Then
            ' 勘定科目に基づいてプルダウン範囲名を決定
            Select Case 勘定科目
                Case "売上"
                    プルダウン範囲名 = "売上選択肢"
                Case "原価"
                    プルダウン範囲名 = "原価選択肢"
                Case "原価率"
                    プルダウン範囲名 = "原価率選択肢"
                Case "粗利"
                    プルダウン範囲名 = "粗利選択肢"
                Case "粗利率"
                    プルダウン範囲名 = "粗利率選択肢"
                Case Else
                    プルダウン範囲名 = "その他選択肢"
            End Select
            
            ' プルダウンリストを設定（厳格なバリデーション）
            With コメントシート.Cells(i, "D").Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                     Operator:=xlBetween, Formula1:="=" & プルダウン範囲名
                .IgnoreBlank = False  ' 空白を許可しない
                .InCellDropdown = True
                .InputTitle = "選択肢の入力"
                .ErrorTitle = "入力エラー"
                .InputMessage = "リストから適切な選択肢を選んでください"
                .ErrorMessage = "「" & 勘定科目 & "」に対する正しい選択肢を選択してください。入力されたデータは無効です。"
                .ShowInput = True
                .ShowError = True
            End With
        End If
    Next i
    
    MsgBox "プルダウンの設定が完了しました。", vbInformation
End Sub

' 名前付き範囲を作成するサブルーチン（初期設定用）
Sub 名前付き範囲の作成()
    Dim プルダウンシート As Worksheet
    
    Set プルダウンシート = ThisWorkbook.Sheets("プルダウン")
    
    ' 既存の名前付き範囲を削除
    On Error Resume Next
    ThisWorkbook.Names("売上選択肢").Delete
    ThisWorkbook.Names("原価選択肢").Delete
    ThisWorkbook.Names("原価率選択肢").Delete
    ThisWorkbook.Names("粗利選択肢").Delete
    ThisWorkbook.Names("粗利率選択肢").Delete
    ThisWorkbook.Names("その他選択肢").Delete
    On Error GoTo 0
    
    ' 名前付き範囲の作成
    ' プルダウンシートのA1に「勘定項目」というヘッダーがあると想定
    ' B1から順に売上の選択肢、C1から順に原価の選択肢、という形式を想定
    
    ' 売上選択肢（B列）
    ThisWorkbook.Names.Add Name:="売上選択肢", RefersTo:= _
        "=プルダウン!$B$2:$B$11"
        
    ' 原価選択肢（C列）
    ThisWorkbook.Names.Add Name:="原価選択肢", RefersTo:= _
        "=プルダウン!$C$2:$C$11"
        
    ' 原価率選択肢（D列）
    ThisWorkbook.Names.Add Name:="原価率選択肢", RefersTo:= _
        "=プルダウン!$D$2:$D$11"
        
    ' 粗利選択肢（E列）
    ThisWorkbook.Names.Add Name:="粗利選択肢", RefersTo:= _
        "=プルダウン!$E$2:$E$11"
        
    ' 粗利率選択肢（F列）
    ThisWorkbook.Names.Add Name:="粗利率選択肢", RefersTo:= _
        "=プルダウン!$F$2:$F$11"
        
    ' その他選択肢（共通選択肢として使用可能）
    ThisWorkbook.Names.Add Name:="その他選択肢", RefersTo:= _
        "=プルダウン!$G$2:$G$11"
    
    MsgBox "名前付き範囲の作成が完了しました。", vbInformation
End Sub

' プルダウンシートの初期化
Sub プルダウンシート初期化()
    Dim プルダウンシート As Worksheet
    Dim i As Integer
    
    ' プルダウンシートがなければ作成
    On Error Resume Next
    Set プルダウンシート = ThisWorkbook.Sheets("プルダウン")
    On Error GoTo 0
    
    If プルダウンシート Is Nothing Then
        Set プルダウンシート = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        プルダウンシート.Name = "プルダウン"
    End If
    
    ' ヘッダー設定
    プルダウンシート.Range("A1").Value = "勘定項目"
    プルダウンシート.Range("B1").Value = "売上"
    プルダウンシート.Range("C1").Value = "原価"
    プルダウンシート.Range("D1").Value = "原価率"
    プルダウンシート.Range("E1").Value = "粗利"
    プルダウンシート.Range("F1").Value = "粗利率"
    プルダウンシート.Range("G1").Value = "その他"
    
    ' サンプルデータ設定（各列に選択肢1～10を設定）
    For i = 1 To 10
        プルダウンシート.Cells(i + 1, "B").Value = "売上選択肢" & i
        プルダウンシート.Cells(i + 1, "C").Value = "原価選択肢" & i
        プルダウンシート.Cells(i + 1, "D").Value = "原価率選択肢" & i
        プルダウンシート.Cells(i + 1, "E").Value = "粗利選択肢" & i
        プルダウンシート.Cells(i + 1, "F").Value = "粗利率選択肢" & i
        プルダウンシート.Cells(i + 1, "G").Value = "選択肢" & i
    Next i
    
    ' 列幅の調整
    プルダウンシート.Columns("A:G").AutoFit
    
    MsgBox "プルダウンシートを初期化しました。必要に応じて選択肢を編集してください。", vbInformation
End Sub

' メイン実行プロシージャ
Sub 実行()
    ' プルダウンシートの初期化
    Call プルダウンシート初期化
    
    ' 名前付き範囲の作成
    Call 名前付き範囲の作成
    
    ' プルダウンの設定
    Call 勘定科目ごとにプルダウン設定
    
    ' イベント処理の設定
    Call データ貼付検証イベント設定
End Sub

' データの貼り付けや変更を検証するイベント設定
Sub データ貼付検証イベント設定()
    Dim コメントシート As Worksheet
    
    ' イベント処理を有効にする
    Application.EnableEvents = True
    
    ' シートの設定
    Set コメントシート = ThisWorkbook.Sheets("コメント")
    
    ' ThisWorkbookではなくシートで直接設定
    ' コードのモジュールタイプ変更のため通知
    MsgBox "データ貼り付け検証を有効にしました。" & vbCrLf & vbCrLf & _
           "注意: イベント処理を完全に有効にするには、" & vbCrLf & _
           "このコードをシートモジュール「コメント」に移動してください。", vbInformation
End Sub

' ※※※ 以下のコードは「コメント」シートのモジュールに配置してください ※※※
' シートモジュールを開くには、VBエディタでシート「コメント」を右クリックし、
' 「コードの表示」を選択します

' 以下のコードをコメントシートモジュールにコピー＆ペーストしてください
'
' Private Sub Worksheet_Change(ByVal Target As Range)
'     ' D列の変更のみを監視
'     If Not Intersect(Target, Me.Columns("D")) Is Nothing Then
'         On Error GoTo ErrorHandler
'         Application.EnableEvents = False
'         
'         Dim cell As Range
'         Dim 検証対象範囲 As Range
'         Dim 勘定科目 As String
'         Dim 入力値 As Variant
'         Dim プルダウン範囲名 As String
'         Dim 有効な選択肢 As Range
'         Dim 有効値 As Boolean
'         
'         ' 変更された範囲内のD列の各セルに対して処理
'         Set 検証対象範囲 = Intersect(Target, Me.Columns("D"))
'         
'         For Each cell In 検証対象範囲.Cells
'             ' セルが空でなければ検証
'             If Not IsEmpty(cell.Value) Then
'                 ' 同じ行のB列から勘定科目を取得
'                 勘定科目 = Me.Cells(cell.Row, "B").Value
'                 入力値 = cell.Value
'                 
'                 ' 勘定科目に基づいてプルダウン範囲名を決定
'                 Select Case 勘定科目
'                     Case "売上"
'                         プルダウン範囲名 = "売上選択肢"
'                     Case "原価"
'                         プルダウン範囲名 = "原価選択肢"
'                     Case "原価率"
'                         プルダウン範囲名 = "原価率選択肢"
'                     Case "粗利"
'                         プルダウン範囲名 = "粗利選択肢"
'                     Case "粗利率"
'                         プルダウン範囲名 = "粗利率選択肢"
'                     Case Else
'                         プルダウン範囲名 = "その他選択肢"
'                 End Select
'                 
'                 ' 勘定科目が空でない場合のみ検証
'                 If 勘定科目 <> "" Then
'                     ' 名前付き範囲から有効な選択肢を取得
'                     Set 有効な選択肢 = ThisWorkbook.Names(プルダウン範囲名).RefersToRange
'                     
'                     ' 入力値が有効かどうかチェック
'                     有効値 = False
'                     Dim rngCell As Range
'                     For Each rngCell In 有効な選択肢
'                         If rngCell.Value = 入力値 Then
'                             有効値 = True
'                             Exit For
'                         End If
'                     Next rngCell
'                     
'                     ' 入力値が無効な場合はエラーメッセージを表示し、入力をクリア
'                     If Not 有効値 Then
'                         MsgBox "「" & 勘定科目 & "」に対して無効な選択肢「" & 入力値 & "」が入力されました。" & vbCrLf & _
'                                "正しい選択肢を選択してください。", vbExclamation, "入力エラー"
'                         cell.Value = ""
'                     End If
'                 End If
'             End If
'         Next cell
'         
'         Application.EnableEvents = True
'         Exit Sub
'         
' ErrorHandler:
'         Application.EnableEvents = True
'         MsgBox "エラーが発生しました: " & Err.Description, vbCritical
'     End If
' End Sub
'
' Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
'     ' D列でのみ動作
'     If Not Intersect(Target, Me.Columns("D")) Is Nothing Then
'         ' 右クリックメニューに選択肢を表示する代わりに、プルダウンを表示
'         Cancel = False  ' 右クリックメニューを表示
'     End If
' End Sub