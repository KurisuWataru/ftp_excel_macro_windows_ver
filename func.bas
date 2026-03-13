Attribute VB_Name = "func"

' ファイルパスシートからデータ範囲を取得する
' @returns {Variant} シートのデータ範囲の値（2次元配列）
Function getfilePathRangeValues()
    
    'ファイルパスシートを読み込み
    Dim filePathSheet As Worksheet
    Dim filePathRange As Range
    Dim lastRow As Long
    Dim lastCol As Long
    Dim lastRowA As Long, lastRowB As Long, lastRowC As Long, lastRowD As Long
    
    Set filePathSheet = ThisWorkbook.Sheets("ファイルパス")
    
    'A列からD列の各最終行を取得
    lastRowA = filePathSheet.Cells(filePathSheet.Rows.Count, "A").End(xlUp).Row
    lastRowB = filePathSheet.Cells(filePathSheet.Rows.Count, "B").End(xlUp).Row
    lastRowC = filePathSheet.Cells(filePathSheet.Rows.Count, "C").End(xlUp).Row
    lastRowD = filePathSheet.Cells(filePathSheet.Rows.Count, "D").End(xlUp).Row
    
    '最大値を取得
    lastRow = Application.WorksheetFunction.Max(lastRowA, lastRowB, lastRowC, lastRowD)
    
    '最終列を1行目で判定
    lastCol = filePathSheet.Cells(1, filePathSheet.Columns.Count).End(xlToLeft).Column
    
    'データが存在する範囲のみを取得
    Set filePathRange = filePathSheet.Range(filePathSheet.Cells(1, 1), filePathSheet.Cells(lastRow, lastCol))
    
    getfilePathRangeValues = filePathRange.value
    
End Function

' 2次元配列の1行目（ヘッダー行）を取得する
' @param {Variant} ary - 2次元配列
' @returns {Variant} 1行目の配列
Function getHeader(ary As Variant)

    Dim header As Variant
    
    header = Application.Index(ary, 1, 0)
    
    getHeader = header
    
End Function

' ヘッダー行から列名とインデックスの辞書を作成する
' @description 列名をキーとし、インデックスを値とするコレクションを生成
' @param {Variant} ary - 2次元配列
' @returns {Object} 列名をキー、インデックスを値とするコレクション
Function makeHeaderDict(ary As Variant)

    Dim header As Variant
    Dim i As Long
    Dim filePathRangeHeaderDict As Object
    
    Set filePathRangeHeaderDict = CreateObject("Scripting.Dictionary")
    
    header = getHeader(ary)
    
    For i = LBound(header) To UBound(header)
    
        filePathRangeHeaderDict.Add header(i), i
        
    Next i
    
    Set makeHeaderDict = filePathRangeHeaderDict
End Function

' ルートディレクトリのパスを処理し、適切な区切り文字で終了させる
' @param {String} path - 既存のパス（参照用）
' @param {String} rootPath - 処理対象のルートパス
' @param {String} separator - パス区切り文字（"¥" または "/"）
' @param {Long} rowIndex - 処理中の行番号
' @param {String} pathType - パスの種類（"ローカル" または "リモート"）
' @returns {String} 処理済みのルートパス
' @throws {MsgBox} パスが空で先頭行の場合にエラーメッセージを表示
Function processRootPath(ByVal path As String, ByVal rootPath As String, ByVal separator As String, ByVal rowIndex As Long, ByVal pathType As String) As String
    'パスが空の場合のチェック
    If rootPath = "" Then
        If rowIndex = 2 Then
            MsgBox "先頭に" & pathType & "のルートディレクトリのパスを入力してください"
            Exit Function
        End If
        processRootPath = path
        Exit Function
    End If
    
    'パスが区切り文字で終了しているか確認。終了していない場合追加
    If Right(rootPath, 1) <> separator Then
        processRootPath = rootPath & separator
    Else
        processRootPath = rootPath
    End If
End Function

' ローカルパスを組み立てる（Windows形式のパス区切り文字を使用）
' @param {String} rootPath - ルートディレクトリのパス
' @param {String} relativePath - 相対パス
' @param {String} pathDescription - パスの説明（エラーメッセージ用）
' @returns {String} 組み立てられた完全なローカルパス
' @throws {MsgBox} 相対パスが空の場合にエラーメッセージを表示
' @throws {End} 相対パスが空の場合にプログラムを終了
Function buildLocalPath(ByVal rootPath As String, ByVal relativePath As String, ByVal pathDescription As String) As String
    'ルートパスが空の場合は処理しない
    If rootPath = "" Then
        Exit Function
    End If
    
    'relative pathが入力されているか確認
    If relativePath <> "" Then
        'relativePathの先頭と末尾の「¥」を削除
        If Left(relativePath, 1) = "¥" Then
            relativePath = Mid(relativePath, 2)
        End If
        If Right(relativePath, 1) = "¥" Then
            relativePath = Left(relativePath, Len(relativePath) - 1)
        End If
        buildLocalPath = rootPath & relativePath
    Else
        MsgBox pathDescription & "を必ず入力してください"
        End
    End If
End Function

' アップロード時削除フラグをPowerShell形式に変換する
' @param {String} deleteFlag - アップロード時削除フラグ（TRUE/FALSE または True/False）
' @returns {String} PowerShell形式のアップロード時削除フラグ（$true または $false）
Function convertDeleteFlag(ByVal deleteFlag As String) As String
    If deleteFlag = "True" Then
        convertDeleteFlag = "$true"
    ElseIf deleteFlag = "False" Then
        convertDeleteFlag = "$false"
    Else
        convertDeleteFlag = "$false"  ' デフォルトは削除しない
    End If
End Function

' 処理モードが「削除のみ」かどうかを判定する
' @param {String} modeValue - 処理モード列の値
' @returns {Boolean} 「削除のみ」の場合 True
Function isDeleteOnlyMode(ByVal modeValue As String) As Boolean
    isDeleteOnlyMode = (Trim(modeValue) = "削除のみ")
End Function

' リモートパスを組み立てる（Unix形式のパス区切り文字を使用）
' @param {String} rootPath - ルートディレクトリのパス
' @param {String} relativePath - 相対パス
' @param {String} pathDescription - パスの説明（エラーメッセージ用）
' @returns {String} 組み立てられた完全なリモートパス
' @throws {MsgBox} 相対パスが空の場合にエラーメッセージを表示
' @throws {End} 相対パスが空の場合にプログラムを終了
Function buildRemotePath(ByVal rootPath As String, ByVal relativePath As String, ByVal pathDescription As String) As String
    'ルートパスが空の場合は処理しない
    If rootPath = "" Then
        Exit Function
    End If
    
    'relative pathが入力されているか確認
    If relativePath <> "" Then
        'relativePathの先頭と末尾の「/」を削除
        If Left(relativePath, 1) = "/" Then
            relativePath = Mid(relativePath, 2)
        End If
        If Right(relativePath, 1) = "/" Then
            relativePath = Left(relativePath, Len(relativePath) - 1)
        End If
        buildRemotePath = rootPath & relativePath
    Else
        MsgBox pathDescription & "を必ず入力してください"
        End
    End If
End Function

' Collectionの内容をテキストファイルに書き込む（ANSI形式）
' @param {Collection} batchCommand - 書き込むコマンドのコレクション
' @param {String} filePath - 出力ファイルのパス
' @returns {void}
Function writeFile(ByVal batchCommand As Collection, ByVal filePath As String)
    Dim fileNum As Integer
    Dim cmd As Variant
    
    'ファイル番号を取得
    fileNum = FreeFile
    
    'ファイルを書き込みモードでオープン
    Open filePath For Output As #fileNum
    
    'Collectionの各要素を1行ずつ書き込み
    For Each cmd In batchCommand
        Print #fileNum, cmd
    Next cmd
    
    'ファイルを閉じる
    Close #fileNum
End Function

' Collectionの内容をUTF-8形式のテキストファイルに書き込む
' @param {Collection} psCommand - 書き込むコマンドのコレクション
' @param {String} filePath - 出力ファイルのパス
' @returns {void}
Function writeFileUTF8(ByVal psCommand As Collection, ByVal filePath As String)
    Dim stream As Object
    Dim cmd As Variant
    Dim text As String

    ' Collectionの内容を1つの文字列にまとめる（改行区切り）
    For Each cmd In psCommand
        text = text & cmd & vbCrLf
    Next cmd

    ' ADODB.Streamを使ってUTF-8で書き込み
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' テキスト
    stream.Charset = "UTF-8"
    stream.Open
    stream.WriteText text
    stream.SaveToFile filePath, 2 ' 2=上書き
    stream.Close
    Set stream = Nothing
End Function

