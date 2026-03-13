Attribute VB_Name = "main"
Sub main()
    
    '変数を宣言
    Dim filePathRangeValues As Variant
    Dim filePathRangeHeaderDict As Object
    Dim localRootDirPath As String
    Dim remoteRootDirPath As String
    Dim i As Long
    Dim psCommand As New Collection
    Dim localPaths As New Collection
    Dim remotePaths As New Collection
    Dim deleteFlags As New Collection
    Dim processingModes As New Collection
    
    'ファイルパスシートのセルの値を取得
    filePathRangeValues = getfilePathRangeValues()
    
    'ファイルパスシートのヘッダーのインデックスを取得
    Set filePathRangeHeaderDict = makeHeaderDict(filePathRangeValues)
    'データの行数チェック
    If UBound(filePathRangeValues) < 2 Then
        MsgBox "データが不足しています。最低でも2行（ヘッダー行 + データ行）が必要です。"
        Exit Sub
    End If
    
    'PowerShellスクリプトの可変部分を作成
    For i = 2 To UBound(filePathRangeValues)
        
        '2番目のインデックス（2行目）の空白チェック
        If i = 2 Then
            Dim emptyColumns As String
            Dim headerName As Variant
            Dim skipColumn As Boolean
            emptyColumns = ""
            
            ' 2行目の処理モードを事前に取得
            Dim firstRowMode As String
            If filePathRangeHeaderDict.Exists("処理モード") Then
                firstRowMode = CStr(filePathRangeValues(i, filePathRangeHeaderDict("処理モード")) & "")
            Else
                firstRowMode = ""
            End If
            
            For Each headerName In filePathRangeHeaderDict.Keys
                skipColumn = False
                
                ' アップロード時削除フラグ・処理モードのカラムはチェックをスキップ
                If headerName = "アップロード時削除フラグ" Or headerName = "処理モード" Then
                    skipColumn = True
                End If
                
                ' 削除のみモードの場合、ローカルパス関連列もスキップ
                If isDeleteOnlyMode(firstRowMode) Then
                    If headerName = "ローカルのルートディレクトリのパス" Or _
                       headerName = "アップロード対象のファイル・フォルダの相対パス" Then
                        skipColumn = True
                    End If
                End If
                
                If Not skipColumn Then
                    Dim colIndex As Long
                    colIndex = filePathRangeHeaderDict(headerName)
                    Debug.Print "行" & i & "列" & colIndex & "(" & headerName & ")の値: " & filePathRangeValues(i, colIndex)
                    If IsEmpty(filePathRangeValues(i, colIndex)) Or Trim(filePathRangeValues(i, colIndex)) = "" Then
                        If emptyColumns = "" Then
                            emptyColumns = "「" & headerName & "」"
                        Else
                            emptyColumns = emptyColumns & ", 「" & headerName & "」"
                        End If
                    End If
                End If
            Next headerName
            
            If emptyColumns <> "" Then
                MsgBox "2行目の" & emptyColumns & "が空白です。すべての項目を入力してください。"
                Exit Sub
            End If
        End If
        
        '処理モードの読み取り
        Dim currentMode As String
        If filePathRangeHeaderDict.Exists("処理モード") Then
            currentMode = CStr(filePathRangeValues(i, filePathRangeHeaderDict("処理モード")) & "")
        Else
            currentMode = ""
        End If
        processingModes.Add currentMode
        
        If isDeleteOnlyMode(currentMode) Then
            '削除のみモードではローカルパスは不要
            localPaths.Add ""
        Else
            'ローカルパスの処理
            localRootDirPath = processRootPath(localRootDirPath, filePathRangeValues(i, filePathRangeHeaderDict("ローカルのルートディレクトリのパス")), "¥", i, "ローカル")
            localPaths.Add buildLocalPath(localRootDirPath, filePathRangeValues(i, filePathRangeHeaderDict("アップロード対象のファイル・フォルダの相対パス")), "アップロード対象のファイル・フォルダの相対パス")
        End If
        
        'リモートパスの処理
        remoteRootDirPath = processRootPath(remoteRootDirPath, filePathRangeValues(i, filePathRangeHeaderDict("リモートのルートディレクトリのパス")), "/", i, "リモート")
        remotePaths.Add buildRemotePath(remoteRootDirPath, filePathRangeValues(i, filePathRangeHeaderDict("アップロード先の相対パス")), "アップロードするディレクトリの相対パス")
        
        'アップロード時削除フラグの処理
        deleteFlags.Add convertDeleteFlag(filePathRangeValues(i, filePathRangeHeaderDict("アップロード時削除フラグ")))
        
    Next i
    
    psCommand.Add "# ================================"
    psCommand.Add "# WinSCP.NET FTPアップロード自動化"
    psCommand.Add "# メイン処理ファイル"
    psCommand.Add "# ================================"
    psCommand.Add ""
    psCommand.Add "# ログファイルの設定"
    psCommand.Add "$logFilePath = ""$PSScriptRoot¥display_log.txt"""
    psCommand.Add ""
    psCommand.Add "# トランスクリプトを開始（すべての出力を自動記録）"
    psCommand.Add "Start-Transcript -Path $logFilePath -Append | Out-Null"
    psCommand.Add ""
    psCommand.Add "# 共通処理ファイルの読み込み"
    psCommand.Add "$ftpFuncPath = ""$PSScriptRoot¥ftp_func.ps1"""
    psCommand.Add "if (-not (Test-Path $ftpFuncPath)) {"
    psCommand.Add "    Write-Host ""共通処理ファイルが見つかりません: $ftpFuncPath"" -ForegroundColor Red"
    psCommand.Add "    Write-Host ""何かキーを押して終了してください..."""
    psCommand.Add "    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')"
    psCommand.Add "    exit 1"
    psCommand.Add "}"
    psCommand.Add "Write-Host ""共通処理を読み込み中..."""
    psCommand.Add ". $ftpFuncPath"
    psCommand.Add ""
    psCommand.Add "# 初期化エラーのチェック"
    psCommand.Add "if ($InitializationError) {"
    psCommand.Add "    Write-Host ""初期化処理でエラーが発生しました。処理を中断します。"" -ForegroundColor Red"
    psCommand.Add "    Write-Host ""何かキーを押して終了してください..."""
    psCommand.Add "    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')"
    psCommand.Add "    exit 1"
    psCommand.Add "}"
    psCommand.Add ""
    psCommand.Add "# 接続しようとしているFTPサーバーのチェック"
    psCommand.Add "$ipCheckTargetPath = ""$PSScriptRoot¥ip_check.ps1"""
    psCommand.Add "if (-not (Test-Path $ipCheckTargetPath)) {"
    psCommand.Add "    Write-Host ""IPチェックファイルが見つかりません: $ipCheckTargetPath"" -ForegroundColor Red"
    psCommand.Add "    Write-Host ""何かキーを押して終了してください..."""
    psCommand.Add "    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')"
    psCommand.Add "    exit 1"
    psCommand.Add "}"
    psCommand.Add "Write-Host ""IPチェックファイルを読み込み中..."""
    psCommand.Add ". $ipCheckTargetPath"
    psCommand.Add ""
    psCommand.Add "# FTPホストのチェック処理"
    psCommand.Add "if ($targetEnvironments.ContainsKey($ftpHost)) {"
    psCommand.Add "    $environmentName = $targetEnvironments[$ftpHost]"
    psCommand.Add "    Write-Host ""接続先環境: $environmentName です。"" -ForegroundColor Green"
    psCommand.Add "    "
    psCommand.Add "    do {"
    psCommand.Add "        Write-Host ""処理を続行しますか？ (yes/no): "" -ForegroundColor Yellow -NoNewline"
    psCommand.Add "        $userResponse = $Host.UI.ReadLine().ToLower()"
    psCommand.Add "        "
    psCommand.Add "        if ($userResponse -eq ""yes"" -or $userResponse -eq ""y"") {"
    psCommand.Add "            Write-Host ""処理を続行します。"" -ForegroundColor Green"
    psCommand.Add "            break"
    psCommand.Add "        }"
    psCommand.Add "        elseif ($userResponse -eq ""no"" -or $userResponse -eq ""n"") {"
    psCommand.Add "            Write-Host ""処理を中断します。"" -ForegroundColor Red"
    psCommand.Add "            Write-Host ""何かキーを押して終了してください..."""
    psCommand.Add "            $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')"
    psCommand.Add "            exit 0"
    psCommand.Add "        }"
    psCommand.Add "        else {"
    psCommand.Add "            Write-Host ""無効な入力です。'yes'または'no'を入力してください。"" -ForegroundColor Red"
    psCommand.Add "        }"
    psCommand.Add "    } while ($true)"
    psCommand.Add "} else {"
    psCommand.Add "    Write-Host ""警告: 接続先IPアドレス '$ftpHost' は確認済環境ではありません。"" -ForegroundColor Yellow"
    psCommand.Add "    Write-Host ""確認済環境一覧:"" -ForegroundColor Yellow"
    psCommand.Add "    foreach ($ip in $targetEnvironments.Keys) {"
    psCommand.Add "        Write-Host ""  $ip : $($targetEnvironments[$ip])"" -ForegroundColor Yellow"
    psCommand.Add "    }"
    psCommand.Add "    Write-Host """""
    psCommand.Add "    "
    psCommand.Add "    do {"
    psCommand.Add "        Write-Host ""未確認環境への接続を続行しますか？ (yes/no): "" -ForegroundColor Yellow -NoNewline"
    psCommand.Add "        $userResponse = $Host.UI.ReadLine().ToLower()"
    psCommand.Add "        "
    psCommand.Add "        if ($userResponse -eq ""yes"" -or $userResponse -eq ""y"") {"
    psCommand.Add "            Write-Host ""処理を続行します。"" -ForegroundColor Green"
    psCommand.Add "            break"
    psCommand.Add "        }"
    psCommand.Add "        elseif ($userResponse -eq ""no"" -or $userResponse -eq ""n"") {"
    psCommand.Add "            Write-Host ""処理を中断します。"" -ForegroundColor Red"
    psCommand.Add "            Write-Host ""何かキーを押して終了してください..."""
    psCommand.Add "            $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')"
    psCommand.Add "            exit 0"
    psCommand.Add "        }"
    psCommand.Add "        else {"
    psCommand.Add "            Write-Host ""無効な入力です。'yes'または'no'を入力してください。"" -ForegroundColor Red"
    psCommand.Add "        }"
    psCommand.Add "    } while ($true)"
    psCommand.Add "}"
    psCommand.Add ""
    psCommand.Add "# メイン処理"
    psCommand.Add "$session = New-Object WinSCP.Session"
    psCommand.Add ""
    psCommand.Add "try {"
    psCommand.Add "    # ログ設定"
    psCommand.Add "    $session.SessionLogPath = $LOG_PATH"
    psCommand.Add ""
    psCommand.Add "    # 進捗表示イベントハンドラを設定（成功/失敗を出し分け）"
    psCommand.Add "    $session.add_FileTransferred({"
    psCommand.Add "        param($sender, $e)"
    psCommand.Add "        $fileNameOnly = Split-Path $e.FileName -Leaf"
    psCommand.Add "        if ($e.Error -eq $null) {"
    psCommand.Add "            Write-Host (""{0} 完了"" -f $fileNameOnly)"
    psCommand.Add "        } else {"
    psCommand.Add "            Write-Host (""{0} 失敗: {1}"" -f $fileNameOnly, $e.Error.Message) -ForegroundColor Red"
    psCommand.Add "        }"
    psCommand.Add "    })"
    psCommand.Add ""
    psCommand.Add "    # 転送進捗イベント（転送中の失敗も捕捉して記録）"
    psCommand.Add "    $session.add_FileTransferProgress({"
    psCommand.Add "        param($e)"
    psCommand.Add "        if ($e.Error -ne $null -and $e.FileName) {"
    psCommand.Add "            if (-not $global:BatchErrorFiles) { $global:BatchErrorFiles = @{} }"
    psCommand.Add "            $global:BatchErrorFiles[$e.FileName] = $true"
    psCommand.Add "        }"
    psCommand.Add "    })"
    psCommand.Add ""
    psCommand.Add "    # エラー発生時も処理を継続するためのハンドラ"
    psCommand.Add "    $session.add_QueryReceived({"
    psCommand.Add "        param($sender, $e)"
    psCommand.Add "        Write-Host (""バッチエラー: {0}"" -f $e.Message) -ForegroundColor Yellow"
    psCommand.Add "        # 失敗したローカルファイルパスをログから抽出し、共有スコープに記録"
    psCommand.Add "        $failedPath = Get-FilePathFromErrorMessage -errorMessage $e.Message"
    psCommand.Add "        if ($failedPath) {"
    psCommand.Add "            if (-not $global:BatchErrorFiles) { $global:BatchErrorFiles = @{} }"
    psCommand.Add "            if (-not $global:BatchErrorFiles.ContainsKey($failedPath)) { $global:BatchErrorFiles[$failedPath] = $true }"
    psCommand.Add "        }"
    psCommand.Add "        $e.Continue()"
    psCommand.Add "    })"
    psCommand.Add ""
    psCommand.Add "    # FTPサーバーに接続"
    psCommand.Add "    Write-Host ""FTPサーバーに接続中..."""
    psCommand.Add "    $session.Open($sessionOptions)"
    psCommand.Add "    Write-Host ""接続が成功しました。"""
    psCommand.Add ""
    psCommand.Add "    # ファイル処理を実行"
    
    For i = 1 To localPaths.Count
        If isDeleteOnlyMode(processingModes(i)) Then
            psCommand.Add "deleteAction " & """" & remotePaths(i) & """" & " $session"
        Else
            psCommand.Add "action " & """" & remotePaths(i) & """" & " " & """" & localPaths(i) & """" & " " & deleteFlags(i) & " $session"
        End If
    Next i

    psCommand.Add ""
    psCommand.Add "} catch [WinSCP.SessionException] {"
    psCommand.Add "    Write-Host ""WinSCPセッションエラーが発生しました: $($_.Exception.Message)"" -ForegroundColor Red"
    psCommand.Add "    Write-Host ""FTPの操作前に処理を終了しました。""
    psCommand.Add "} catch {"
    psCommand.Add "    if ($_.Exception.Message -match ""BackupError:(.+)"") {"
    psCommand.Add "        Write-Host ""バックアップ処理でエラーが発生したため、安全のため処理を中断しました。"" -ForegroundColor Red"
    psCommand.Add "    } else {"
    psCommand.Add "        Write-Host ""エラーが発生しました: $_"" -ForegroundColor Red"
    psCommand.Add "    }"
    psCommand.Add "} finally {"
    psCommand.Add "    # セッションを適切に終了"
    psCommand.Add "    if ($session -ne $null) {"
    psCommand.Add "        try {"
    psCommand.Add "            $session.Dispose()"
    psCommand.Add "        } catch {"
    psCommand.Add "            Write-Host ""セッションの破棄中にエラーが発生しました: $_"" -ForegroundColor Red"
    psCommand.Add "        }"
    psCommand.Add "    }"
    psCommand.Add ""
    psCommand.Add "    Write-Host ""`n================ 処理結果 ================="""
    psCommand.Add "    Write-Host ""[バックアップ成功]"" -ForegroundColor Green"
    psCommand.Add "    if ($BackupSuccess.Count) { "
    psCommand.Add "        Write-Host ""$($BackupSuccess.Count)件"""
    psCommand.Add "        $BackupSuccess | ForEach-Object { Write-Host ""  $_"" -ForegroundColor Green } "
    psCommand.Add "    } else { "
    psCommand.Add "        Write-Host ""  なし"" -ForegroundColor Gray "
    psCommand.Add "    }"
    psCommand.Add "    Write-Host ""[バックアップ失敗]"" -ForegroundColor Red"
    psCommand.Add "    if ($BackupFailed.Count) { "
    psCommand.Add "        Write-Host ""$($BackupFailed.Count)件"""
    psCommand.Add "        $BackupFailed | ForEach-Object { Write-Host ""  $_"" -ForegroundColor Red } "
    psCommand.Add "    } else { "
    psCommand.Add "        Write-Host ""  なし"" -ForegroundColor Gray "
    psCommand.Add "    }"
    psCommand.Add "    Write-Host ""[バックアップ不要]"" -ForegroundColor Yellow"
    psCommand.Add "    if ($BackupNotNeeded.Count) { "
    psCommand.Add "        Write-Host ""$($BackupNotNeeded.Count)件"""
    psCommand.Add "        $BackupNotNeeded | ForEach-Object { Write-Host ""  $_"" -ForegroundColor Yellow } "
    psCommand.Add "    } else { "
    psCommand.Add "        Write-Host ""  なし"" -ForegroundColor Gray "
    psCommand.Add "    }"
    psCommand.Add "    Write-Host ""[アップロード成功]"" -ForegroundColor Green"
    psCommand.Add "    if ($UploadSuccess.Count) { "
    psCommand.Add "        # 成功したファイルの総数を計算"
    psCommand.Add "        $successFileCount = 0"
    psCommand.Add "        foreach ($path in $UploadSuccess) {"
    psCommand.Add "            if ($UploadSuccessFiles.ContainsKey($path)) {"
    psCommand.Add "                $successFileCount += $UploadSuccessFiles[$path].Count"
    psCommand.Add "            } else {"
    psCommand.Add "                $successFileCount += 1  # 単一ファイルの場合"
    psCommand.Add "            }"
    psCommand.Add "        }"
    psCommand.Add "        Write-Host ""$($successFileCount)件"""
    psCommand.Add "        $UploadSuccess | ForEach-Object { "
    psCommand.Add "            Write-Host ""  $_"" -ForegroundColor Green"
    psCommand.Add "            # 個別ファイルの詳細を表示"
    psCommand.Add "            if ($UploadSuccessFiles.ContainsKey($_)) {"
    psCommand.Add "                $UploadSuccessFiles[$_] | ForEach-Object { Write-Host ""`t$_"" -ForegroundColor Green }"
    psCommand.Add "            }"
    psCommand.Add "        } "
    psCommand.Add "    } else { "
    psCommand.Add "        Write-Host ""  なし"" -ForegroundColor Gray "
    psCommand.Add "    }"
    psCommand.Add "    Write-Host ""[アップロード失敗]"" -ForegroundColor Red"
    psCommand.Add "    if ($UploadFailed.Count) { "
    psCommand.Add "        # 失敗したファイルの総数を計算"
    psCommand.Add "        $failedFileCount = 0"
    psCommand.Add "        foreach ($path in $UploadFailed) {"
    psCommand.Add "            if ($UploadFailedFiles.ContainsKey($path)) {"
    psCommand.Add "                $failedFileCount += $UploadFailedFiles[$path].Count"
    psCommand.Add "            } else {"
    psCommand.Add "                $failedFileCount += 1  # 単一ファイルの場合"
    psCommand.Add "            }"
    psCommand.Add "        }"
    psCommand.Add "        Write-Host ""$($failedFileCount)件"""
    psCommand.Add "        $UploadFailed | ForEach-Object { "
    psCommand.Add "            Write-Host ""  $_"" -ForegroundColor Red"
    psCommand.Add "            # 個別ファイルの詳細を表示"
    psCommand.Add "            if ($UploadFailedFiles.ContainsKey($_)) {"
    psCommand.Add "                $UploadFailedFiles[$_] | ForEach-Object { Write-Host ""`t$_"" -ForegroundColor Red }"
    psCommand.Add "            }"
    psCommand.Add "        } "
    psCommand.Add "    } else { "
    psCommand.Add "        Write-Host ""  なし"" -ForegroundColor Gray "
    psCommand.Add "    }"
    psCommand.Add "    Write-Host ""[削除成功]"" -ForegroundColor Green"
    psCommand.Add "    if ($DeleteSuccess.Count) { "
    psCommand.Add "        Write-Host ""$($DeleteSuccess.Count)件"""
    psCommand.Add "        $DeleteSuccess | ForEach-Object { Write-Host ""  $_"" -ForegroundColor Green } "
    psCommand.Add "    } else { "
    psCommand.Add "        Write-Host ""  なし"" -ForegroundColor Gray "
    psCommand.Add "    }"
    psCommand.Add "    Write-Host ""[削除失敗]"" -ForegroundColor Red"
    psCommand.Add "    if ($DeleteFailed.Count) { "
    psCommand.Add "        Write-Host ""$($DeleteFailed.Count)件"""
    psCommand.Add "        $DeleteFailed | ForEach-Object { Write-Host ""  $_"" -ForegroundColor Red } "
    psCommand.Add "    } else { "
    psCommand.Add "        Write-Host ""  なし"" -ForegroundColor Gray "
    psCommand.Add "    }"
    psCommand.Add "    Write-Host ""[削除スキップ]"" -ForegroundColor Yellow"
    psCommand.Add "    if ($DeleteSkipped.Count) { "
    psCommand.Add "        Write-Host ""$($DeleteSkipped.Count)件"""
    psCommand.Add "        $DeleteSkipped | ForEach-Object { Write-Host ""  $_"" -ForegroundColor Yellow } "
    psCommand.Add "    } else { "
    psCommand.Add "        Write-Host ""  なし"" -ForegroundColor Gray "
    psCommand.Add "    }"
    psCommand.Add "    Write-Host ""===============================================`n"""
    psCommand.Add "    Write-Host ""処理が完了しました。何かキーを押して終了してください..."""
    psCommand.Add "    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')"
    psCommand.Add "}"
    psCommand.Add ""
    psCommand.Add "# トランスクリプトを停止"
    psCommand.Add "Stop-Transcript | Out-Null"
    
    'PowerShellスクリプトを出力
    Call writeFileUTF8(psCommand, ThisWorkbook.path & "¥main_process.ps1")

    MsgBox "ファイルの作成が完了しました。"
    
End Sub

Sub makeConfig()

    Dim ftpHost As String
    Dim ftpUser As String
    Dim ftpPassword As String
    Dim winscpDllPath As String
    Dim backupPath As String

    dim psCommand as new collection

    ftpHost = ThisWorkbook.Sheets("基本設定").Range("A2").Value
    ftpUser = ThisWorkbook.Sheets("基本設定").Range("A5").Value
    ftpPassword = ThisWorkbook.Sheets("基本設定").Range("A8").Value
    winscpDllPath = ThisWorkbook.Sheets("基本設定").Range("A11").Value
    backupPath = ThisWorkbook.Sheets("基本設定").Range("A14").Value

    if ftpHost = "" or ftpUser = "" or ftpPassword = "" or winscpDllPath = "" or backupPath = "" then
        MsgBox "すべての項目を入力してください。"
        Exit Sub
    End If
    
    psCommand.Add "# FTP接続設定"
    psCommand.Add "$ftpHost = '" & ftpHost & "'"
    psCommand.Add "$ftpUser = '" & ftpUser & "'"
    psCommand.Add "$ftpPassword = '" & ftpPassword & "'"
    psCommand.Add ""
    psCommand.Add "# WinSCP.NET.dllのパス"
    psCommand.Add "$winscpDllPath = '" & winscpDllPath & "'"
    psCommand.Add ""
    psCommand.Add "# バックアップファイルの保存先パス"
    psCommand.Add "$backupPath = '" & backupPath & "'"

    'config.ps1を作成
    Call writeFileUTF8(psCommand, ThisWorkbook.path & "¥config.ps1")

    MsgBox "ファイルの作成が完了しました。"

End Sub