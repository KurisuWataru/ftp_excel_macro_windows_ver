<# 
========================================
 WinSCP.NET FTPアップロード自動化（共通ライブラリ）
========================================
#>

# ------------------------------------------------------------
# 直接実行はNG（ライブラリとして dot-source される前提）
# ------------------------------------------------------------
if ((Split-Path $MyInvocation.InvocationName -Leaf) -eq $MyInvocation.MyCommand.Name) {
    Write-Host "このファイルは共通処理ライブラリです。" -ForegroundColor Yellow
    Write-Host "FTP処理を実行するには main_process.ps1 を使用してください。" -ForegroundColor Yellow
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    exit 0
}

#region 初期化（設定読み込み / DLLロード / グローバル状態）

# 初期化エラーフラグ（main_process.ps1 側でチェックされます）
$script:InitializationError = $false
$InitializationError = $false

function Write-Info {
    param([string]$message)
    Write-Host $message
}

function Write-Warn {
    param([string]$message)
    Write-Host $message -ForegroundColor Yellow
}

function Write-Err {
    param([string]$message)
    Write-Host $message -ForegroundColor Red
}

function Initialize-ResultState {
    <#
    処理結果（成功/失敗など）を格納する変数を初期化します。
    これらは main_process.ps1 の finally ブロックで集計表示に使われます。
    #>

    # 処理結果集計用の配列
    $script:BackupSuccess   = @()
    $script:BackupFailed    = @()
    $script:BackupNotNeeded = @()
    $script:UploadSuccess   = @()
    $script:UploadFailed    = @()
    $script:DeleteSuccess   = @()
    $script:DeleteFailed    = @()
    $script:DeleteSkipped   = @()

    # 個別ファイルの結果を記録するための辞書（キーは localBasePath）
    $script:UploadSuccessFiles = @{}
    $script:UploadFailedFiles = @{}

    # バッチエラー（QueryReceived / FileTransferProgress）で拾ったローカル側失敗ファイルパスを一時保持
    # ※イベントハンドラから参照するため global を使用
    $global:BatchErrorFiles = @{}
}

function Import-WinScpDll {
    <#
    WinSCP.NET.dll をロードします。
    ※ config.ps1 の読み込みは「必ずスクリプト直下（関数の外）」で行います。
       （関数内で dot-source すると、その変数が関数スコープに閉じてしまい、
        `$ftpHost` 等が後続処理で見えなくなるため）
    #>
    param(
        [Parameter(Mandatory)]
        [string]$dllPath
    )

    if (-not (Test-Path $dllPath)) {
        Write-Err "WinSCP.NET.dll が見つかりません: $dllPath"
        Write-Info "WinSCP.NET.dll をダウンロードして、指定したパスに配置してください。"
        $script:InitializationError = $true
        $InitializationError = $true
        return
    }

    try {
        Add-Type -Path $dllPath
    } catch {
        Write-Err "WinSCP.NET.dll の読み込みに失敗しました: $_"
        $script:InitializationError = $true
        $InitializationError = $true
        return
    }
}

# 初期化を実行
Initialize-ResultState

# 設定ファイルの読み込み（スクリプトスコープで dot-source するのが重要）
$configPath = Join-Path $PSScriptRoot "config.ps1"
if (-not (Test-Path $configPath)) {
    Write-Err "設定ファイルが見つかりません: $configPath"
    Write-Info "エクセルマクロで config.ps1 を作成してください。"
    $script:InitializationError = $true
    $InitializationError = $true
    return
}
. $configPath

Import-WinScpDll -dllPath $winscpDllPath
if ($script:InitializationError -or $InitializationError) {
    # DLLロードに失敗した場合は、ここで止めます（後続の WinSCP 型参照で追加エラーを出さないため）
    return
}

# ログファイルのパスを設定（main_process.ps1 が参照）
$LOG_PATH = Join-Path $PSScriptRoot "WinSCP.log"

# バックアップフォルダのパスを設定（設定ファイルから読み込み）
$BACKUP_PATH = $backupPath

#endregion 初期化

#region WinSCPセッションオプション（読み込み時に作成される）

function Initialize-WinSCPSession {
    <#
    WinSCP への接続情報（SessionOptions）を作成します。
    main_process.ps1 では `$session.Open($sessionOptions)` として使用します。
    #>
    Write-Host "スクリプト実行開始: $(Get-Date)"

    # WinSCP セッションオプションの設定
    $sessionOptions = New-Object WinSCP.SessionOptions
    $sessionOptions.Protocol = [WinSCP.Protocol]::Ftp
    $sessionOptions.HostName = $ftpHost
    $sessionOptions.UserName = $ftpUser
    $sessionOptions.Password = $ftpPassword
    $sessionOptions.FtpSecure = [WinSCP.FtpSecure]::Explicit
    # SSL証明書の検証を無効化
    $sessionOptions.GiveUpSecurityAndAcceptAnyTlsHostCertificate = $true
    
    return $sessionOptions
}

# セッションオプションを初期化（main_process.ps1 が参照）
$sessionOptions = Initialize-WinSCPSession

#endregion WinSCPセッションオプション

#region メイン処理の入口（外部から呼ばれる関数）

function action {
    param (
        [Parameter(Mandatory)][string]$remotePath,
        [Parameter(Mandatory)][string]$localPath,
        [Parameter(Mandatory)][bool]$deleteAfterBackup,
        [Parameter(Mandatory)][WinSCP.Session]$session
    )

    $remoteParentDir = Get-RemoteParentDirectory -remotePath $remotePath

    Backup-RemoteFile -remotePath $remotePath -childPath $remoteParentDir -deleteAfterBackup $deleteAfterBackup -session $session
    Upload-LocalFile -localPath $localPath -childPath $remoteParentDir -session $session
}

function deleteAction {
    <#
    リモートファイル/フォルダのバックアップ取得後に削除のみ行います（アップロードは行いません）。
    バックアップは安全のため必ず取得し、削除前にユーザー確認プロンプトを表示します。
    #>
    param (
        [Parameter(Mandatory)][string]$remotePath,
        [Parameter(Mandatory)][WinSCP.Session]$session
    )

    $remoteParentDir = Get-RemoteParentDirectory -remotePath $remotePath

    Backup-RemoteFile -remotePath $remotePath -childPath $remoteParentDir -deleteAfterBackup $true -session $session
}

#endregion メイン処理の入口

#region バックアップ（リモート → ローカル）

function Backup-RemoteFile {
    param (
        [Parameter(Mandatory)][string]$remotePath,
        [Parameter(Mandatory)][string]$childPath,
        [Parameter(Mandatory)][bool]$deleteAfterBackup,
        [Parameter(Mandatory)][WinSCP.Session]$session
    )

    # 念のため、明示的にルートディレクトリに移動（相対パス事故を防止）
    Move-SessionToRootDirectory -session $session
    Write-Host "リモートのファイルの存在を確認: $remotePath"
    
    try {
        # ここで GetFileInfo が成功すれば、対象が存在する（ファイル/フォルダ）
        $fileInfo = $session.GetFileInfo($remotePath)
        Write-Host "リモートのファイルのバックアップを取得: $remotePath"
        
        $backupDir = Join-Path $BACKUP_PATH $childPath
        if (-not (Test-Path $backupDir)) {
            New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
        }
        
        $transferOptions = New-TransferOptions
        $transferResult = $session.GetFiles($remotePath, "$backupDir\", $False, $transferOptions)
        
        if ($transferResult.IsSuccess) {
            $script:BackupSuccess += $remotePath
            Write-Host "バックアップ成功: $remotePath"

            # バックアップ成功後、アップロード時削除フラグが true の場合は「確認してから」削除します
            if ($deleteAfterBackup) {
                try {
                    $shouldDelete = Read-YesNo -message "バックアップ対象 '$remotePath' を削除しますか？ (yes/no): "
                    if (-not $shouldDelete) {
                        Write-Warn "削除をスキップします。バックアップは完了しました。"
                        $script:DeleteSkipped += $remotePath
                        return
                    }

                    # ファイルとフォルダの両方に RemoveFiles を使用（WinSCP的に扱いやすい）
                    Write-Host "対象を削除: $remotePath" -ForegroundColor Yellow
                    try {
                        $removalResult = $session.RemoveFiles($remotePath)

                        if ($removalResult.IsSuccess) {
                            if ($fileInfo.IsDirectory) {
                                Write-Host "フォルダ削除完了: $remotePath （検証中...）" -ForegroundColor Yellow
                            } else {
                                Write-Host "ファイル削除完了: $remotePath （検証中...）" -ForegroundColor Yellow
                            }

                            # 削除結果の検証（存在確認。無ければ例外になる）
                            # QueryReceived の Continue() により IsSuccess が true でも
                            # 実際には削除されていない場合があるため、検証結果を最終判定とする
                            try {
                                $verifyInfo = $session.GetFileInfo($remotePath)
                                Write-Host "エラー: 削除後もファイル/フォルダがまだ存在しています。削除失敗として記録します。" -ForegroundColor Red
                                $script:DeleteFailed += $remotePath
                            } catch [WinSCP.SessionRemoteException] {
                                Write-Host "検証: 正常に削除されました" -ForegroundColor Green
                                $script:DeleteSuccess += $remotePath
                            }
                        } else {
                            # WinSCPのネイティブなエラーメッセージを使用
                            $winscpErrorMessage = if ($removalResult.Failures.Count -gt 0) {
                                $removalResult.Failures[0].Message
                            } else {
                                "削除に失敗しました"
                            }

                            # 詳細なエラー情報を表示
                            Write-Host "`n================ WinSCP削除エラー ================" -ForegroundColor Red
                            Write-Host "対象パス: $remotePath" -ForegroundColor Red
                            Write-Host "WinSCPエラー: $winscpErrorMessage" -ForegroundColor Red
                            Write-Host "失敗件数: $($removalResult.Failures.Count)" -ForegroundColor Red

                            # 各失敗の詳細を表示
                            foreach ($failure in $removalResult.Failures) {
                                Write-Host "  失敗: $($failure.FileName) - $($failure.Message)" -ForegroundColor Red
                            }
                            Write-Host "===============================================`n" -ForegroundColor Red

                            $script:DeleteFailed += $remotePath
                            # WinSCPのエラーメッセージをそのまま使用して例外を投げる
                            throw $winscpErrorMessage
                        }
                    } catch [WinSCP.SessionRemoteException] {
                        # ファイルが存在しない場合は成功として扱う
                        Write-Host "対象が存在しないため削除をスキップ: $remotePath" -ForegroundColor Yellow
                        $script:DeleteSkipped += $remotePath
                    }
                } catch {
                    # WinSCPのエラーメッセージをそのまま再スロー
                    throw
                }
            }
        } else {
            Write-Host "バックアップ失敗: $remotePath" -ForegroundColor Red
            foreach ($failure in $transferResult.Failures) {
                Write-Host "$($failure.Message)" -ForegroundColor Red
            }
            # バックアップ失敗時は処理を中断
            throw ""
        }
    } catch [WinSCP.SessionRemoteException] {
        Write-Host "リモートにファイルが存在しないため $remotePath はバックアップを作成しませんでした。"
        $script:BackupNotNeeded += $remotePath
        if ($deleteAfterBackup) {
            $script:DeleteSkipped += $remotePath
        }
    } catch {
        $script:BackupFailed += $remotePath
        # バックアップ処理中のエラーでも処理を中断
        throw "BackupError:$remotePath"
    }
}

#endregion バックアップ

#region アップロード（ローカル → リモート）

function Upload-LocalFile {
    param (
        [Parameter(Mandatory)][string]$localPath,
        [Parameter(Mandatory)][string]$childPath,
        [Parameter(Mandatory)][WinSCP.Session]$session
    )
    
    <#
    ローカル側のパス（ファイル/フォルダ）を、指定リモートディレクトリにアップロードします。

    処理の流れ：
    1) PutFiles() 実行
    2) 結果を解析して成功/失敗を集計
    3) 成功したものがあれば、アップロード先に chmod（権限設定）
    #>

    Write-Host "ローカルのファイルをアップロード: $localPath"
    
    try {
        # 念のため、明示的にルートディレクトリに移動（相対パス事故を防止）
        Move-SessionToRootDirectory -session $session
        $transferOptions = New-TransferOptions
        
        # リモートパスの末尾スラッシュを正規化（重複スラッシュ防止）
        $normalizedChildPath = (ConvertTo-UnixPath -path $childPath) -replace '/+$',''
        $transferResult = $session.PutFiles($localPath, "$normalizedChildPath/", $False, $transferOptions)
        
        # フォルダかファイルかを判定して適切な処理関数を呼び出し
        $isDirectory = Test-Path $localPath -PathType Container
        if ($isDirectory) {
            Process-DirectoryUploadResult -localPath $localPath -transferResult $transferResult
        } else {
            Process-SingleFileUploadResult -localPath $localPath -transferResult $transferResult
        }
        
        # アップロード完了後に権限を設定（実際に成功した転送がある場合のみ）
        $hasSuccessfulTransfers = ($transferResult.Transfers | Where-Object { $_.Error -eq $null }).Count -gt 0
        
        if ($hasSuccessfulTransfers) {
            Write-Host "アップロード完了。権限を設定中..."
            
            # 権限設定を実行（パス計算・chmod のいずれが失敗しても処理を継続）
            try {
                $uploadedRemotePath = Get-UploadedRemotePath -normalizedChildPath $normalizedChildPath -localPath $localPath
                Set-RemoteFilePermissions -remotePath $uploadedRemotePath -session $session
            } catch {
                Write-Warn "  警告: 権限設定でエラーが発生しましたが、処理を継続します: $_"
            }
        } else {
            Write-Host "アップロードが完全に失敗したため、権限設定をスキップします。" -ForegroundColor Yellow
        }
    } catch {
        Write-Host "アップロード処理中にエラーが発生しました: $_" -ForegroundColor Red
        $script:UploadFailed += $localPath
    }
}

# フォルダアップロードの結果を処理する関数
function Process-DirectoryUploadResult {
    param (
        [string]$localPath,
        [object]$transferResult
    )
    
    <#
    PutFiles() の戻り値（TransferOperationResult）を見て、
    - どのファイルが成功したか
    - どのファイルが失敗したか
    を「1ファイル単位」で集計します。

    WinSCP の結果は主に2種類あります：
    - Transfers : 実際に転送を試みた各ファイルの結果（成功なら Error が null）
    - Failures  : まとめて実行したときに Transfers だけでは拾えない失敗

    さらに main_process.ps1 側のイベント（QueryReceived / FileTransferProgress）で
    「転送前に弾かれたパス」を $global:BatchErrorFiles に貯めているため、
    それもここで失敗として補完します。
    #>
    
    $successCount = 0
    $failureCount = 0
    
    # 成功/失敗を TransferEventArgs の Error プロパティで判定
    foreach ($transfer in $transferResult.Transfers) {
        if ($transfer.Error -eq $null) {
            if (Add-SuccessFile -localBasePath $localPath -filePath $transfer.FileName) {
                $successCount++
            }
        } else {
            if (Add-FailedFile -localBasePath $localPath -filePath $transfer.FileName -errorMessage $transfer.Error.Message) {
                $failureCount++
            }
        }
    }

    # Transfers に含まれない失敗を Failures から補完
    foreach ($failure in $transferResult.Failures) {
        $failedPath = Get-FilePathFromErrorMessage -errorMessage $failure.Message
        if ($failedPath) {
            if (Add-FailedFile -localBasePath $localPath -filePath $failedPath -errorMessage $failure.Message) {
                $failureCount++
            }
        } else {
            if (Add-FailedFile -localBasePath $localPath -filePath "[ファイル名不明]" -errorMessage $failure.Message) {
                $failureCount++
            }
        }
    }

    # QueryReceived / FileTransferProgress で拾った失敗を補完（取得とクリアを共通化）
    $batchErrorPaths = Get-AndClearBatchErrorFilePaths
    foreach ($absPath in $batchErrorPaths) {
        if (Add-FailedFile -localBasePath $localPath -filePath $absPath -errorMessage "転送前エラーにより転送されませんでした") {
            $failureCount++
        }
    }
    
    # フォルダ全体の結果を判定
    if ($failureCount -eq 0) {
        $script:UploadSuccess += $localPath
        Write-Host "$localPath のアップロードが成功しました。"
    } elseif ($successCount -eq 0) {
        $script:UploadFailed += $localPath
        Write-Host "$localPath のアップロードが失敗しました。"
    } else {
        $script:UploadSuccess += $localPath
        $script:UploadFailed += $localPath
        Write-Host "$localPath のアップロードが部分的に成功しました。"
    }
}

# 単一ファイルアップロードの結果を処理する関数
function Process-SingleFileUploadResult {
    param (
        [string]$localPath,
        [object]$transferResult
    )
    
    <#
    単一ファイル PutFiles() の結果を、最終的に
    - 成功（UploadSuccess）
    - 失敗（UploadFailed）
    のどちらに入れるか判定します。

    注意：WinSCP は「IsSuccess が False」でも部分的に Transfers が成功している場合があります。
    そのため、Transfers と Failures の両方を見て判定しています。
    #>

    $hasErrors = $false
    $hasSuccessfulTransfer = $false
    
    # 実際に転送されたファイルがあるかチェック（Error が null のもの）
    $successfulTransfers = $transferResult.Transfers | Where-Object { $_.Error -eq $null }
    if ($successfulTransfers -and $successfulTransfers.Count -gt 0) {
        $hasSuccessfulTransfer = $true
    }
    
    # TransferResultの失敗をチェック
    if (-not $transferResult.IsSuccess) {
        $hasErrors = $true
        foreach ($failure in $transferResult.Failures) {
            Write-Host "  エラー: $($failure.Message)" -ForegroundColor Red
        }
    }
    
    # Transfersに含まれるエラーもチェック
    $failedTransfers = $transferResult.Transfers | Where-Object { $_.Error -ne $null }
    if ($failedTransfers -and $failedTransfers.Count -gt 0) {
        $hasErrors = $true
        foreach ($transfer in $failedTransfers) {
            Write-Host "  エラー: $($transfer.Error.Message)" -ForegroundColor Red
        }
    }
    
    # QueryReceived / FileTransferProgress で拾った失敗をチェック（取得とクリアを共通化）
    $batchErrorPaths = Get-AndClearBatchErrorFilePaths
    if ($batchErrorPaths -contains $localPath) {
        $hasErrors = $true
        Write-Host "  エラー: 転送前エラーにより転送されませんでした" -ForegroundColor Red
    }
    
    # 転送されたファイルが1つもない場合もエラーとみなす
    if ($transferResult.Transfers.Count -eq 0) {
        $hasErrors = $true
        Write-Host "  エラー: ファイルが転送されませんでした" -ForegroundColor Red
    }
    
    # 最終的な成功/失敗の判定
    # エラーがあり、かつ成功した転送がない場合は完全失敗
    if ($hasErrors -and -not $hasSuccessfulTransfer) {
        $script:UploadFailed += $localPath
        Write-Host "$localPath のアップロードが失敗しました。" -ForegroundColor Red
    } elseif ($hasErrors -and $hasSuccessfulTransfer) {
        # エラーはあるが部分的に成功している場合（通常は単一ファイルではありえないが念のため）
        $script:UploadFailed += $localPath
        Write-Host "$localPath のアップロードが失敗しました（部分的な転送がありました）。" -ForegroundColor Red
    } else {
        $script:UploadSuccess += $localPath
        Write-Host "$localPath のアップロードが成功しました。"
    }
}

#endregion アップロード

#region ユーティリティ（上の処理が内部で使う部品）

function Get-CommandErrorDetail {
    <#
    ExecuteCommand の結果から ErrorOutput / Output を確認し、エラー原因文字列を返します。
    両方空の場合は「原因不明」を返します。
    #>
    param([Parameter(Mandatory)][object]$result)

    $errorDetail = if ($result.ErrorOutput) {
        $result.ErrorOutput.Trim()
    } elseif ($result.Output) {
        $result.Output.Trim()
    } else {
        "原因不明"
    }
    return $errorDetail
}

# ファイル・フォルダの権限を設定する関数
# ファイル: 664 (rw-rw-r--), フォルダ: 775 (rwxrwxr-x)
function Set-RemoteFilePermissions {
    param (
        [string]$remotePath,
        [WinSCP.Session]$session,
        [string]$filePermission = "664",
        [string]$folderPermission = "775"
    )
    
    try {
        # リモートパスが「ファイル」か「フォルダ」かを判定します
        $fileInfo = $session.GetFileInfo($remotePath)

        if ($fileInfo.IsDirectory) {
            # -----------------------------
            # フォルダの場合：chmod → 中身も再帰
            # -----------------------------
            Write-Host "フォルダ権限を設定中: $remotePath ($folderPermission)"

            $chmodCommand = "SITE CHMOD $folderPermission $remotePath"
            $result = $session.ExecuteCommand($chmodCommand)

            if ($result.IsSuccess) {
                Write-Host "  ✓ フォルダ権限設定完了: $remotePath" -ForegroundColor Green
            } else {
                $errorDetail = Get-CommandErrorDetail -result $result
                Write-Host "  ✗ フォルダ権限設定失敗: $remotePath - $errorDetail" -ForegroundColor Yellow
            }

            # フォルダ内のファイル・サブフォルダの権限も設定（失敗しても処理は継続）
            try {
                $remoteFiles = $session.ListDirectory($remotePath)
                foreach ($file in $remoteFiles.Files) {
                    if ($file.Name -eq "." -or $file.Name -eq "..") { continue }
                    $childPath = ("$remotePath/$($file.Name)" -replace "//", "/")
                    try {
                        Set-RemoteFilePermissions -remotePath $childPath -session $session -filePermission $filePermission -folderPermission $folderPermission
                    } catch {
                        Write-Warn "  警告: 権限設定に失敗（処理継続）: $childPath - $_"
                    }
                }
            } catch {
                Write-Warn "  警告: サブフォルダの権限設定でエラー: $_"
            }
        } else {
            # -----------------------------
            # ファイルの場合：chmod のみ
            # -----------------------------
            Write-Host "ファイル権限を設定中: $remotePath ($filePermission)"

            $chmodCommand = "SITE CHMOD $filePermission $remotePath"
            $result = $session.ExecuteCommand($chmodCommand)

            if ($result.IsSuccess) {
                Write-Host "  ✓ ファイル権限設定完了: $remotePath" -ForegroundColor Green
            } else {
                $errorDetail = Get-CommandErrorDetail -result $result
                Write-Host "  ✗ ファイル権限設定失敗: $remotePath - $errorDetail" -ForegroundColor Yellow
            }
        }
    } catch {
        Write-Warn "  ✗ 権限設定エラー: $remotePath - $_"
    }
}

function New-TransferOptions {
    <#
    WinSCP の転送オプションを生成します。
    現状はバイナリ転送固定（テキスト/改行変換を避けるため）です。
    #>
    $transferOptions = New-Object WinSCP.TransferOptions
    $transferOptions.TransferMode = [WinSCP.TransferMode]::Binary
    return $transferOptions
}

function Move-SessionToRootDirectory {
    <#
    WinSCP セッションのカレントディレクトリを / に移動します。
    相対パス解釈の事故を防ぐため、処理の前に毎回呼ぶ方針です。
    #>
    param(
        [Parameter(Mandatory)]
        [WinSCP.Session]$session
    )
    $null = $session.ExecuteCommand("CWD /")
}

# ファイルパスから相対パスを抽出する共通関数
function Get-RelativePath {
    param (
        [string]$fullPath,
        [string]$basePath
    )
    
    if ($fullPath -and $fullPath.StartsWith($basePath, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $fullPath.Substring($basePath.Length).TrimStart([char[]]"\/")
    } else {
        return Split-Path $fullPath -Leaf
    }
}

# エラーメッセージからファイルパスを抽出する関数
function Get-FilePathFromErrorMessage {
    param (
        [string]$errorMessage
    )
    
    if ($errorMessage -match "(?i)(?:file|file or folder) '([^']+)'") {
        return $matches[1]
    }
    return $null
}

function Read-YesNo {
    <#
    yes/no の入力を安全に受け取るための共通関数です。
    戻り値: $true (yes), $false (no)
    #>
    param(
        [Parameter(Mandatory)]
        [string]$message
    )

    do {
        Write-Host $message -ForegroundColor Yellow -NoNewline
        $userResponse = $Host.UI.ReadLine().ToLower()

        if ($userResponse -eq "yes" -or $userResponse -eq "y") { return $true }
        if ($userResponse -eq "no" -or $userResponse -eq "n") { return $false }

        Write-Host "無効な入力です。'yes' または 'no' を入力してください。" -ForegroundColor Red
    } while ($true)
}

function ConvertTo-UnixPath {
    <#
    Windows の \ を / に変換し、// のような重複を潰します。
    #>
    param([Parameter(Mandatory)][string]$path)
    return (($path -replace "\\", "/") -replace "//+", "/")
}

function Get-RemoteParentDirectory {
    <#
    リモートファイルの「親フォルダ」を取得します（Unixスラッシュで返す）。
    例: /common/css/style.css → /common/css
    #>
    param([Parameter(Mandatory)][string]$remotePath)

    $parent = ConvertTo-UnixPath -path (Split-Path $remotePath -Parent)
    if (-not $parent.StartsWith('/')) { $parent = '/' + $parent }
    return $parent
}

function Get-AndClearBatchErrorFilePaths {
    <#
    main_process.ps1 側のイベントで収集している「転送前に弾かれたパス」を取得してクリアします。
    - 戻り値: パス文字列の配列（無ければ空配列）
    #>
    if (-not $global:BatchErrorFiles) { return @() }
    $paths = @($global:BatchErrorFiles.Keys)
    $global:BatchErrorFiles = @{}
    return $paths
}

function Get-UploadedRemotePath {
    <#
    アップロード先のリモートパス（chmod 対象）を計算します。
    PutFiles() の仕様上、ファイルでもフォルダでも「末尾要素（Leaf）」を childPath 配下に作るため、
    単純に `$normalizedChildPath/$leaf` で統一できます。
    ルート直下へのアップロード時は $normalizedChildPath が空文字列になるため AllowEmptyString を付与。
    #>
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$normalizedChildPath,
        [Parameter(Mandatory)][string]$localPath
    )
    $leaf = Split-Path $localPath -Leaf
    return ("$normalizedChildPath/$leaf" -replace "//+", "/")
}

# 失敗ファイルを記録する共通関数
function Add-FailedFile {
    param (
        [string]$localBasePath,
        [string]$filePath,
        [string]$errorMessage = "転送失敗"
    )
    
    $relativePath = Get-RelativePath -fullPath $filePath -basePath $localBasePath
    
    if (-not $script:UploadFailedFiles.ContainsKey($localBasePath)) {
        $script:UploadFailedFiles[$localBasePath] = @()
    }
    
    # 実行回数ベースのカウント：重複を許可して常に追加
    $script:UploadFailedFiles[$localBasePath] += $relativePath
    Write-Host "    ✗ ${relativePath}: $errorMessage" -ForegroundColor Red
    return $true  
}

# 成功ファイルを記録する共通関数
function Add-SuccessFile {
    param (
        [string]$localBasePath,
        [string]$filePath
    )
    
    $relativePath = Get-RelativePath -fullPath $filePath -basePath $localBasePath
    
    if (-not $script:UploadSuccessFiles.ContainsKey($localBasePath)) {
        $script:UploadSuccessFiles[$localBasePath] = @()
    }
    
    # 実行回数ベースのカウント
    $script:UploadSuccessFiles[$localBasePath] += $relativePath
    Write-Host "    ✓ $relativePath" -ForegroundColor Green
    return $true
}

#endregion ユーティリティ
