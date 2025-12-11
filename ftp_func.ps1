# ================================
# WinSCP.NET FTPアップロード自動化
# ================================

# このファイルは共通処理ライブラリとして使用されます
# 直接実行する場合は main_process.ps1 を使用してください
if ((Split-Path $MyInvocation.InvocationName -Leaf) -eq $MyInvocation.MyCommand.Name) {
    Write-Host "このファイルは共通処理ライブラリです。" -ForegroundColor Yellow
    Write-Host "FTP処理を実行するには main_process.ps1 を使用してください。" -ForegroundColor Yellow
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    exit 0
}

# 初期化エラーフラグ
$script:InitializationError = $false

# 設定ファイルの読み込み
$configPath = "$PSScriptRoot\config.ps1"
if (-not (Test-Path $configPath)) {
    Write-Host "設定ファイルが見つかりません: $configPath" -ForegroundColor Red
    Write-Host "エクセルマクロでconfig.ps1を作成してください。"
    $script:InitializationError = $true
    return
}
. $configPath

# WinSCP.NET.dllの読み込み
if (-not (Test-Path $winscpDllPath)) {
    Write-Host "WinSCP.NET.dllが見つかりません: $winscpDllPath" -ForegroundColor Red
    Write-Host "WinSCP.NET.dllをダウンロードしてスクリプトと同じフォルダに配置してください。"
    $script:InitializationError = $true
    return
}

try {
    Add-Type -Path $winscpDllPath
} catch {
    Write-Host "WinSCP.NET.dllの読み込みに失敗しました: $_" -ForegroundColor Red
    $script:InitializationError = $true
    return
}

# ログファイルのパスを設定
$LOG_PATH = "$PSScriptRoot\WinSCP.log"
# バックアップフォルダのパスを設定（設定ファイルから読み込み）
$BACKUP_PATH = $backupPath

# ========================================
# グローバル変数の初期化
# ========================================

# 処理結果集計用の配列
$script:BackupSuccess   = @()
$script:BackupFailed    = @()
$script:BackupNotNeeded = @()
$script:UploadSuccess   = @()
$script:UploadFailed    = @()

# 個別ファイルの結果を記録するための辞書
$script:UploadSuccessFiles = @{}
$script:UploadFailedFiles = @{}

# バッチエラー（QueryReceived）で拾ったローカル側の失敗ファイルの絶対パスを一時保持
$global:BatchErrorFiles = @{}

# ========================================
# ユーティリティ関数
# ========================================

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
        # リモートパスの情報を取得
        $fileInfo = $session.GetFileInfo($remotePath)
        
        if ($fileInfo.IsDirectory) {
            # フォルダの場合
            Write-Host "フォルダ権限を設定中: $remotePath ($folderPermission)"
            $chmodCommand = "SITE CHMOD $folderPermission $remotePath"
            $result = $session.ExecuteCommand($chmodCommand)
            
            if ($result.IsSuccess) {
                Write-Host "  ✓ フォルダ権限設定完了: $remotePath" -ForegroundColor Green
            } else {
                Write-Host "  ✗ フォルダ権限設定失敗: $remotePath - $($result.ErrorOutput)" -ForegroundColor Yellow
            }
            
            # フォルダ内のファイル・サブフォルダの権限も再帰的に設定
            try {
                $remoteFiles = $session.ListDirectory($remotePath)
                foreach ($file in $remoteFiles.Files) {
                    if ($file.Name -ne "." -and $file.Name -ne "..") {
                        $childPath = "$remotePath/$($file.Name)" -replace "//", "/"
                        Set-RemoteFilePermissions -remotePath $childPath -session $session -filePermission $filePermission -folderPermission $folderPermission
                    }
                }
            } catch {
                Write-Host "  警告: サブフォルダの権限設定でエラー: $_" -ForegroundColor Yellow
            }
        } else {
            # ファイルの場合
            Write-Host "ファイル権限を設定中: $remotePath ($filePermission)"
            $chmodCommand = "SITE CHMOD $filePermission $remotePath"
            $result = $session.ExecuteCommand($chmodCommand)
            
            if ($result.IsSuccess) {
                Write-Host "  ✓ ファイル権限設定完了: $remotePath" -ForegroundColor Green
            } else {
                Write-Host "  ✗ ファイル権限設定失敗: $remotePath - $($result.ErrorOutput)" -ForegroundColor Yellow
            }
        }
    } catch {
        Write-Host "  ✗ 権限設定エラー: $remotePath - $_" -ForegroundColor Yellow
    }
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

# ========================================
# WinSCPセッションオプションの初期化
# ========================================

function Initialize-WinSCPSession {
    Write-Host "スクリプト実行開始: $(Get-Date)"
    
    # WinSCPセッションオプションの設定
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

# セッションオプションを初期化
$sessionOptions = Initialize-WinSCPSession

#関数を宣言
function action {
    param (
        [string]$remotePath,
        [string]$localPath,
        [bool]$deleteAfterBackup = $false,
        [WinSCP.Session]$session
    )

    # リモートパスのディレクトリ部分（Unix スラッシュ維持）を取得
    $childPath = (
        (Split-Path $remotePath -Parent) -replace "\\", "/"
    )
    if (-not $childPath.StartsWith('/')) {
        # 先頭にスラッシュが無い場合は追加（絶対パス化）
        $childPath = '/' + $childPath
    }
    
    Backup-RemoteFile $remotePath $childPath $deleteAfterBackup $session
    Upload-LocalFile $localPath $childPath $session
}

function Backup-RemoteFile {
    param (
        [string]$remotePath,
        [string]$childPath,
        [bool]$deleteAfterBackup = $false,
        [WinSCP.Session]$session
    )
    # 明示的にルートディレクトリに移動
    $null = $session.ExecuteCommand("CWD /")
    Write-Host "リモートのファイルの存在を確認: $remotePath"
    
    try {
        $fileInfo = $session.GetFileInfo($remotePath)
        Write-Host "リモートのファイルのバックアップを取得: $remotePath"
        
        $backupDir = Join-Path $BACKUP_PATH $childPath
        if (-not (Test-Path $backupDir)) {
            New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
        }
        
        $transferOptions = New-Object WinSCP.TransferOptions
        $transferOptions.TransferMode = [WinSCP.TransferMode]::Binary
        
        $transferResult = $session.GetFiles($remotePath, "$backupDir\", $False, $transferOptions)
        
        if ($transferResult.IsSuccess) {
            $script:BackupSuccess += $remotePath
            Write-Host "バックアップ成功: $remotePath"

            # バックアップ成功後、削除フラグがtrueの場合はリモートファイル/フォルダを削除
            if ($deleteAfterBackup) {
                try {
                    # 削除実行の確認
                    do {
                        Write-Host "バックアップ対象 '$remotePath' を削除しますか？ (yes/no): " -ForegroundColor Yellow -NoNewline
                        $userResponse = $Host.UI.ReadLine().ToLower()

                        if ($userResponse -eq "yes" -or $userResponse -eq "y") {
                            Write-Host "削除を実行します。" -ForegroundColor Green
                            break
                        }
                        elseif ($userResponse -eq "no" -or $userResponse -eq "n") {
                            Write-Host "削除をスキップします。" -ForegroundColor Yellow
                            Write-Host "バックアップは完了しましたが、削除は実行されませんでした。" -ForegroundColor Cyan
                            return
                        }
                        else {
                            Write-Host "無効な入力です。'yes'または'no'を入力してください。" -ForegroundColor Red
                        }
                    } while ($true)

                    Write-Host "バックアップ対象を削除: $remotePath"

                    # リモートパスがファイルかフォルダかを判定
                    $fileInfo = $session.GetFileInfo($remotePath)

                    # ファイルとフォルダの両方にRemoveFilesを使用（より信頼性が高い）
                    Write-Host "対象を削除: $remotePath"
                    try {
                        $removalResult = $session.RemoveFiles($remotePath)

                        if ($removalResult.IsSuccess) {
                            if ($fileInfo.IsDirectory) {
                                Write-Host "フォルダ削除成功: $remotePath" -ForegroundColor Yellow
                            } else {
                                Write-Host "ファイル削除成功: $remotePath" -ForegroundColor Yellow
                            }

                            # 削除結果の検証（オプション）
                            try {
                                $verifyInfo = $session.GetFileInfo($remotePath)
                                Write-Host "警告: 削除後もファイル/フォルダがまだ存在しています" -ForegroundColor Yellow
                            } catch [WinSCP.SessionRemoteException] {
                                # SessionRemoteExceptionはファイルが存在しないことを意味するので成功
                                Write-Host "検証: 正常に削除されました" -ForegroundColor Green
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

                            # WinSCPのエラーメッセージをそのまま使用して例外を投げる
                            throw $winscpErrorMessage
                        }
                    } catch [WinSCP.SessionRemoteException] {
                        # ファイルが存在しない場合は成功として扱う
                        Write-Host "対象が存在しないため削除をスキップ: $remotePath" -ForegroundColor Yellow
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
    } catch {
        $script:BackupFailed += $remotePath
        # バックアップ処理中のエラーでも処理を中断
        throw "BackupError:$remotePath"
    }
}

# フォルダアップロードの結果を処理する関数
function Process-DirectoryUploadResult {
    param (
        [string]$localPath,
        [object]$transferResult
    )
    
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

    # QueryReceived で拾った失敗を補完
    if ($global:BatchErrorFiles.Keys.Count -gt 0) {
        foreach ($absPath in $global:BatchErrorFiles.Keys) {
            if (Add-FailedFile -localBasePath $localPath -filePath $absPath -errorMessage "転送前エラーにより転送されませんでした") {
                $failureCount++
            }
        }
        $global:BatchErrorFiles = @{}
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
    
    # QueryReceived で拾った失敗をチェック
    if ($global:BatchErrorFiles.Keys.Count -gt 0) {
        foreach ($absPath in $global:BatchErrorFiles.Keys) {
            # 現在のlocalPathと一致するかチェック
            if ($absPath -eq $localPath) {
                $hasErrors = $true
                Write-Host "  エラー: 転送前エラーにより転送されませんでした" -ForegroundColor Red
            }
        }
        # バッチエラーをクリア
        $global:BatchErrorFiles = @{}
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

function Upload-LocalFile {
    param (
        [string]$localPath,
        [string]$childPath,
        [WinSCP.Session]$session
    )
    
    Write-Host "ローカルのファイルをアップロード: $localPath"
    
    try {
        # 明示的にルートディレクトリに移動
        $null = $session.ExecuteCommand("CWD /")
        $transferOptions = New-Object WinSCP.TransferOptions
        $transferOptions.TransferMode = [WinSCP.TransferMode]::Binary
        
        # リモートパスの末尾スラッシュを正規化（重複スラッシュ防止）
        $normalizedChildPath = ($childPath -replace '/+$','')
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
            
            # アップロード先のリモートパスを計算
            $uploadedRemotePath = if ($isDirectory) {
                # ディレクトリの場合、パス名を追加
                $localDirName = Split-Path $localPath -Leaf
                "$normalizedChildPath/$localDirName" -replace "//", "/"
            } else {
                # ファイルの場合、ファイル名を追加
                $localFileName = Split-Path $localPath -Leaf
                "$normalizedChildPath/$localFileName" -replace "//", "/"
            }
            
            # 権限設定を実行（エラーがあっても処理を継続）
            try {
                Set-RemoteFilePermissions -remotePath $uploadedRemotePath -session $session
            } catch {
                Write-Host "  警告: 権限設定でエラーが発生しましたが、処理を継続します: $_" -ForegroundColor Yellow
            }
        } else {
            Write-Host "アップロードが完全に失敗したため、権限設定をスキップします。" -ForegroundColor Yellow
        }
    } catch {
        Write-Host "アップロード処理中にエラーが発生しました: $_" -ForegroundColor Red
        $script:UploadFailed += $localPath
    }
}


