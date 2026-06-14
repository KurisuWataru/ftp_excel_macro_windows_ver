# FTP Excel Macro Windows

Excelで管理されたファイルリストに基づいて、WinSCPを使用したFTP転送を自動化するツールです。
PowerShellスクリプトを生成し、安全かつ確実にファイルをアップロードします。

## 特徴

* **Excelでの簡単管理**: 転送対象のファイルやフォルダ、FTP接続情報をExcelシートで管理できます。
* **自動バックアップ**: アップロード前にリモートサーバー上の既存ファイルのバックアップを自動取得します。
* **権限設定**: アップロード後のファイル・フォルダのパーミッション（chmod）を自動設定します。
* **安全性**: 接続先IPアドレスのチェック機能により、誤った環境への接続を防止します。
* **詳細なログ**: 転送の成功・失敗、バックアップ状況などを詳細にログ出力します。

## 動作環境

* Windows 10/11
* Microsoft Excel (マクロが有効な状態)
* PowerShell 5.1 以上
* [WinSCP .NET Assembly](https://winscp.net/eng/docs/library) (`WinSCP.NET.dll`)

## セットアップ

1. このリポジトリをクローンするか、GitHub の Releases ページから ZIP をダウンロードして展開します。
2. WinSCPの公式サイトから `.NET Assembly / COM Library` パッケージをダウンロードし、展開します。
3. 展開したフォルダ内の `WinSCP.NET.dll` を、このツールのフォルダ（`FTPコマンド作成マクロ_windows版.xlsm`と同じ場所）または任意の場所に配置します。
   * ※ `WinSCP.exe` も同じフォルダにある必要がある場合があります（依存関係による）。

## 使い方

### 1. Excelでの設定

1. `FTPコマンド作成マクロ_windows版.xlsm` を開きます。
2. **「基本設定」シート**:
   * FTPホスト名、ユーザー名、パスワードを入力します。
   * `WinSCP.NET.dll` のフルパスを指定します。
   * バックアップの保存先パスを指定します。
3. **「ファイルパス」シート**:
   * `ローカルのルートディレクトリのパス`: ローカルファイルのベースパス。
   * `アップロード対象のファイル・フォルダの相対パス`: アップロードするファイル名やフォルダ名。
   * `リモートのルートディレクトリのパス`: アップロード先のベースパス。
   * `アップロード先の相対パス`: アップロード先のサブディレクトリなど。
   * `アップロード時削除フラグ`: アップロード（バックアップ）後にリモートファイルを削除してからアップロードする場合は `True` を指定。
   * `処理モード`: リモートのファイル・フォルダを削除のみ行う（アップロードしない）場合は `削除のみ` と入力。空欄の場合は従来通りアップロードが実行されます。

#### 削除のみモード

「処理モード」列に `削除のみ` と入力すると、該当行はアップロードを行わず、リモートファイル/フォルダのバックアップ取得後に削除のみを実行します。

* ローカルパス関連の列（`ローカルのルートディレクトリのパス`、`アップロード対象のファイル・フォルダの相対パス`）は空欄で構いません。
* バックアップは安全のため必ず取得されます。
* 削除前にユーザー確認プロンプトが表示されます。

| 処理モード列の値 | 動作 |
| --- | --- |
| 空欄 / `アップロード` | 従来通りバックアップ + アップロード（`action`） |
| `削除のみ` | バックアップ + 削除のみ（`deleteAction`） |

### 2. 接続先環境の登録（推奨）

誤送信防止のため、接続先のIPアドレスと環境名を `ip_check.ps1` に登録しておくことを推奨します。
未登録のIPアドレスに接続しようとすると、警告が表示されます。

`ip_check.ps1` をテキストエディタで開き、以下のように設定します：

```powershell
$targetEnvironments = @{
    "192.168.1.10" = "開発環境"
    "192.168.1.20" = "本番環境"
}
```

### 3. スクリプトの生成

1. Excelのマクロボタン（または開発タブからマクロ実行）を押して、設定ファイルと実行スクリプトを生成します。
   * `makeConfig`: `config.ps1` を生成します。
   * `main`: `main_process.ps1` を生成します。

### 4. 転送の実行

1. 生成された `main_process.ps1` を右クリックし、「PowerShell で実行」を選択します。
2. コンソール画面が開き、接続先の確認や処理の続行確認が表示されるので、指示に従って操作します。
3. 処理完了後、結果が表示されます。ログファイル (`display_log.txt`, `WinSCP.log`) も生成されます。

## ファイル構成

* `FTPコマンド作成マクロ_windows版.xlsm`: 設定およびスクリプト生成用Excelファイル
* `ftp_func.ps1`: FTP処理の共通関数ライブラリ
* `ip_check.ps1`: 接続先IP確認用設定ファイル
* `config.ps1`: 生成される設定ファイル（パスワードが含まれるため、Git管理外にすることをお勧めします）
* `main_process.ps1`: 生成されるメイン実行スクリプト
* `docker-compose.yml`: ローカル検証用 FTP サーバーの定義
* `docker/Dockerfile`: `fauria/vsftpd` をベースに、SSL 設定の追記と自己署名証明書の生成を行うイメージ定義
* `docker/vsftpd-ssl.conf`: 標準 conf に追記する Explicit FTPS 用の設定

## ローカル開発環境

実 FTP サーバーに接続せず、手元の Windows で FTP 接続・バックアップ・アップロードを確認するための手順です。本番向けの `makeConfig` や `ftp_func.ps1` の挙動は変更しません。

### 前提

* Docker Desktop（または Docker CLI）が利用可能であること
* WinSCP.NET.dll が配置済みであること
* PowerShell 5.1 以上

### 1. ローカル FTP サーバーの起動

リポジトリルートで以下を実行します。Explicit FTPS 用の設定と自己署名証明書を組み込んだイメージをビルドして起動するため、`--build` を付けます（`docker/Dockerfile` を使用）。

```powershell
docker compose up -d --build
```

2 回目以降で `docker/` 配下に変更がなければ `--build` は省略できます。

```powershell
docker compose up -d
```

停止する場合:

```powershell
docker compose down
```

接続情報（`docker-compose.yml` と `config.ps1.example` と整合）:

| 項目 | 値 |
| --- | --- |
| ホスト | `127.0.0.1` |
| ポート | `21`（Explicit FTPS） |
| ユーザー名 | `localdev` |
| パスワード | `localdev-pass` |
| FTP データディレクトリ（ホスト） | `local-data/ftp-root` |
| FTP データディレクトリ（コンテナ内） | `/home/vsftpd/localdev` |
| バックアップ保存先 | `local-data/backup` |

ホストの `local-data/ftp-root` はコンテナ内の `/home/vsftpd/localdev`（FTP ユーザー `localdev` のホーム）にマウントされており、両者は同じ内容を共有します。FTP でアップロードしたファイルはこのディレクトリに格納されます。

`ftp_func.ps1` は `FtpSecure::Explicit` と証明書検証の無効化を使用するため、ローカル FTP は vsftpd + TLS（Explicit FTPS）で起動します。証明書は `docker/Dockerfile` のビルド時に自己署名で生成され、イメージへ組み込まれます（手動準備は不要）。パッシブモード用にポート `21100-21110` を公開し、パッシブ応答アドレスは `127.0.0.1` を返すよう設定しています。

### 2. 設定ファイルの準備

#### `config.ps1`

`config.ps1.example` を `config.ps1` にコピーし、`winscpDllPath` を手元の WinSCP.NET.dll のフルパスに変更します。

```powershell
Copy-Item config.ps1.example config.ps1
```

Excel マクロ（`FTPコマンド作成マクロ_windows版.xlsm`）から `makeConfig` で生成しても構いません。その場合は FTP ホスト・ユーザー・パスワード・バックアップ先をローカル向けの値に変更してください。

**注意:** `$ftpHost` は `127.0.0.1` を使用してください。`ip_check.ps1` の `ContainsKey` はホスト名の完全一致で判定するため、`localhost` とは別扱いになります。

#### `ip_check.ps1`

`ip_check.ps1.example` の内容を `ip_check.ps1` に反映し、ローカルホストを開発環境として登録します。

```powershell
$targetEnvironments = @{
    "127.0.0.1" = "ローカル開発環境"
}
```

### 3. `main_process.ps1` の生成と実行

1. Excel マクロの `main` で `main_process.ps1` を生成します。
2. `main_process.ps1` を PowerShell で実行します。

`config.ps1` は上記の example から手動作成できますが、`main_process.ps1` は Excel マクロの `main` で生成する必要があります。

### 4. 手動確認手順

1. `docker compose up -d --build` でコンテナが起動していることを確認する
2. `config.ps1` と `ip_check.ps1` をローカル向けに準備する
3. `main_process.ps1` を実行し、「接続が成功しました」が表示されることを確認する
4. `local-data/ftp-root` にテスト用ファイルを配置し、Excel「ファイルパス」シートでバックアップ取得を実行する
5. ローカルファイルをアップロードし、`local-data/ftp-root` 上に反映されることを確認する
6. `WinSCP.log` に接続エラーがないことを確認する

ローカル検証で生成された `local-data/backup` および `local-data/ftp-root` 配下のデータは `.gitignore` で除外されます。

### 5. アップロード結果の確認

FTP でアップロードしたファイルは、FTP ユーザー `localdev` のホームディレクトリ（コンテナ内 **`/home/vsftpd/localdev`**）に格納されます。ここはホストの `local-data/ftp-root` にマウントされているため、次のいずれの方法でも確認できます。

* **コンテナ内を直接確認する**

  ```powershell
  docker exec local-ftp-server ls -la /home/vsftpd/localdev
  ```

* **ホスト側（マウント先）を確認する**

  ```powershell
  dir .\local-data\ftp-root\
  ```

* **FTP クライアントで一覧する**

  ```powershell
  curl --ftp-ssl --ssl-reqd -k --user localdev:localdev-pass ftp://127.0.0.1/
  ```

アップロードしたファイルがこれらの一覧に表示されれば成功です。

## 注意事項

* セキュリティのため、`config.ps1` にはパスワードが含まれます。このファイルはバージョン管理システム（Gitなど）にコミットしないように注意してください（`.gitignore` への追加を推奨）。
* 本ツールを使用する際は、事前にテスト環境で十分に動作確認を行ってください。
