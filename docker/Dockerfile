# ローカル開発用 vsftpd イメージ
#
# fauria/vsftpd は「仮想ユーザ + 環境変数」で設定する設計のため、conf を丸ごと
# bind mount / 置換すると以下の問題が起きる:
#   - ホスト所有(非root)の conf は vsftpd に拒否される (500 OOPS: config file not owned ...)
#   - :ro だとエントリポイントの追記が Read-only エラーになる
#   - background=NO や仮想ユーザ設定(guest_enable 等)が抜けて即終了する
# そこで標準 conf は活かし、不足している SSL 設定だけを追記する。
#
# 接続側(WinSCP)は Explicit FTPS で接続するため自己署名証明書も用意する。

# 自己署名証明書をビルド時に生成 (元イメージに openssl が無いため別ステージで作成)
FROM alpine:3 AS certs
RUN apk add --no-cache openssl \
    && openssl req -x509 -nodes -days 3650 -newkey rsa:2048 \
        -keyout /vsftpd.pem -out /vsftpd.pem \
        -subj "/CN=localhost"

FROM fauria/vsftpd

# 標準 conf に SSL 設定を追記
COPY vsftpd-ssl.conf /tmp/vsftpd-ssl.conf
RUN cat /tmp/vsftpd-ssl.conf >> /etc/vsftpd/vsftpd.conf && rm /tmp/vsftpd-ssl.conf

# 証明書を配置 (COPY により root 所有になる)
COPY --from=certs /vsftpd.pem /etc/vsftpd/vsftpd.pem
RUN chmod 600 /etc/vsftpd/vsftpd.pem
