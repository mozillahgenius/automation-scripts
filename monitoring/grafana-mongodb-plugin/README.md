# Grafana + MongoDB Plugin (Docker)

## 概要

Grafana + Open MongoDB Plugin の Docker 環境構成。
MongoDB データソースを Grafana で可視化するための設定一式。

## 構成

- Grafana 8.4.2 (Docker)
- Open MongoDB Grafana Plugin (MongoDB Proxy 経由)
- Docker Compose による環境構築

## 起動方法

```bash
docker-compose up -d
```

Grafana UI: http://localhost:3000

## セットアップ

1. `docker-compose.yml` の `GF_SECURITY_ADMIN_PASSWORD` を設定
2. `grafana_data/plugins/open-mongodb-grafana/server/config/default.json` に MongoDB 接続情報を設定
3. MongoDB Proxy を起動: `cd grafana_data/plugins/open-mongodb-grafana/server && node mongodb-proxy.js`
4. Grafana で MongoDB データソースを追加（Proxy URL: `http://localhost:3333`）
