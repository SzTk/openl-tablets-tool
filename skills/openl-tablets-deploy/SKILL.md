---
name: openl-tablets-deploy
description: openl-tablets-create / openl-tablets-edit で作成・編集した Excel を Azure 上の OpenL Tablets インスタンスにデプロイし、REST API として動作確認する。
license: MIT
metadata:
  argument-hint: "<excel_file_path>"
allowed-tools: Bash Read
---

# OpenL Tablets クラウドデプロイ・スキル

作成・編集した Excel ルールを Azure 上の OpenL Tablets (`openltablets/ws`) にデプロイし、
REST API として正しく動作することを確認する。

## 前提

deploy-service と OpenL Tablets が Azure Container Instances にデプロイ済みであること
（`deploy/azure/README.md` 参照）。

## ステップ

### Step 1: 設定の確認

`~/.config/openl-tablets-tool/deploy.env` を確認する。

存在しない場合、ユーザーに以下を確認して保存する:

```bash
mkdir -p ~/.config/openl-tablets-tool
cat > ~/.config/openl-tablets-tool/deploy.env << 'EOF'
OPENL_DEPLOY_SERVICE_URL=http://<azure-ip>:8000
RESOURCE_GROUP=openl-demo-rg
CONTAINER_GROUP=openl-demo
EOF
```

- `OPENL_DEPLOY_SERVICE_URL`: deploy-service のベース URL
- `RESOURCE_GROUP` / `CONTAINER_GROUP`: `deploy/azure/start.sh` 用（デフォルト値で良ければそのまま）

### Step 2: サーバー起動確認

```bash
source ~/.config/openl-tablets-tool/deploy.env
curl -sf "$OPENL_DEPLOY_SERVICE_URL/health"
```

応答がない場合、ACI を起動する:

```bash
RESOURCE_GROUP="$RESOURCE_GROUP" CONTAINER_GROUP="$CONTAINER_GROUP" \
  ./deploy/azure/start.sh
```

`/health` が 200 を返すまで、5 秒間隔で最大 120 秒リトライする。
それでも応答しない場合はユーザーにエラーを報告して停止する。

### Step 3: 対象ファイルとサービス名の確認

引数 `$ARGUMENTS` が指定されていればそれを対象の Excel ファイルとする。
指定がなければ、カレントディレクトリの `.xlsx` を一覧してユーザーに選択を求める。

ファイル名からサービス名を提案する（例: `ShopPolicy.xlsx` → `shop-policy`）。
ユーザーに確認・変更の機会を与える。

### Step 4: デプロイ

```bash
curl -X POST "$OPENL_DEPLOY_SERVICE_URL/deploy" \
  -F "file=@<対象Excelの絶対パス>" \
  -F "service_name=<サービス名>"
```

レスポンスの `endpoint` と `swagger_url` を保持する。
エラー（400 / 504 など）が返った場合は内容をそのままユーザーに表示する。

### Step 5: API テスト

```bash
curl "<swagger_url>"
```

返ってきた OpenAPI 定義から、利用可能なメソッドとパラメータ型を確認する。
代表的な 1 メソッドについて、パラメータ型に合うサンプル値を組み立てて呼び出す:

```bash
curl -X POST "<endpoint>/<MethodName>" \
  -H "Content-Type: application/json" \
  -d '<サンプルJSON>'
```

レスポンスをユーザーに表示する。

### Step 6: 完了報告

```
✅ デプロイ完了

サービス名: <service_name>
エンドポイント: <endpoint>
Swagger: <swagger_url>

テスト結果:
<curl の出力>

不要になったら停止できます:
RESOURCE_GROUP=<RESOURCE_GROUP> CONTAINER_GROUP=<CONTAINER_GROUP> ./deploy/azure/stop.sh
```

## 注意事項

- `deploy.env` に保存した URL ・リソースグループ名は、ユーザー固有の Azure 環境を前提とする。
  他のユーザーが利用する場合は各自で設定する。
- Excel ファイルは `.xlsx` のみ対応。`.xls` 等は事前に変換が必要。
