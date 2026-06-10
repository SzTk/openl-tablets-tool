# OpenL Tablets Cloud Deploy Service — Design Spec

**Date:** 2026-06-09  
**Status:** Approved  

## Context

`openl-tablets-create` および `openl-tablets-edit` スキルで Excel ルール定義を作成・編集できるようになった。しかし作成した Excel が実際に OpenL Tablets API として動作する様子を、スキル利用者がローカル Docker なしで確認できない。

`how-to-deploy-excel-to-api.md` に手動手順はあるが、スキルには組み込まれておらず、エンドツーエンドの確認体験が欠けている。

**目的:** Excel 作成/編集後に、Azure 上の OpenL Tablets インスタンスへ自動デプロイし、REST API テストまでスキル内で完結させる。

---

## アーキテクチャ

```
ユーザー (スキル利用者)
  │
  │ openl-tablets-create / openl-tablets-edit スキルで Excel 生成済み
  │
  ▼
openl-tablets-deploy スキル (Claude Code)
  │
  │ 1. ACI サーバー起動確認 → 停止中なら az container start
  │ 2. POST /deploy に Excel をアップロード
  │ 3. エンドポイント URL + curl サンプルを受け取る
  │ 4. curl でテスト呼び出し → 結果をユーザーに表示
  │
  ▼
Deploy Service (FastAPI) — Azure Container Instance
  │
  │ a. Excel ファイルを受信
  │ b. rules.xml + rules-deploy.xml を生成
  │ c. /deployment/{service_name}/ に配置
  │
  ▼
OpenL Tablets (openltablets/ws) — 同一 ACI コンテナグループ
  │
  │ deployment/ を監視 → 自動デプロイ (repo-file モード)
  │
  ▼
REST API: http://<azure-ip>:9080/REST/{service_name}/...
```

**永続化:** Azure Files (Storage Account + File Share) を `/deployment` にマウントし、ACI 再起動後もルールを保持する。

---

## コンポーネント詳細

### 1. Deploy Service (`deploy-service/`)

**技術スタック:** Python 3.12+, FastAPI, uvicorn, python-multipart

**ディレクトリ構成:**
```
deploy-service/
├── main.py          # FastAPI アプリ
├── requirements.txt
└── Dockerfile
```

**API エンドポイント:**

| Method | Path | 説明 |
|--------|------|------|
| `POST` | `/deploy` | Excel アップロード → デプロイ |
| `GET` | `/services` | デプロイ済みサービス一覧 |
| `DELETE` | `/services/{name}` | アンデプロイ |
| `GET` | `/health` | ヘルスチェック |

**`POST /deploy` リクエスト:**
```
multipart/form-data:
  file: <.xlsx ファイル>
  service_name: "shop-policy"   (省略時はファイル名 (小文字・ハイフン区切り) から自動生成)
```

**`POST /deploy` レスポンス:**
```json
{
  "service_name": "shop-policy",
  "endpoint": "http://<azure-ip>:9080/REST/shop-policy",
  "swagger_url": "http://<azure-ip>:9080/REST/shop-policy/api-docs"
}
```

> **前提:** OpenL Tablets (`repo-file` モード) はデプロイフォルダへのファイル追加を自動検知して再デプロイする。この挙動は初回セットアップ時に `docker compose up` でローカル検証すること。

**内部処理 (`POST /deploy`):**
1. Excel ファイルを受信・保存
2. `rules.xml` + `rules-deploy.xml` を `how-to-deploy-excel-to-api.md` のテンプレートで生成
3. `$DEPLOYMENT_PATH/{service_name}/` ディレクトリを作成し配置
4. OpenL Tablets の REST API (`/REST/{service_name}`) が応答するまでポーリング (最大 60 秒)
5. エンドポイント URL と Swagger URL をレスポンスとして返却（curl サンプルはスキルが Swagger を参照して生成）

**環境変数:**
- `DEPLOYMENT_PATH` — OpenL が監視するデプロイフォルダのパス (例: `/deployment`)
- `OPENL_BASE_URL` — OpenL Tablets の内部 URL (例: `http://openl:8080`)

---

### 2. Docker 構成 (`docker-compose.yml`)

```yaml
services:
  openl:
    image: openltablets/ws
    ports: ["9080:8080"]
    environment:
      - production-repository.factory=repo-file
      - production-repository.uri=/tmp/openl
    volumes:
      - deployment:/tmp/openl/deployment

  deploy-service:
    build: ./deploy-service
    ports: ["8000:8000"]
    volumes:
      - deployment:/deployment
    environment:
      - DEPLOYMENT_PATH=/deployment
      - OPENL_BASE_URL=http://openl:8080
    depends_on:
      - openl

volumes:
  deployment:
```

---

### 3. Azure 構成 (`deploy/azure/`)

**ディレクトリ構成:**
```
deploy/azure/
├── deploy.sh     # 初回: ACI 作成 + Azure Files 設定
├── start.sh      # az container start（再起動・コスト節約後の起動）
├── stop.sh       # az container stop（使用後に停止）
└── README.md     # セットアップ手順
```

**ACI 構成:**
- コンテナグループに `openl` + `deploy-service` の 2 コンテナを収容
- Azure Files を `/deployment` にマウント（永続化）
- ポート 9080 (OpenL REST API) と 8000 (Deploy Service) を公開

**ライフサイクル管理:**
- デプロイ前: スキルが `az container show` でサーバー状態確認 → 停止中なら `start.sh` を実行
- 使用後: `stop.sh` でコスト節約 (ユーザーが手動実行 or スキルが案内)

---

### 4. スキル `openl-tablets-deploy` (`skills/openl-tablets-deploy/`)

**目的:** `openl-tablets-create` / `openl-tablets-edit` 後に呼び出し、Azure へのデプロイと API テストをガイドする。

**Azure URL の管理:** `config.yaml` は使用しない。スキル初回実行時に Deploy Service の URL (例: `http://<azure-ip>:8000`) をユーザーに確認し、Claude のメモリに保存する。

**フロー:**
```
1. 設定確認
   Deploy Service URL をメモリから取得、なければユーザーに確認してメモリに保存

2. サーバー状態確認
   GET /health → 応答なしなら az container start を実行
   ヘルスチェック通過まで待機（最大 120 秒）

3. デプロイ対象の確認
   ローカルの .xlsx ファイルを確認
   サービス名をファイル名から提案 → ユーザーが承認または変更

4. デプロイ実行
   POST /deploy に Excel をアップロード
   成功レスポンスからエンドポイント URL を取得

5. API テスト
   Swagger URL から利用可能なエンドポイントを確認
   代表的な curl コマンドを実行してレスポンスを表示

6. 完了報告
   エンドポイント URL と curl サンプルをユーザーに提示
   「不要になったら deploy/azure/stop.sh で停止してください」と案内
```

---

## ファイル構成（追加分のみ）

```
openl-tablets-tool/
├── deploy-service/              # 新規
│   ├── main.py
│   ├── requirements.txt
│   └── Dockerfile
├── deploy/
│   └── azure/                   # 新規
│       ├── deploy.sh
│       ├── start.sh
│       ├── stop.sh
│       └── README.md
├── docker-compose.yml           # 新規
└── skills/
    └── openl-tablets-deploy/    # 新規
        └── SKILL.md
```

---

## 検証方法

1. **ローカル確認:** `docker compose up` → `POST http://localhost:8000/deploy` で Excel を送信 → `http://localhost:9080/REST/{service}` が応答することを確認
2. **Azure 確認:** `deploy.sh` 実行 → `openl-tablets-deploy` スキルで E2E フローを実行 → curl でルール実行結果を確認
3. **停止確認:** `stop.sh` 実行後、`start.sh` で再起動しルールが保持されていることを確認（Azure Files 永続化の検証）
