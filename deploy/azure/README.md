# Azure へのデプロイ (Container Instances)

OpenL Tablets (`openltablets/ws`) と deploy-service を同一の Azure Container Instance (ACI) コンテナグループにデプロイする。

## 前提

- Azure CLI (`az`) がインストール済みで `az login` 済みであること
- deploy-service イメージをビルドし、Azure 上から取得可能なレジストリに push 済みであること

```bash
docker build -t <registry>/openl-deploy-service:latest ../../deploy-service
docker push <registry>/openl-deploy-service:latest
```

## 初回デプロイ

```bash
export DEPLOY_SERVICE_IMAGE=<registry>/openl-deploy-service:latest
export RESOURCE_GROUP=openl-demo-rg     # 任意
export LOCATION=japaneast               # 任意
./deploy.sh
```

実行後に表示される IP アドレスを使い、`OPENL_PUBLIC_URL` を `http://<ip>:8080` として
deploy-service の環境変数を更新する場合は、コンテナグループを再作成するか
`az container create` を再実行する（ACI は稼働中コンテナの環境変数を直接更新できない）。

## 起動 / 停止（コスト節約）

```bash
./start.sh   # 再起動。IP アドレスが表示される
./stop.sh    # 停止（課金を止める）
```

## エンドポイント

- OpenL Tablets REST API: `http://<ip>:8080/<service-name>`
- Deploy Service: `http://<ip>:8000`
