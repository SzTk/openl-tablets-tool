# Azure へのデプロイ (Container Instances)

OpenL Tablets (`openltablets/ws`) と deploy-service を同一の Azure Container Instance (ACI) コンテナグループにデプロイする。

## 前提

- Azure CLI (`az`) がインストール済みで `az login` 済みであること
- deploy-service イメージをビルドし、Azure 上から取得可能なレジストリ（例: ACR）に push 済みであること

```bash
docker build -t openl-deploy-service:latest ../../deploy-service

az acr login --name <registry-name>
docker tag openl-deploy-service:latest <login-server>/openl-deploy-service:latest
docker push <login-server>/openl-deploy-service:latest
```

- レジストリの認証情報を取得しておくこと（ACR の場合）

```bash
az acr credential show --name <registry-name>
```

## 初回デプロイ

```bash
export DEPLOY_SERVICE_IMAGE=<login-server>/openl-deploy-service:latest
export REGISTRY_SERVER=<login-server>
export REGISTRY_USERNAME=<acr-username>
export REGISTRY_PASSWORD=<acr-password>
export RESOURCE_GROUP=openl-demo-rg     # 任意
export LOCATION=japaneast               # 任意
export CONTAINER_GROUP=openl-demo       # 任意
./deploy.sh
```

公開 URL は `dnsNameLabel` から `http://${CONTAINER_GROUP}.${LOCATION}.azurecontainer.io` として
事前に決まるため、IP アドレスの確定を待って再作成する必要はない。
`DNS_NAME_LABEL` は Azure リージョン内で一意である必要があるため、衝突した場合は
`DNS_NAME_LABEL` を別の値に設定して再実行する。

## 起動 / 停止（コスト節約）

```bash
./start.sh   # 再起動。IP アドレスと FQDN が表示される
./stop.sh    # 停止（課金を止める）
```

## エンドポイント

- OpenL Tablets REST API: `http://<container-group>.<location>.azurecontainer.io:8080/<service-name>`
- OpenL Tablets Swagger UI: `http://<container-group>.<location>.azurecontainer.io:8080/swagger-ui.html?urls.primaryName=<service-name>`
- Deploy Service: `http://<container-group>.<location>.azurecontainer.io:8000`
