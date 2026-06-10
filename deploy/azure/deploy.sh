#!/usr/bin/env bash
set -euo pipefail

# Required: build and push the deploy-service image first, e.g.
#   docker build -t <registry>/openl-deploy-service:latest deploy-service
#   docker push <registry>/openl-deploy-service:latest
: "${DEPLOY_SERVICE_IMAGE:?Set DEPLOY_SERVICE_IMAGE to the pushed deploy-service image}"

RESOURCE_GROUP="${RESOURCE_GROUP:-openl-demo-rg}"
LOCATION="${LOCATION:-japaneast}"
STORAGE_ACCOUNT="${STORAGE_ACCOUNT:-openldemostorage}"
SHARE_NAME="${SHARE_NAME:-openl-deployment}"
CONTAINER_GROUP="${CONTAINER_GROUP:-openl-demo}"

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

az group create --name "$RESOURCE_GROUP" --location "$LOCATION" --output none

az storage account create \
  --name "$STORAGE_ACCOUNT" \
  --resource-group "$RESOURCE_GROUP" \
  --location "$LOCATION" \
  --sku Standard_LRS \
  --output none

STORAGE_KEY=$(az storage account keys list \
  --account-name "$STORAGE_ACCOUNT" \
  --resource-group "$RESOURCE_GROUP" \
  --query '[0].value' -o tsv)

az storage share create \
  --name "$SHARE_NAME" \
  --account-name "$STORAGE_ACCOUNT" \
  --account-key "$STORAGE_KEY" \
  --output none

OPENL_PUBLIC_URL="${OPENL_PUBLIC_URL:-}"
if [[ -z "$OPENL_PUBLIC_URL" ]]; then
  # Placeholder; ACI assigns the public IP only after creation. Re-run
  # start.sh to print the IP, then update the running group if needed.
  OPENL_PUBLIC_URL="http://PENDING:8080"
fi

export LOCATION CONTAINER_GROUP DEPLOY_SERVICE_IMAGE OPENL_PUBLIC_URL SHARE_NAME STORAGE_ACCOUNT STORAGE_KEY

envsubst < "$SCRIPT_DIR/container-group.yaml.template" > "$SCRIPT_DIR/container-group.yaml"

az container create \
  --resource-group "$RESOURCE_GROUP" \
  --file "$SCRIPT_DIR/container-group.yaml" \
  --output none

az container show \
  --name "$CONTAINER_GROUP" \
  --resource-group "$RESOURCE_GROUP" \
  --query "{state:instanceView.state, ip:ipAddress.ip}" -o table
