#!/usr/bin/env bash
set -euo pipefail

# Required: build and push the deploy-service image to a registry first, e.g.
#   az acr login --name <registry-name>
#   docker tag openl-deploy-service:latest <login-server>/openl-deploy-service:latest
#   docker push <login-server>/openl-deploy-service:latest
: "${DEPLOY_SERVICE_IMAGE:?Set DEPLOY_SERVICE_IMAGE to the pushed deploy-service image}"

# Required: credentials for the registry that hosts DEPLOY_SERVICE_IMAGE, e.g.
#   az acr credential show --name <registry-name>
: "${REGISTRY_SERVER:?Set REGISTRY_SERVER to the registry login server}"
: "${REGISTRY_USERNAME:?Set REGISTRY_USERNAME to the registry username}"
: "${REGISTRY_PASSWORD:?Set REGISTRY_PASSWORD to the registry password}"

RESOURCE_GROUP="${RESOURCE_GROUP:-openl-demo-rg}"
LOCATION="${LOCATION:-japaneast}"
STORAGE_ACCOUNT="${STORAGE_ACCOUNT:-openldemostorage}"
SHARE_NAME="${SHARE_NAME:-openl-deployment}"
CONTAINER_GROUP="${CONTAINER_GROUP:-openl-demo}"
DNS_NAME_LABEL="${DNS_NAME_LABEL:-$CONTAINER_GROUP}"

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

# dnsNameLabel makes the public URL predictable before creation, avoiding
# the chicken-and-egg problem of needing the assigned IP up front.
OPENL_PUBLIC_URL="${OPENL_PUBLIC_URL:-http://${DNS_NAME_LABEL}.${LOCATION}.azurecontainer.io:8080}"

export LOCATION CONTAINER_GROUP DEPLOY_SERVICE_IMAGE DNS_NAME_LABEL OPENL_PUBLIC_URL \
  REGISTRY_SERVER REGISTRY_USERNAME REGISTRY_PASSWORD SHARE_NAME STORAGE_ACCOUNT STORAGE_KEY

envsubst < "$SCRIPT_DIR/container-group.yaml.template" > "$SCRIPT_DIR/container-group.yaml"

az container create \
  --resource-group "$RESOURCE_GROUP" \
  --file "$SCRIPT_DIR/container-group.yaml" \
  --output none

az container show \
  --name "$CONTAINER_GROUP" \
  --resource-group "$RESOURCE_GROUP" \
  --query "{state:instanceView.state, ip:ipAddress.ip, fqdn:ipAddress.fqdn}" -o table

echo "OPENL_PUBLIC_URL=$OPENL_PUBLIC_URL"
echo "Deploy Service:  http://${DNS_NAME_LABEL}.${LOCATION}.azurecontainer.io:8000"
