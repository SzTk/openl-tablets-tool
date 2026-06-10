#!/usr/bin/env bash
set -euo pipefail

RESOURCE_GROUP="${RESOURCE_GROUP:-openl-demo-rg}"
CONTAINER_GROUP="${CONTAINER_GROUP:-openl-demo}"

az container start --name "$CONTAINER_GROUP" --resource-group "$RESOURCE_GROUP"

az container show \
  --name "$CONTAINER_GROUP" \
  --resource-group "$RESOURCE_GROUP" \
  --query "{state:instanceView.state, ip:ipAddress.ip}" -o table
