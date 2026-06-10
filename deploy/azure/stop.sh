#!/usr/bin/env bash
set -euo pipefail

RESOURCE_GROUP="${RESOURCE_GROUP:-openl-demo-rg}"
CONTAINER_GROUP="${CONTAINER_GROUP:-openl-demo}"

az container stop --name "$CONTAINER_GROUP" --resource-group "$RESOURCE_GROUP"
