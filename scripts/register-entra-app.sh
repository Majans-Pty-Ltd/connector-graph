#!/usr/bin/env bash
# Register Entra app for connector-graph and store credentials in 1Password.
#
# Prerequisites: az CLI authenticated, op CLI authenticated
# Usage: bash scripts/register-entra-app.sh

set -euo pipefail

APP_NAME="Majans-Graph-MCP-Agent"
VAULT="Majans Dev"

echo "=== Step 1: Register Entra app ==="
APP_ID=$(az ad app create \
  --display-name "$APP_NAME" \
  --sign-in-audience AzureADMyOrg \
  --query appId -o tsv)
echo "App ID: $APP_ID"

echo "=== Step 2: Create service principal ==="
az ad sp create --id "$APP_ID" > /dev/null 2>&1 || true
echo "Service principal created"

echo "=== Step 3: Add application permissions ==="
# Microsoft Graph app ID: 00000003-0000-0000-c000-000000000000
GRAPH_APP_ID="00000003-0000-0000-c000-000000000000"

# User.ReadWrite.All
az ad app permission add --id "$APP_ID" --api "$GRAPH_APP_ID" --api-permissions 741f803b-c850-494e-b5df-cde7c675a1ca=Role
# Group.ReadWrite.All
az ad app permission add --id "$APP_ID" --api "$GRAPH_APP_ID" --api-permissions 62a82d76-70ea-41e2-9197-370581804d09=Role
# Directory.Read.All
az ad app permission add --id "$APP_ID" --api "$GRAPH_APP_ID" --api-permissions 7ab1d382-f21e-4acd-a863-ba3e13f7da61=Role
# Reports.Read.All
az ad app permission add --id "$APP_ID" --api "$GRAPH_APP_ID" --api-permissions 230c1aed-a721-4c5d-9cb4-a90514e508ef=Role
echo "Permissions added"

echo "=== Step 4: Grant admin consent ==="
az ad app permission admin-consent --id "$APP_ID"
echo "Admin consent granted"

echo "=== Step 5: Create client secret (2-year expiry) ==="
SECRET=$(az ad app credential reset \
  --id "$APP_ID" \
  --display-name "connector-graph" \
  --years 2 \
  --query password -o tsv)
echo "Secret created"

# Get tenant ID
TENANT_ID=$(az account show --query tenantId -o tsv)

echo "=== Step 6: Store in 1Password ==="
op item create \
  --category "API Credential" \
  --title "Graph MCP Agent" \
  --vault "$VAULT" \
  --tags "azure,entra,graph,mcp" \
  "tenant_id=$TENANT_ID" \
  "client_id=$APP_ID" \
  "client_secret=$SECRET"
echo "1Password item created"

echo ""
echo "=== Done ==="
echo "App Name:      $APP_NAME"
echo "Client ID:     $APP_ID"
echo "Tenant ID:     $TENANT_ID"
echo "1Password:     $VAULT / Graph MCP Agent"
echo ""
echo "Next: copy .env.template to .env, or use op run --env-file=.env.template -- npm start"
