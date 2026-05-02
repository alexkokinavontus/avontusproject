#!/bin/bash
# =============================================================================
# AzureReader — Deploy to Azure Static Web Apps
# Run this in Azure Cloud Shell: https://shell.azure.com
# Repo: https://github.com/alexkokinavontus/avontusproject
# =============================================================================

set -e

# ── Config ────────────────────────────────────────────────────────────────────
TENANT_ID="bd98204b-b981-4d03-8796-356d537927eb"
SUBSCRIPTION_NAME="avontus"          # or use subscription ID
RESOURCE_GROUP="rg-azurereader"
LOCATION="eastus2"
SWA_NAME="azurereader-avontus"
GITHUB_REPO="alexkokinavontus/avontusproject"
GITHUB_BRANCH="main"
APP_LOCATION="/"
OUTPUT_LOCATION="dist"
BUILD_PRESET="react"

# Colors
GREEN='\033[0;32m'
BLUE='\033[0;34m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
NC='\033[0m'

echo -e "${BLUE}"
echo "╔══════════════════════════════════════════════════════╗"
echo "║         AzureReader — Azure Deployment Script        ║"
echo "║                   avontus.com tenant                 ║"
echo "╚══════════════════════════════════════════════════════╝"
echo -e "${NC}"

# ── Step 1: Set subscription ──────────────────────────────────────────────────
echo -e "${YELLOW}[1/7] Setting subscription...${NC}"
az account set --subscription "$SUBSCRIPTION_NAME" 2>/dev/null || \
az account set --subscription "$TENANT_ID"
CURRENT_SUB=$(az account show --query name -o tsv)
echo -e "${GREEN}✓ Using subscription: $CURRENT_SUB${NC}"

# ── Step 2: Create resource group ─────────────────────────────────────────────
echo -e "${YELLOW}[2/7] Creating resource group '$RESOURCE_GROUP' in $LOCATION...${NC}"
az group create \
  --name "$RESOURCE_GROUP" \
  --location "$LOCATION" \
  --output none
echo -e "${GREEN}✓ Resource group ready${NC}"

# ── Step 3: Prompt for GitHub token ───────────────────────────────────────────
echo -e "${YELLOW}[3/7] GitHub authentication${NC}"
echo ""
echo "  You need a GitHub Personal Access Token with scopes:"
echo "    ✓ repo"
echo "    ✓ workflow"
echo ""
echo "  Create one at: https://github.com/settings/tokens/new"
echo ""
read -s -p "  Paste your GitHub PAT: " GITHUB_TOKEN
echo ""
if [ -z "$GITHUB_TOKEN" ]; then
  echo -e "${RED}✗ No token provided. Exiting.${NC}"
  exit 1
fi
echo -e "${GREEN}✓ Token received${NC}"

# ── Step 4: Prompt for Azure client secret ────────────────────────────────────
echo -e "${YELLOW}[4/7] Azure app secret (for AzureReader app registration)${NC}"
echo ""
echo "  App:    AzureReader (3977e66a-cdf1-419d-9d0d-70e8cf3a76ed)"
echo "  Tenant: bd98204b-b981-4d03-8796-356d537927eb"
echo ""
read -s -p "  Paste your Azure client secret: " AZURE_CLIENT_SECRET
echo ""
if [ -z "$AZURE_CLIENT_SECRET" ]; then
  echo -e "${RED}✗ No secret provided. Exiting.${NC}"
  exit 1
fi
echo -e "${GREEN}✓ Secret received${NC}"

# ── Step 5: Create Static Web App linked to GitHub ────────────────────────────
echo -e "${YELLOW}[5/7] Creating Static Web App '$SWA_NAME'...${NC}"
echo "  Linking to: https://github.com/$GITHUB_REPO (branch: $GITHUB_BRANCH)"

SWA_OUTPUT=$(az staticwebapp create \
  --name "$SWA_NAME" \
  --resource-group "$RESOURCE_GROUP" \
  --location "$LOCATION" \
  --source "https://github.com/$GITHUB_REPO" \
  --branch "$GITHUB_BRANCH" \
  --app-location "$APP_LOCATION" \
  --output-location "$OUTPUT_LOCATION" \
  --login-with-github \
  --token "$GITHUB_TOKEN" \
  --output json 2>&1)

if echo "$SWA_OUTPUT" | grep -q '"id"'; then
  SWA_URL=$(echo "$SWA_OUTPUT" | python3 -c "import sys,json; d=json.load(sys.stdin); print(d.get('defaultHostname',''))" 2>/dev/null || echo "")
  echo -e "${GREEN}✓ Static Web App created${NC}"
  [ -n "$SWA_URL" ] && echo -e "  URL: ${BLUE}https://$SWA_URL${NC}"
else
  echo -e "${RED}✗ SWA creation failed. Output:${NC}"
  echo "$SWA_OUTPUT"
  exit 1
fi

# ── Step 6: Set app setting (VITE_AZURE_CLIENT_SECRET) ───────────────────────
echo -e "${YELLOW}[6/7] Setting VITE_AZURE_CLIENT_SECRET app setting...${NC}"
az staticwebapp appsettings set \
  --name "$SWA_NAME" \
  --resource-group "$RESOURCE_GROUP" \
  --setting-names "VITE_AZURE_CLIENT_SECRET=$AZURE_CLIENT_SECRET" \
  --output none
echo -e "${GREEN}✓ App setting configured${NC}"

# ── Step 7: Add GitHub repo secrets via GitHub API ───────────────────────────
echo -e "${YELLOW}[7/7] Configuring GitHub repo secrets...${NC}"

# Get SWA deployment token
DEPLOY_TOKEN=$(az staticwebapp secrets list \
  --name "$SWA_NAME" \
  --resource-group "$RESOURCE_GROUP" \
  --query "properties.apiKey" -o tsv 2>/dev/null || echo "")

if [ -z "$DEPLOY_TOKEN" ]; then
  echo -e "${YELLOW}  ⚠ Could not retrieve deployment token automatically.${NC}"
  echo "    Get it from: Azure Portal → $SWA_NAME → Manage deployment token"
  echo "    Then add as GitHub secret: AZURE_STATIC_WEB_APPS_API_TOKEN"
else
  # Get repo public key for secret encryption
  PUBKEY_RESP=$(curl -s -H "Authorization: token $GITHUB_TOKEN" \
    "https://api.github.com/repos/$GITHUB_REPO/actions/secrets/public-key")
  KEY_ID=$(echo "$PUBKEY_RESP" | python3 -c "import sys,json; print(json.load(sys.stdin)['key_id'])" 2>/dev/null || echo "")
  PUB_KEY=$(echo "$PUBKEY_RESP" | python3 -c "import sys,json; print(json.load(sys.stdin)['key'])" 2>/dev/null || echo "")

  if [ -n "$KEY_ID" ] && [ -n "$PUB_KEY" ]; then
    # Encrypt and upload AZURE_STATIC_WEB_APPS_API_TOKEN
    ENC_DEPLOY=$(python3 -c "
import base64, sys
from cryptography.hazmat.primitives.asymmetric.x25519 import X25519PublicKey
from cryptography.hazmat.primitives.asymmetric import padding as apadding
from cryptography.hazmat.primitives import hashes
from nacl.public import PublicKey, SealedBox
import nacl.encoding
pub = PublicKey(base64.b64decode('$PUB_KEY'))
box = SealedBox(pub)
enc = box.encrypt(b'$DEPLOY_TOKEN')
print(base64.b64encode(enc).decode())
" 2>/dev/null || echo "")

    ENC_SECRET=$(python3 -c "
import base64
from nacl.public import PublicKey, SealedBox
pub = PublicKey(base64.b64decode('$PUB_KEY'))
box = SealedBox(pub)
enc = box.encrypt(b'$AZURE_CLIENT_SECRET')
print(base64.b64encode(enc).decode())
" 2>/dev/null || echo "")

    if [ -n "$ENC_DEPLOY" ]; then
      # Upload AZURE_STATIC_WEB_APPS_API_TOKEN
      curl -s -X PUT \
        -H "Authorization: token $GITHUB_TOKEN" \
        -H "Content-Type: application/json" \
        "https://api.github.com/repos/$GITHUB_REPO/actions/secrets/AZURE_STATIC_WEB_APPS_API_TOKEN" \
        -d "{\"encrypted_value\":\"$ENC_DEPLOY\",\"key_id\":\"$KEY_ID\"}" > /dev/null
      echo -e "${GREEN}  ✓ AZURE_STATIC_WEB_APPS_API_TOKEN added to GitHub${NC}"
    fi

    if [ -n "$ENC_SECRET" ]; then
      # Upload AZURE_CLIENT_SECRET
      curl -s -X PUT \
        -H "Authorization: token $GITHUB_TOKEN" \
        -H "Content-Type: application/json" \
        "https://api.github.com/repos/$GITHUB_REPO/actions/secrets/AZURE_CLIENT_SECRET" \
        -d "{\"encrypted_value\":\"$ENC_SECRET\",\"key_id\":\"$KEY_ID\"}" > /dev/null
      echo -e "${GREEN}  ✓ AZURE_CLIENT_SECRET added to GitHub${NC}"
    fi
  else
    echo -e "${YELLOW}  ⚠ Could not auto-upload secrets (pynacl may not be available).${NC}"
    echo "    Manually add these 2 secrets to: https://github.com/$GITHUB_REPO/settings/secrets/actions"
    echo ""
    echo "    Secret 1 name:  AZURE_STATIC_WEB_APPS_API_TOKEN"
    echo "    Secret 1 value: $DEPLOY_TOKEN"
    echo ""
    echo "    Secret 2 name:  AZURE_CLIENT_SECRET"
    echo "    Secret 2 value: (your Azure app secret)"
  fi
fi

# ── Summary ───────────────────────────────────────────────────────────────────
echo ""
echo -e "${BLUE}══════════════════════════════════════════════════════${NC}"
echo -e "${GREEN}  Deployment complete!${NC}"
echo -e "${BLUE}══════════════════════════════════════════════════════${NC}"
echo ""
[ -n "$SWA_URL" ] && echo "  App URL:        https://$SWA_URL"
echo "  Resource group: $RESOURCE_GROUP"
echo "  SWA name:       $SWA_NAME"
echo "  GitHub repo:    https://github.com/$GITHUB_REPO"
echo ""
echo "  Next steps:"
echo "  1. Push code to the '$GITHUB_BRANCH' branch to trigger deployment"
echo "  2. Monitor: https://github.com/$GITHUB_REPO/actions"
echo "  3. Azure Portal → $SWA_NAME → view deployment logs"
echo ""
echo -e "${YELLOW}  ⚠ Rotate your Azure client secret after first successful deploy${NC}"
echo ""
