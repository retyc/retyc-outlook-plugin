#!/bin/sh
# Resolves runtime env vars into the manifest.xml AND the nginx config at container start.
# Run by nginx:alpine's /docker-entrypoint.sh as part of /docker-entrypoint.d/, BEFORE nginx
# loads its config.
#
# This lets the same image serve any environment (preprod, prod, staging) just by passing
# different env vars:
#
#   docker run \
#     -e BASE_URL=https://outlook-addin.retyc.com \
#     -e LANDING_URL=https://retyc.com/ \
#     ghcr.io/retyc/retyc-outlook-plugin

set -e

: "${BASE_URL:=https://outlook.retyc.com}"
: "${LANDING_URL:=https://retyc.com/}"

# Source placeholder used in the canonical manifest.xml. Same string is what the dev tooling
# (scripts/dev-setup.js, webpack CopyPlugin transform) targets.
SOURCE_URL="https://outlook.retyc.com"

# 1. manifest.xml — the URLs Outlook fetches (taskpane, icons, manifest itself).
MANIFEST_TEMPLATE=/etc/retyc/manifest.xml.template
MANIFEST_TARGET=/usr/share/nginx/html/manifest.xml
if [ ! -f "$MANIFEST_TEMPLATE" ]; then
  echo "[retyc-entrypoint] missing template at $MANIFEST_TEMPLATE — image is broken" >&2
  exit 1
fi
sed "s|${SOURCE_URL}|${BASE_URL}|g" "$MANIFEST_TEMPLATE" > "$MANIFEST_TARGET"

# 2. nginx config — the bare-domain redirect target. We re-render the whole site config
#    rather than mutating /etc/nginx/conf.d/default.conf in place, so a restart with a new
#    LANDING_URL takes effect without rebuilding the image.
NGINX_TEMPLATE=/etc/retyc/nginx.conf.template
NGINX_TARGET=/etc/nginx/conf.d/default.conf
if [ ! -f "$NGINX_TEMPLATE" ]; then
  echo "[retyc-entrypoint] missing template at $NGINX_TEMPLATE — image is broken" >&2
  exit 1
fi
sed "s|__LANDING_URL__|${LANDING_URL}|g" "$NGINX_TEMPLATE" > "$NGINX_TARGET"

echo "[retyc-entrypoint] BASE_URL=${BASE_URL}  LANDING_URL=${LANDING_URL}"
