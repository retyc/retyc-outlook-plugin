# Per-environment configuration. Defaults live here only; each stage re-declares what it uses.
ARG RETYC_API_BASE_URL=https://api.retyc.com
ARG RETYC_AUTH_BASE_URL=https://auth.retyc.com

#
# Build stage — produces /src/dist with all static assets
#
FROM node:24-alpine AS builder

ARG VERSION=dev
ARG RETYC_API_BASE_URL

WORKDIR /src

COPY package.json package-lock.json ./
RUN npm ci

COPY tsconfig.json tsconfig.node.json vite.config.mts taskpane.html ./
COPY public ./public
COPY manifest.xml ./
COPY src ./src

RUN VITE_RETYC_API_URL=$RETYC_API_BASE_URL npm run build

#
# Prod stage (nginx serving the static bundle)
#
FROM nginx:1.30.5-alpine-slim AS prod

ARG VERSION=dev
LABEL org.opencontainers.image.source="https://github.com/retyc/retyc-outlook-plugin"
LABEL org.opencontainers.image.description="Retyc add-in for Outlook — static asset host (taskpane, manifest, icons)."
LABEL org.opencontainers.image.licenses="MIT"
LABEL org.opencontainers.image.version="${VERSION}"

ARG RETYC_API_BASE_URL
ARG RETYC_AUTH_BASE_URL

RUN rm /etc/nginx/conf.d/default.conf
COPY ./.docker/nginx.conf /etc/nginx/conf.d/default.conf

# Point the CSP connect-src at this environment's API and OIDC issuer origins (build time only).
# nginx.conf carries the prod origins, so this rewrites nothing with the default build args.
# Only the origin of each base URL is kept — a path, if any, is dropped. An empty or malformed
# value fails the build: it would otherwise silently drop the origin from connect-src.
RUN origin() { printf '%s' "$1" | sed -nE 's#^(https?://[^/[:space:]]+).*#\1#p'; } \
 && api_origin="$(origin "$RETYC_API_BASE_URL")" \
 && auth_origin="$(origin "$RETYC_AUTH_BASE_URL")" \
 && if [ -z "$api_origin" ] || [ -z "$auth_origin" ]; then \
      echo "RETYC_API_BASE_URL and RETYC_AUTH_BASE_URL must be absolute http(s) URLs" >&2; exit 1; \
    fi \
 && sed -i \
      -e "s#https://api.retyc.com#${api_origin}#g" \
      -e "s#https://auth.retyc.com#${auth_origin}#g" \
      /etc/nginx/conf.d/default.conf \
 && nginx -t
COPY --from=builder /src/dist /usr/share/nginx/html

EXPOSE 80

HEALTHCHECK --interval=30s --timeout=3s --start-period=5s \
  CMD wget -q --spider http://localhost/taskpane.html || exit 1
