#
# Build stage — produces /src/dist with all static assets
#
FROM node:22-alpine AS builder

ARG VERSION=dev

WORKDIR /src

# Lockfiles first for layer caching
COPY package.json package-lock.json ./
RUN npm ci

# Sources required by webpack
COPY tsconfig.json webpack.config.js ./
COPY assets ./assets
COPY manifest.xml ./
COPY src ./src

# Build. The manifest URLs stay as `https://localhost:3000` here — they are
# resolved at container start via /docker-entrypoint.d/40-substitute-base-url.sh
# so the same image can serve any environment by changing the BASE_URL env var.
RUN npm run build

#
# Prod stage (nginx serving the static bundle)
#
FROM nginx:1.30.0-alpine-slim AS prod

ARG VERSION=dev
LABEL org.opencontainers.image.source="https://github.com/retyc/retyc-outlook-plugin"
LABEL org.opencontainers.image.description="Retyc add-in for Outlook — static asset host (taskpane, manifest, icons)."
LABEL org.opencontainers.image.licenses="MIT"
LABEL org.opencontainers.image.version="${VERSION}"

# nginx config + manifest are kept as templates outside the web root. The entrypoint
# script renders both at container start, so the same image works for any environment.
RUN rm /etc/nginx/conf.d/default.conf && mkdir -p /etc/retyc
COPY ./.docker/nginx.conf /etc/retyc/nginx.conf.template

# Static bundle. The manifest is moved out of the web root and kept as a template too.
COPY --from=builder /src/dist /usr/share/nginx/html
RUN mv /usr/share/nginx/html/manifest.xml /etc/retyc/manifest.xml.template

# Runtime substitution. nginx:alpine's /docker-entrypoint.sh runs every */*.sh in
# /docker-entrypoint.d/ before launching nginx.
COPY ./.docker/40-substitute-base-url.sh /docker-entrypoint.d/40-substitute-base-url.sh
RUN chmod +x /docker-entrypoint.d/40-substitute-base-url.sh

# Defaults match the canonical placeholder in manifest.xml — running the image with no env
# overrides serves a manifest that's directly usable for the prod deployment at outlook.retyc.com.
# Override either env var to retarget preprod / staging / local docker testing without rebuilding.
ENV BASE_URL=https://outlook.retyc.com \
    LANDING_URL=https://retyc.com/

EXPOSE 80

HEALTHCHECK --interval=30s --timeout=3s --start-period=5s \
  CMD wget -q --spider http://localhost/taskpane.html || exit 1
