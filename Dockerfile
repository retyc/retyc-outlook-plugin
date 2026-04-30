#
# Build stage — produces /src/dist with all static assets
#
FROM node:24-alpine AS builder

ARG VERSION=dev
ARG VITE_RETYC_API_URL
ENV VITE_RETYC_API_URL=${VITE_RETYC_API_URL}

WORKDIR /src

COPY package.json package-lock.json ./
RUN npm ci

COPY tsconfig.json tsconfig.node.json vite.config.mts taskpane.html ./
COPY public ./public
COPY manifest.xml ./
COPY src ./src

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

RUN rm /etc/nginx/conf.d/default.conf
COPY ./.docker/nginx.conf /etc/nginx/conf.d/default.conf
COPY --from=builder /src/dist /usr/share/nginx/html

EXPOSE 80

HEALTHCHECK --interval=30s --timeout=3s --start-period=5s \
  CMD wget -q --spider http://localhost/taskpane.html || exit 1
