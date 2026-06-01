# syntax=docker/dockerfile:1.6
ARG HUGO_VERSION=0.139.0
ARG NGINX_VERSION=1.27-alpine

FROM hugomods/hugo:${HUGO_VERSION} AS build
ARG SCRIPT_VERSION=dev
WORKDIR /src
COPY . .
RUN hugo --minify --gc --destination /out && \
    sed -i "s/__KMS_VERSION__/${SCRIPT_VERSION}/g" /out/scripts/*.ps1

FROM nginx:${NGINX_VERSION}
LABEL org.opencontainers.image.source="https://forgejo.viktorbarzin.me/viktor/kms-website" \
      org.opencontainers.image.description="kms.viktorbarzin.me — KMS activation reference"

# Strip default config so the bundled simple one stays
RUN rm -f /etc/nginx/conf.d/default.conf
COPY nginx.conf /etc/nginx/conf.d/default.conf
COPY --from=build /out /usr/share/nginx/html
EXPOSE 80
HEALTHCHECK --interval=30s --timeout=3s --start-period=5s --retries=3 \
  CMD wget -qO- http://127.0.0.1/ >/dev/null 2>&1 || exit 1
