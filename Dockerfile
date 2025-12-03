# Debian 13
FROM debian:trixie AS build

RUN <<EOF
  apt update
  apt install -y nodejs npm node-grunt-cli
EOF

RUN --mount=type=bind,source=./,target=/temp,rw <<EOF
  cd /temp/build

  npm install

  grunt

  # Persist the built files in the image
  cp -r /temp/deploy/sdkjs /sdkjs
EOF

FROM alpine:3.22

COPY --from=build /sdkjs /sdkjs
