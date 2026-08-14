#!/usr/bin/env bash
set -euo pipefail

readonly REMOTE_HOST="gpu-server"
readonly REMOTE_DIR="/root/code/docx-dsl"

printf -v cargo_args ' %q' "$@"

rsync -az --delete \
  --exclude '/target/' \
  --exclude '/.git/' \
  ./ "${REMOTE_HOST}:${REMOTE_DIR}/"

ssh "${REMOTE_HOST}" \
  "cd '${REMOTE_DIR}' && export PATH=/root/.cargo/bin:\$PATH CARGO_INCREMENTAL=1 RUSTUP_DIST_SERVER=https://rsproxy.cn RUSTUP_UPDATE_ROOT=https://rsproxy.cn/rustup RUSTY_V8_MIRROR=https://gh-proxy.com/https://github.com/denoland/rusty_v8/releases/download && cargo test${cargo_args}"
