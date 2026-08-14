#!/usr/bin/env bash
set -euo pipefail

readonly PROJECT_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
readonly VALIDATOR_PROJECT="${PROJECT_ROOT}/tools/openxml-validator/OpenXmlValidator.csproj"
readonly NUGET_CONFIG="${PROJECT_ROOT}/NuGet.Config"

if [[ -n "${DOTNET_BIN:-}" ]]; then
  readonly DOTNET="${DOTNET_BIN}"
elif command -v dotnet >/dev/null 2>&1; then
  readonly DOTNET="$(command -v dotnet)"
elif [[ -x /root/.dotnet/dotnet ]]; then
  readonly DOTNET="/root/.dotnet/dotnet"
else
  printf '%s\n' 'Open XML integration tests require .NET SDK 8.' >&2
  printf '%s\n' 'Fix: install SDK 8.0.127, or set DOTNET_BIN to the dotnet executable.' >&2
  exit 2
fi

readonly TEMP_DIR="$(mktemp -d)"
trap 'rm -rf "${TEMP_DIR}"' EXIT

cd "${PROJECT_ROOT}"
cargo build --locked
"${DOTNET}" restore "${VALIDATOR_PROJECT}" --configfile "${NUGET_CONFIG}" --locked-mode
"${PROJECT_ROOT}/target/debug/docx-jsx" \
  "${PROJECT_ROOT}/tests/fixtures/openxml-valid.tsx" \
  --output "${TEMP_DIR}/valid.docx"
"${DOTNET}" run --project "${VALIDATOR_PROJECT}" --no-restore -- "${TEMP_DIR}/valid.docx"

printf '%s' 'not a DOCX package' > "${TEMP_DIR}/invalid.docx"
if "${DOTNET}" run --project "${VALIDATOR_PROJECT}" --no-restore -- "${TEMP_DIR}/invalid.docx" \
  2> "${TEMP_DIR}/invalid.jsonl"; then
  printf '%s\n' 'OpenXmlValidator unexpectedly accepted an invalid package.' >&2
  exit 1
fi
grep -q 'package-open-failed' "${TEMP_DIR}/invalid.jsonl"
