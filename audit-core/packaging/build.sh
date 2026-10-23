#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
WORKSPACE_DIR="$(cd "$ROOT_DIR/.." && pwd)"
TARGET="x86_64-unknown-linux-gnu.2.28"
TARGET_DIR="$WORKSPACE_DIR/target/x86_64-unknown-linux-gnu/release"

for command_name in cargo rustup zig objdump readelf; do
    command -v "$command_name" >/dev/null 2>&1 || {
        echo "missing required build command: $command_name" >&2
        exit 1
    }
done

rustup target add x86_64-unknown-linux-gnu
cd "$WORKSPACE_DIR"
cargo zigbuild --locked --release --workspace --target "$TARGET"

check_glibc_cap() {
    local binary="$1"
    local versions maximum
    versions="$(objdump -T "$binary" | sed -n 's/.*\(GLIBC_[0-9][0-9.]*\).*/\1/p' | sed 's/^GLIBC_//' | sort -V -u)"
    maximum="$(printf '%s\n' "$versions" | tail -n 1)"
    if [[ -n "$maximum" ]] && [[ "$(printf '%s\n' "2.28" "$maximum" | sort -V | tail -n 1)" != "2.28" ]]; then
        echo "GLIBC requirement exceeds 2.28: $binary requires $maximum" >&2
        exit 1
    fi
    printf 'GLIBC max: %s (%s)\n' "${maximum:-none (static or no GLIBC symbols)}" "$binary"
    readelf -d "$binary" | grep -E 'NEEDED|RUNPATH|RPATH' || true
}

for binary_name in umya-probe rust-fastpath-probe; do
    binary_path="$TARGET_DIR/$binary_name"
    [[ -x "$binary_path" ]] || {
        echo "expected release binary missing: $binary_path" >&2
        exit 1
    }
    check_glibc_cap "$binary_path"
done

if [[ "${ENABLE_GO_EDITOR:-0}" == "1" ]]; then
    command -v go >/dev/null 2>&1 || { echo "ENABLE_GO_EDITOR=1 but Go is unavailable" >&2; exit 1; }
    [[ -f "$ROOT_DIR/core-go/go.mod" ]] || { echo "Go editor is not enabled: core-go/go.mod is absent" >&2; exit 1; }
    mkdir -p "$WORKSPACE_DIR/target"
    (cd "$ROOT_DIR/core-go" && CGO_ENABLED=0 GOOS=linux GOARCH=amd64 go build -trimpath -o "$WORKSPACE_DIR/target/core-go-editor" ./src)
fi

echo "ABI-checked probe binaries are in $TARGET_DIR."
echo "Deb creation is intentionally withheld until a production CLI and target-system acceptance exist."
