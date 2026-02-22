#!/usr/bin/env bash
set -euo pipefail

print_usage() {
  cat <<'EOF'
Usage:
  ./scripts/migrate_wsl_codex_to_windows.sh [options]

Options:
  --wsl-codex-home PATH   Source Codex home inside WSL (default: $HOME/.codex)
  --win-codex-home PATH   Target Codex home under /mnt/... (auto-detected if omitted)
  --no-auth               Do not copy auth.json
  --no-config             Do not copy config.toml
  --no-history            Do not copy history.jsonl
  --no-backup             Skip backup of existing Windows Codex home
  -h, --help              Show this help

Examples:
  ./scripts/migrate_wsl_codex_to_windows.sh
  ./scripts/migrate_wsl_codex_to_windows.sh --win-codex-home /mnt/c/Users/<your-user>/.codex
EOF
}

require_cmd() {
  if ! command -v "$1" >/dev/null 2>&1; then
    echo "error: required command not found: $1" >&2
    exit 1
  fi
}

count_jsonl_files() {
  local dir="$1"
  if [ -d "$dir" ]; then
    find "$dir" -type f -name '*.jsonl' 2>/dev/null | wc -l | tr -d '[:space:]'
  else
    echo "0"
  fi
}

resolve_windows_codex_home() {
  local win_profile_win win_profile_wsl
  if command -v cmd.exe >/dev/null 2>&1 && command -v wslpath >/dev/null 2>&1; then
    win_profile_win="$(cmd.exe /c "echo %USERPROFILE%" 2>/dev/null | tr -d '\r' | tail -n 1)"
    if [ -n "$win_profile_win" ]; then
      win_profile_wsl="$(wslpath -u "$win_profile_win")"
      echo "${win_profile_wsl}/.codex"
      return
    fi
  fi
  echo "/mnt/c/Users/${USER}/.codex"
}

WSL_CODEX_HOME="${HOME}/.codex"
WIN_CODEX_HOME=""
COPY_AUTH=1
COPY_CONFIG=1
COPY_HISTORY=1
BACKUP=1

while [ "$#" -gt 0 ]; do
  case "$1" in
    --wsl-codex-home)
      [ "$#" -ge 2 ] || { echo "error: --wsl-codex-home requires a value" >&2; exit 1; }
      WSL_CODEX_HOME="$2"
      shift 2
      ;;
    --win-codex-home)
      [ "$#" -ge 2 ] || { echo "error: --win-codex-home requires a value" >&2; exit 1; }
      WIN_CODEX_HOME="$2"
      shift 2
      ;;
    --no-auth)
      COPY_AUTH=0
      shift
      ;;
    --no-config)
      COPY_CONFIG=0
      shift
      ;;
    --no-history)
      COPY_HISTORY=0
      shift
      ;;
    --no-backup)
      BACKUP=0
      shift
      ;;
    -h|--help)
      print_usage
      exit 0
      ;;
    *)
      echo "error: unknown option: $1" >&2
      print_usage
      exit 1
      ;;
  esac
done

require_cmd rsync
require_cmd python3

if [ -z "$WIN_CODEX_HOME" ]; then
  WIN_CODEX_HOME="$(resolve_windows_codex_home)"
fi

if [ ! -d "$WSL_CODEX_HOME" ]; then
  echo "error: WSL Codex home does not exist: $WSL_CODEX_HOME" >&2
  exit 1
fi

if [ ! -d "$WSL_CODEX_HOME/sessions" ]; then
  echo "error: source sessions directory not found: $WSL_CODEX_HOME/sessions" >&2
  exit 1
fi

echo "Source (WSL):   $WSL_CODEX_HOME"
echo "Target (Windows): $WIN_CODEX_HOME"
src_session_count="$(count_jsonl_files "$WSL_CODEX_HOME/sessions")"
dst_session_count_before="$(count_jsonl_files "$WIN_CODEX_HOME/sessions")"
echo "Session files: source=$src_session_count destination(before)=$dst_session_count_before"

if [ "$BACKUP" -eq 1 ] && [ -d "$WIN_CODEX_HOME" ]; then
  backup_path="${WIN_CODEX_HOME}.backup.$(date +%Y%m%d-%H%M%S)"
  echo "Creating backup: $backup_path"
  mkdir -p "$backup_path"
  if ! rsync -a \
    --exclude 'tmp/' \
    --exclude '.sandbox/' \
    --exclude '.sandbox-secrets/' \
    --exclude 'log/' \
    --exclude '*.lock' \
    "$WIN_CODEX_HOME/" "$backup_path/"; then
    echo "warning: backup hit transient file errors; continuing with migration."
  fi
fi

mkdir -p "$WIN_CODEX_HOME/sessions"
echo "Copying sessions..."
rsync -a "$WSL_CODEX_HOME/sessions/" "$WIN_CODEX_HOME/sessions/"
dst_session_count_after="$(count_jsonl_files "$WIN_CODEX_HOME/sessions")"
echo "Session files: destination(after)=$dst_session_count_after"
if [ "$dst_session_count_after" -lt "$src_session_count" ]; then
  echo "warning: destination session count is lower than source count."
fi

if [ "$COPY_HISTORY" -eq 1 ] && [ -f "$WSL_CODEX_HOME/history.jsonl" ]; then
  echo "Copying history.jsonl..."
  cp -f "$WSL_CODEX_HOME/history.jsonl" "$WIN_CODEX_HOME/history.jsonl"
fi

if [ "$COPY_AUTH" -eq 1 ] && [ -f "$WSL_CODEX_HOME/auth.json" ]; then
  echo "Copying auth.json..."
  cp -f "$WSL_CODEX_HOME/auth.json" "$WIN_CODEX_HOME/auth.json"
fi

if [ "$COPY_CONFIG" -eq 1 ] && [ -f "$WSL_CODEX_HOME/config.toml" ]; then
  echo "Copying config.toml..."
  cp -f "$WSL_CODEX_HOME/config.toml" "$WIN_CODEX_HOME/config.toml"
fi

echo "Rewriting /mnt/<drive>/... paths in structured fields..."
python3 - "$WIN_CODEX_HOME" <<'PY'
import json
import re
import sys
from pathlib import Path

home = Path(sys.argv[1])
path_keys = {"cwd", "root", "workdir", "workspace_root"}
mnt_rx = re.compile(r"^/mnt/([a-zA-Z])/(.*)$")

def to_windows_path(value: str) -> str:
    match = mnt_rx.match(value)
    if not match:
        return value
    drive = match.group(1).upper()
    rest = match.group(2).replace("/", "\\")
    return f"{drive}:\\{rest}"

def rewrite(value):
    if isinstance(value, dict):
      out = {}
      for k, v in value.items():
          if isinstance(v, str) and k in path_keys:
              out[k] = to_windows_path(v)
          else:
              out[k] = rewrite(v)
      return out
    if isinstance(value, list):
      return [rewrite(item) for item in value]
    return value

def process_jsonl(path: Path) -> tuple[int, int]:
    total = 0
    changed = 0
    output_lines = []
    with path.open("r", encoding="utf-8") as infile:
        for line in infile:
            total += 1
            stripped = line.rstrip("\n")
            if not stripped:
                output_lines.append(stripped)
                continue
            try:
                obj = json.loads(stripped)
            except json.JSONDecodeError:
                output_lines.append(stripped)
                continue
            new_obj = rewrite(obj)
            if new_obj != obj:
                changed += 1
            output_lines.append(json.dumps(new_obj, ensure_ascii=False, separators=(",", ":")))
    if changed:
        with path.open("w", encoding="utf-8", newline="\n") as outfile:
            outfile.write("\n".join(output_lines) + "\n")
    return total, changed

files = list((home / "sessions").rglob("*.jsonl"))
history = home / "history.jsonl"
if history.exists():
    files.append(history)

line_total = 0
line_changed = 0
file_changed = 0
for file_path in files:
    t, c = process_jsonl(file_path)
    line_total += t
    line_changed += c
    if c:
        file_changed += 1

print(f"Processed {len(files)} files, {line_total} lines; updated {file_changed} files ({line_changed} lines changed).")
PY

cat <<EOF
Migration complete.
Next steps:
1. Close this WSL shell.
2. Start your app in CMD/PowerShell.
3. Point Codex home/runtime to Windows paths (for example: C:\\Users\\<your-user>\\.codex and C:\\Users\\<your-user>\\AppData\\Roaming\\npm\\codex.cmd).
EOF
