#!/usr/bin/env bash
set -eu
umask 077
if ! command -v python3 >/dev/null 2>&1; then
    echo 'Missing python3. Install with sudo apt install python3 python3-lxml' >&2
    exit 2
fi
script_dir=$(cd -P -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)
exec python3 -I "$script_dir/clean_office_metadata.py" "$@"
