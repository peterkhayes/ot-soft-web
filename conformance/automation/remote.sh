#!/usr/bin/env bash
# Runs a command on the configured Windows remote machine via SSH.
# Usage: ./remote.sh <command>
# Example: ./remote.sh "python --version"
#
# The command is executed in the repo directory on the remote machine.

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
CONF="$SCRIPT_DIR/remote.conf"

if [[ ! -f "$CONF" ]]; then
    echo "Error: $CONF not found." >&2
    exit 1
fi

SSH_USER=""
SSH_HOST=""
REMOTE_REPO=""
while IFS="=" read -r key value; do
    key="${key%%$'\r'}"
    value="${value%%$'\r'}"
    value="${value%"${value##*[! ]}"}"
    [[ -z "$key" || "$key" =~ ^# ]] && continue
    case "$key" in
        SSH_USER)    SSH_USER="$value" ;;
        SSH_HOST)    SSH_HOST="$value" ;;
        REMOTE_REPO) REMOTE_REPO="$value" ;;
    esac
done < "$CONF"

if [[ -z "$SSH_USER" || -z "$SSH_HOST" || -z "$REMOTE_REPO" ]]; then
    echo "Error: remote.conf must define SSH_USER, SSH_HOST, and REMOTE_REPO." >&2
    exit 1
fi

CMD="$*"
if [[ -z "$CMD" ]]; then
    echo "Usage: $0 <command>" >&2
    exit 1
fi

ssh -o ConnectTimeout=10 "${SSH_USER}@${SSH_HOST}" "cd \"${REMOTE_REPO}\" && ${CMD}"
