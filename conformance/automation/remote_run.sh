#!/usr/bin/env bash
# Triggers a conformance test run on the remote Windows control server.
# Reads SSH_HOST and port from remote.conf.
#
# Usage:
#   ./remote_run.sh                              # run all tests
#   ./remote_run.sh --filter "rcd_defaults$"     # filter by regex
#   ./remote_run.sh --no-cleanup                 # keep VB6 output dirs
#   ./remote_run.sh --verbose                    # verbose logging
#   ./remote_run.sh --status                     # check server status
#   ./remote_run.sh --results                    # get last run results
#   ./remote_run.sh --reload                     # git pull + restart server

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
CONF="$SCRIPT_DIR/remote.conf"

if [ ! -f "$CONF" ]; then
    echo "Error: $CONF not found." >&2
    exit 1
fi

SSH_HOST=""
SERVER_PORT="8377"
while IFS="=" read -r key value; do
    key=$(echo "$key" | tr -d '\r')
    value=$(echo "$value" | tr -d '\r')
    case "$key" in
        SSH_HOST)    SSH_HOST="$value" ;;
        SERVER_PORT) SERVER_PORT="$value" ;;
    esac
done < "$CONF"

if [ -z "$SSH_HOST" ]; then
    echo "Error: remote.conf must define SSH_HOST." >&2
    exit 1
fi

BASE_URL="http://${SSH_HOST}:${SERVER_PORT}"

# Parse arguments
FILTER=""
NO_CLEANUP=false
VERBOSE=false
MODE="run"

while [ $# -gt 0 ]; do
    case "$1" in
        --filter)    FILTER="$2"; shift 2 ;;
        --no-cleanup) NO_CLEANUP=true; shift ;;
        --verbose|-v) VERBOSE=true; shift ;;
        --status)    MODE="status"; shift ;;
        --results)   MODE="results"; shift ;;
        --reload)    MODE="reload"; shift ;;
        *) echo "Unknown argument: $1" >&2; exit 1 ;;
    esac
done

case "$MODE" in
    status)
        curl -s "$BASE_URL/status" | python3 -m json.tool
        ;;
    reload)
        echo "Reloading server (git pull + restart)..."
        RESP=$(curl -s -X POST "$BASE_URL/reload")
        echo "$RESP" | python3 -m json.tool
        echo ""
        echo "Waiting for server to restart..."
        sleep 2
        # Poll until server is back
        for i in $(seq 1 10); do
            if curl -s "$BASE_URL/status" > /dev/null 2>&1; then
                echo "Server is back up."
                exit 0
            fi
            sleep 1
        done
        echo "Warning: server did not come back within 12 seconds."
        exit 1
        ;;
    results)
        RESP=$(curl -s "$BASE_URL/results")
        # Print the output field as plain text, then the results as JSON
        echo "$RESP" | python3 -c "
import sys, json
data = json.load(sys.stdin)
if data.get('output'):
    print(data['output'])
    print()
if data.get('results'):
    print('Exit code:', data['results'].get('exit_code'))
    print('Success:', data['results'].get('success'))
if data.get('running'):
    print('(still running)')
"
        ;;
    run)
        # Build JSON body
        BODY='{'
        SEP=""
        if [ -n "$FILTER" ]; then
            BODY="${BODY}${SEP}\"filter\":\"${FILTER}\""
            SEP=","
        fi
        if [ "$NO_CLEANUP" = "true" ]; then
            BODY="${BODY}${SEP}\"no_cleanup\":true"
            SEP=","
        fi
        if [ "$VERBOSE" = "true" ]; then
            BODY="${BODY}${SEP}\"verbose\":true"
            SEP=","
        fi
        BODY="${BODY}}"

        echo "Sending run request to $BASE_URL/run ..."
        echo "  Body: $BODY"
        echo ""

        RESP=$(curl -s -X POST -H "Content-Type: application/json" -d "$BODY" "$BASE_URL/run")
        echo "$RESP" | python3 -m json.tool

        echo ""
        echo "Tests started. Poll with: $0 --results"
        echo "Or wait and poll automatically..."
        echo ""

        # Poll until done
        while true; do
            sleep 3
            STATUS=$(curl -s "$BASE_URL/status" 2>/dev/null || echo '{"running": true}')
            RUNNING=$(echo "$STATUS" | python3 -c "import sys,json; print(json.load(sys.stdin).get('running', True))")
            if [ "$RUNNING" = "False" ]; then
                break
            fi
            printf "."
        done
        echo ""
        echo ""

        # Print final results
        exec "$0" --results
        ;;
esac
