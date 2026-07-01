#!/usr/bin/env bash
set -Eeuo pipefail

APP_DIR="${APP_DIR:-/opt/tender-dashboard}"
APP_USER="${APP_USER:-tender-dashboard}"
BRANCH="${BRANCH:-codex/s2p}"
HEALTH_URL="${HEALTH_URL:-http://127.0.0.1:8765/}"
BACKUP_DIR="${BACKUP_DIR:-/var/backups/tender-dashboard}"
ENV_FILE="${ENV_FILE:-/etc/tender-dashboard/tender-dashboard.env}"

log() {
    printf '[%s] %s\n' "$(date -Is)" "$*"
}

run_as_app() {
    runuser -u "$APP_USER" -- "$@"
}

if [ ! -d "$APP_DIR/.git" ]; then
    log "Missing Git checkout at $APP_DIR"
    exit 1
fi

cd "$APP_DIR"

current_revision="$(run_as_app git -C "$APP_DIR" rev-parse HEAD)"
run_as_app git -C "$APP_DIR" fetch --prune origin "$BRANCH"
remote_revision="$(run_as_app git -C "$APP_DIR" rev-parse "origin/$BRANCH")"

if [ "$current_revision" = "$remote_revision" ]; then
    log "No changes on origin/$BRANCH"
    exit 0
fi

timestamp="$(date +%Y%m%d-%H%M%S)"
mkdir -p "$BACKUP_DIR/$timestamp"
if [ -f "$APP_DIR/dashboard.server.json" ]; then
    cp "$APP_DIR/dashboard.server.json" "$BACKUP_DIR/$timestamp/dashboard.server.json"
fi
if [ -f "$ENV_FILE" ]; then
    cp "$ENV_FILE" "$BACKUP_DIR/$timestamp/tender-dashboard.env"
fi
chmod -R go-rwx "$BACKUP_DIR/$timestamp"

log "Deploying $current_revision -> $remote_revision"
run_as_app git -C "$APP_DIR" checkout "$BRANCH"
run_as_app git -C "$APP_DIR" merge --ff-only "origin/$BRANCH"

if [ -f "$APP_DIR/requirements.txt" ]; then
    run_as_app "$APP_DIR/.venv/bin/python" -m pip install --upgrade -r "$APP_DIR/requirements.txt"
fi

run_as_app "$APP_DIR/.venv/bin/python" -m py_compile "$APP_DIR/tender_dashboard.py" "$APP_DIR/mail_automation.py"

systemctl restart tender-dashboard.service
sleep 3

http_code="$(curl -sS -o /dev/null -w '%{http_code}' --max-time 10 "$HEALTH_URL" || true)"
case "$http_code" in
    200|401)
        log "Health check OK: HTTP $http_code"
        ;;
    *)
        log "Health check failed: HTTP ${http_code:-000}"
        systemctl status tender-dashboard.service --no-pager || true
        journalctl -u tender-dashboard.service -n 80 --no-pager || true
        exit 1
        ;;
esac

systemctl reload caddy.service || systemctl restart caddy.service
log "Deploy complete"
