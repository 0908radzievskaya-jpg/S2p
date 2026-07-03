#!/usr/bin/env bash
set -Eeuo pipefail

REPO_URL="${REPO_URL:-https://github.com/0908radzievskaya-jpg/S2p.git}"
BRANCH="${BRANCH:-codex/s2p}"
APP_DIR="${APP_DIR:-/opt/tender-dashboard}"
APP_USER="${APP_USER:-tender-dashboard}"
APP_GROUP="${APP_GROUP:-tender-dashboard}"
ENV_DIR="${ENV_DIR:-/etc/tender-dashboard}"
ENV_FILE="${ENV_FILE:-$ENV_DIR/tender-dashboard.env}"
SITE="${TENDER_DASHBOARD_SITE:-194.113.209.237.sslip.io}"
ENABLE_UFW="${ENABLE_UFW:-1}"
ENABLE_AUTO_DEPLOY="${ENABLE_AUTO_DEPLOY:-1}"

log() {
    printf '[%s] %s\n' "$(date -Is)" "$*"
}

require_root() {
    if [ "$(id -u)" -ne 0 ]; then
        echo "Run this installer as root or through sudo." >&2
        exit 1
    fi
}

generate_password() {
    python3 - <<'PY'
import secrets
import string

alphabet = string.ascii_letters + string.digits + "-_"
print("".join(secrets.choice(alphabet) for _ in range(32)))
PY
}

escape_env_value() {
    python3 -c 'import sys; v=sys.argv[1]; print("\"" + v.replace("\\", "\\\\").replace("\"", "\\\"") + "\"")' "$1"
}

run_as_app() {
    runuser -u "$APP_USER" -- "$@"
}

write_caddyfile() {
    local caddyfile="/etc/caddy/Caddyfile"
    local tmp
    tmp="$(mktemp)"

    if [ -n "${ACME_EMAIL:-}" ]; then
        cat > "$tmp" <<EOF
{
    email $ACME_EMAIL
}

$SITE {
    encode zstd gzip
    header {
        X-Content-Type-Options "nosniff"
        Referrer-Policy "same-origin"
        Strict-Transport-Security "max-age=15552000"
    }
    reverse_proxy 127.0.0.1:8765
}
EOF
    else
        cat > "$tmp" <<EOF
$SITE {
    encode zstd gzip
    header {
        X-Content-Type-Options "nosniff"
        Referrer-Policy "same-origin"
        Strict-Transport-Security "max-age=15552000"
    }
    reverse_proxy 127.0.0.1:8765
}
EOF
    fi

    install -m 0644 "$tmp" "$caddyfile"
    rm -f "$tmp"
    caddy fmt --overwrite "$caddyfile"
    caddy validate --config "$caddyfile"
}

require_root

log "Updating apt cache and installing packages"
apt-get update
DEBIAN_FRONTEND=noninteractive apt-get install -y \
    build-essential \
    caddy \
    curl \
    git \
    python3 \
    python3-pip \
    python3-venv \
    ufw \
    unzip

if ! getent group "$APP_GROUP" >/dev/null; then
    groupadd --system "$APP_GROUP"
fi
if ! id -u "$APP_USER" >/dev/null 2>&1; then
    useradd --system --gid "$APP_GROUP" --home-dir "$APP_DIR" --shell /usr/sbin/nologin "$APP_USER"
fi

log "Preparing $APP_DIR"
mkdir -p "$(dirname "$APP_DIR")"
if [ -d "$APP_DIR/.git" ]; then
    chown -R "$APP_USER:$APP_GROUP" "$APP_DIR"
    run_as_app git -C "$APP_DIR" fetch --prune origin "$BRANCH"
    run_as_app git -C "$APP_DIR" checkout "$BRANCH"
    run_as_app git -C "$APP_DIR" merge --ff-only "origin/$BRANCH"
elif [ -e "$APP_DIR" ] && [ "$(find "$APP_DIR" -mindepth 1 -maxdepth 1 | wc -l)" -gt 0 ]; then
    echo "$APP_DIR exists and is not an empty Git checkout. Move it aside before installing." >&2
    exit 1
else
    git clone --branch "$BRANCH" --single-branch "$REPO_URL" "$APP_DIR"
fi
chown -R "$APP_USER:$APP_GROUP" "$APP_DIR"

log "Creating Python virtual environment"
run_as_app python3 -m venv "$APP_DIR/.venv"
run_as_app "$APP_DIR/.venv/bin/python" -m pip install --upgrade pip
if [ -f "$APP_DIR/requirements.txt" ]; then
    run_as_app "$APP_DIR/.venv/bin/python" -m pip install -r "$APP_DIR/requirements.txt"
fi

if [ ! -f "$APP_DIR/dashboard.server.json" ]; then
    cp "$APP_DIR/dashboard.server.example.json" "$APP_DIR/dashboard.server.json"
fi
chown "$APP_USER:$APP_GROUP" "$APP_DIR/dashboard.server.json"
chmod 0640 "$APP_DIR/dashboard.server.json"

log "Writing auth environment file"
install -d -m 0750 -o root -g "$APP_GROUP" "$ENV_DIR"
if [ ! -f "$ENV_FILE" ]; then
    dashboard_user="${TENDER_DASHBOARD_USER:-admin}"
    dashboard_password="${TENDER_DASHBOARD_PASSWORD:-$(generate_password)}"
    umask 0177
    cat > "$ENV_FILE" <<EOF
TENDER_DASHBOARD_USER=$(escape_env_value "$dashboard_user")
TENDER_DASHBOARD_PASSWORD=$(escape_env_value "$dashboard_password")
EOF
    chown root:"$APP_GROUP" "$ENV_FILE"
    chmod 0660 "$ENV_FILE"
else
    dashboard_user="$(grep -E '^TENDER_DASHBOARD_USER=' "$ENV_FILE" | cut -d= -f2- || true)"
    dashboard_password="existing password in $ENV_FILE"
    chown root:"$APP_GROUP" "$ENV_FILE"
    chmod 0660 "$ENV_FILE"
fi

install -d -m 0755 /var/log/tender-dashboard
chown "$APP_USER:$APP_GROUP" /var/log/tender-dashboard

log "Installing systemd units"
install -m 0644 "$APP_DIR/deploy/systemd/tender-dashboard.service" /etc/systemd/system/tender-dashboard.service
install -m 0644 "$APP_DIR/deploy/systemd/tender-dashboard-deploy.service" /etc/systemd/system/tender-dashboard-deploy.service
install -m 0644 "$APP_DIR/deploy/systemd/tender-dashboard-deploy.timer" /etc/systemd/system/tender-dashboard-deploy.timer
chmod +x "$APP_DIR/deploy/scripts/deploy_tender_dashboard.sh"

log "Configuring Caddy for $SITE"
write_caddyfile

systemctl daemon-reload
systemctl enable --now tender-dashboard.service
if [ "$ENABLE_AUTO_DEPLOY" = "1" ]; then
    systemctl enable --now tender-dashboard-deploy.timer
fi
systemctl enable --now caddy.service
systemctl reload caddy.service || systemctl restart caddy.service

if [ "$ENABLE_UFW" = "1" ]; then
    log "Configuring UFW"
    ufw allow 22/tcp
    ufw allow 80/tcp
    ufw allow 443/tcp
    ufw --force enable
fi

http_code="$(curl -sS -o /dev/null -w '%{http_code}' --max-time 10 http://127.0.0.1:8765/ || true)"
case "$http_code" in
    200|401)
        log "Local dashboard health check OK: HTTP $http_code"
        ;;
    *)
        log "Local dashboard health check failed: HTTP ${http_code:-000}"
        systemctl status tender-dashboard.service --no-pager || true
        journalctl -u tender-dashboard.service -n 80 --no-pager || true
        exit 1
        ;;
esac

log "Installation complete"
printf '\nTender Dashboard URL: https://%s/\n' "$SITE"
printf 'Install path: %s\n' "$APP_DIR"
printf 'Auth user: %s\n' "${dashboard_user:-admin}"
printf 'Auth password: %s\n' "$dashboard_password"
printf 'Auth env file: %s\n' "$ENV_FILE"
