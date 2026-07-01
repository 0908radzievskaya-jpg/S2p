#!/usr/bin/env bash
set -Eeuo pipefail

KEY='ssh-ed25519 AAAAC3NzaC1lZDI1NTE5AAAAIDXKfW/ZT01Lt+MeFQ9CfbZTcKMS4TYwMrf7pDpQnhFb tender-dashboard-sync-windows'
AUTHORIZED_KEYS="/root/.ssh/authorized_keys"

if [ "$(id -u)" -ne 0 ]; then
    echo "Run this script as root." >&2
    exit 1
fi

install -d -m 700 /root/.ssh
touch "$AUTHORIZED_KEYS"
sed -i '/tender-dashboard-sync-windows/d' "$AUTHORIZED_KEYS"
printf '%s\n' "$KEY" >> "$AUTHORIZED_KEYS"
chmod 600 "$AUTHORIZED_KEYS"

echo "Installed Windows dashboard sync SSH key for root."
