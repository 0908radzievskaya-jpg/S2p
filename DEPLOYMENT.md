# Развертывание Tender Dashboard на Ubuntu 24.04

Рекомендуемый путь установки: `/opt/tender-dashboard`.

## Быстрая установка

Подключитесь к VPS по SSH и выполните:

```bash
apt-get update
apt-get install -y git
git clone --branch codex/s2p --single-branch https://github.com/0908radzievskaya-jpg/S2p.git /tmp/tender-dashboard-src
cd /tmp/tender-dashboard-src
TENDER_DASHBOARD_USER=admin \
TENDER_DASHBOARD_PASSWORD='replace-with-long-password' \
TENDER_DASHBOARD_SITE=194.113.209.237.sslip.io \
bash deploy/scripts/install_tender_dashboard.sh
```

`194.113.209.237.sslip.io` можно заменить на свой домен после создания DNS A-записи на `194.113.209.237`. Caddy сам получит и продлит Let's Encrypt сертификат.

## Что устанавливает скрипт

- пакеты: `git`, `curl`, `unzip`, `python3`, `python3-venv`, `python3-pip`, `caddy`, `build-essential`, `ufw`;
- проект из `https://github.com/0908radzievskaya-jpg/S2p.git`, ветка `codex/s2p`;
- Python venv: `/opt/tender-dashboard/.venv`;
- server-only конфиг: `/opt/tender-dashboard/dashboard.server.json`;
- секреты Basic Auth: `/etc/tender-dashboard/tender-dashboard.env`;
- systemd-сервис: `tender-dashboard.service`;
- systemd-таймер автообновления: `tender-dashboard-deploy.timer`;
- reverse proxy Caddy на `127.0.0.1:8765`;
- UFW: открыты только `22/tcp`, `80/tcp`, `443/tcp`.

## Команды эксплуатации

```bash
systemctl status tender-dashboard --no-pager
journalctl -u tender-dashboard -f
systemctl status caddy --no-pager
journalctl -u caddy -f
curl -I http://127.0.0.1:8765/
```

Автообновление данных из локальных папок `reports` и `релевантные` описано в `DATA_SYNC.md`.

Ручное обновление:

```bash
/opt/tender-dashboard/deploy/scripts/deploy_tender_dashboard.sh
```

Проверка автообновления:

```bash
systemctl list-timers tender-dashboard-deploy.timer
journalctl -u tender-dashboard-deploy -n 100 --no-pager
```

## Замена конфигурации

Не используйте `config.json` на сервере, если в нем есть IMAP/SMTP/API-секреты.

Редактируйте только серверный конфиг:

```bash
nano /opt/tender-dashboard/dashboard.server.json
systemctl restart tender-dashboard
```

Basic Auth и интеграции можно менять в dashboard через кнопку `Интеграции`. Секреты сохраняются в environment-файл:

```bash
nano /etc/tender-dashboard/tender-dashboard.env
systemctl restart tender-dashboard
```

Резервная интерактивная команда без браузера:

```bash
/opt/tender-dashboard/.venv/bin/python /opt/tender-dashboard/tender_dashboard.py configure-secrets --config /opt/tender-dashboard/dashboard.server.json
systemctl restart tender-dashboard
```

Файл `/opt/tender-dashboard/dashboard.server.json` исключен из Git. Файл `/etc/tender-dashboard/tender-dashboard.env` находится вне репозитория.

## Как работает автообновление

`tender-dashboard-deploy.timer` каждые 2 минуты проверяет `origin/codex/s2p`. Если появился новый commit, скрипт:

1. сохраняет резервную копию `dashboard.server.json` и `/etc/tender-dashboard/tender-dashboard.env` в `/var/backups/tender-dashboard/<timestamp>/`;
2. делает fast-forward merge ветки `codex/s2p`;
3. обновляет зависимости из `requirements.txt`;
4. компилирует ключевые Python-файлы;
5. перезапускает `tender-dashboard.service`;
6. проверяет `http://127.0.0.1:8765/`, считая `200` и `401` здоровым ответом.

## Важные файлы

- `/opt/tender-dashboard` — проект;
- `/opt/tender-dashboard/dashboard.server.json` — серверная конфигурация без почтовых секретов;
- `/etc/tender-dashboard/tender-dashboard.env` — Basic Auth;
- `/etc/systemd/system/tender-dashboard.service` — автозапуск dashboard;
- `/etc/caddy/Caddyfile` — HTTPS reverse proxy;
- `/var/backups/tender-dashboard/` — резервные копии перед деплоем.
