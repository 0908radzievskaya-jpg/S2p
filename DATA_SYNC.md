# Автообновление данных Tender Dashboard

Код панели обновляется на VPS через `tender-dashboard-deploy.timer`. Данные панели не хранятся в Git: их нужно синхронизировать отдельно из локальных папок `reports` и `релевантные`.

## Ручная синхронизация

По умолчанию работает режим `Minimal`: на сервер отправляются только сводные Excel-файлы заявок, JSON-файлы анализа и документы, похожие на ТЗ, смету или аналитическую записку. Полная папка `релевантные` и архивы материалов не выгружаются.

Из PowerShell на Windows:

```powershell
.\deploy\scripts\sync_dashboard_data.ps1
```

Если для SSH используется отдельный ключ:

```powershell
.\deploy\scripts\sync_dashboard_data.ps1 -SshKeyPath "$env:USERPROFILE\.ssh\id_ed25519"
```

Скрипт упакует минимальный набор файлов из локальных папок `reports` и `релевантные`, загрузит его на `root@194.113.209.237`, заменит серверные папки в `/opt/tender-dashboard`, сохранит backup в `/var/backups/tender-dashboard-data/`, поправит права и проверит `http://127.0.0.1:8765/`.

Перед упаковкой скрипт скачивает с сервера `.state/dashboard_statuses.json`. Если в dashboard заявка отмечена как `Не подходит`, локально удаляются связанные файлы только внутри разрешенных папок проекта: `reports`, `релевантные`, `sorted_mail`, `_review_ai_tender`, `archive`. Служебные JSON-файлы анализа сохраняются, чтобы строка dashboard не появилась заново с другим ID.

Отключить локальную очистку для разового запуска:

```powershell
.\deploy\scripts\sync_dashboard_data.ps1 -ApplyRemoteDeletions:$false
```

Для разовой полной выгрузки всех данных:

```powershell
.\deploy\scripts\sync_dashboard_data.ps1 -Mode Full
```

## Автоматическая синхронизация

Установить задачу Windows Task Scheduler на ежедневный запуск в 09:00:

```powershell
.\deploy\scripts\install_windows_data_sync_task.ps1 -DailyAt 09:00 -RunNow
```

С отдельным SSH-ключом:

```powershell
.\deploy\scripts\install_windows_data_sync_task.ps1 -DailyAt 09:00 -SshKeyPath "$env:USERPROFILE\.ssh\id_ed25519" -RunNow
```

После этого Windows будет отправлять свежие данные каждый день в 09:00 без ручного участия. Если предыдущая синхронизация еще выполняется, новый запуск будет пропущен; максимальное время одного запуска по умолчанию — 120 минут.

## Требования

- На Windows должны быть доступны `ssh`, `scp` и `tar`.
- Для полностью автоматической работы нужен SSH-доступ без ввода пароля: ключ в `~/.ssh` или ключ, переданный через `-SshKeyPath`.
- На VPS должен быть установлен Tender Dashboard через `deploy/scripts/install_tender_dashboard.sh`.

## Управление задачей

```powershell
Get-ScheduledTask -TaskName "Tender Dashboard Data Sync"
Start-ScheduledTask -TaskName "Tender Dashboard Data Sync"
Unregister-ScheduledTask -TaskName "Tender Dashboard Data Sync"
```
