# Logwatch for Windows

> *An independent Windows implementation inspired by the classic Linux `logwatch` utility.*

A lightweight, standalone daily monitoring and log reporting utility for Windows Server / Windows workstations...


* **Author:** KaydaulovMS
* **Version:** 1.0 (26.11.2026)
* **Languages:** Auto-detected via Windows UI locale (`ru` / `en`) or forced via config/CLI.

---

## Features

1. **System Telemetry & Top RAM Processes:**
   * OS version, uptime (highlighted if `< 24h`), RAM usage (highlighted if `> 90%`), pending reboot check via Windows Registry.
   * **Top 5 RAM Consumers:** Automatically displays the top 5 memory-consuming processes (`Win32_Process`) whenever RAM usage exceeds `85%`.
2. **Disk Status & Hardware Health:**
   * Local fixed drives (`DriveType=3`), free space in GB and `%`, threshold alerting.
   * **SMART & I/O Health:** Checks WMI `MSStorageDriver_FailurePredictStatus` (`PredictFailure`) and disk/controller/NTFS errors (`disk`, `Ntfs`, `storport`, `storahci`, `nvme`) in the `System` log.
3. **Services, Crashes & Task Scheduler:**
   * Lists `Auto`-start services that are not running (with configurable exclusion list).
   * **Newly Installed Services:** Tracks newly registered services and kernel drivers (`SCM EventID 7045`) with image path, start type, and logon account.
   * **Unexpected Reboots & Service Crashes:** Tracks unexpected power loss (`Kernel-Power 41`), dirty shutdowns (`EventLog 6008`), BSODs (`BugCheck 1001`), and unexpected service terminations (`SCM 7031 / 7034`).
   * **Task Scheduler Errors:** Monitors `Microsoft-Windows-TaskScheduler/Operational` for launch failures (`EventID 101`), action failures (`EventID 103`), and non-zero exit codes (`EventID 201`).
4. **Security & Access Control (`Security` Log):**
   * Successful logons & session unlocks (`EventID 4624` types 2, 7, 10, 11; `EventID 4801`).
   * Grouped failed logon attempts (`EventID 4625`) with source IP and logon type.
   * Account & audit tracking: user creation (`4720`), user deletion (`4726`), local group additions (`4732`), account lockouts (`4740`), and Security log clearing (`1102`).
5. **Software & OS Updates:** Tracks MSI installations (`MsiInstaller 11707`), Registry `Uninstall` entries (`InstallDate`), and Windows Update events (`WindowsUpdateClient 19`).
6. **Critical Errors Summary:** Groups `Critical` (Level 1) and `Error` (Level 2) events from `System` and `Application` logs by source and Event ID.
7. **Noise Protection & Execution Metrics:** Configurable row limit per table (`max_rows`) to prevent oversized emails, plus report generation timing in the footer.

---

## Quick Start & Deployment Layout

1. Download the precompiled **`logwatch-win-x64.exe`** from the [Releases](https://github.com/tehnomagistr/logwatch-for-windows/releases) page (or build it from source using the instructions below), rename it to **`logwatch.exe`**, and place it in a dedicated directory on your server (e.g., `C:\Program Files\Logwatch\`):
2. Place `logwatch.exe` in a dedicated directory on your server (e.g., `C:\Program Files\Logwatch\`):
   ```text
   C:\Program Files\Logwatch\
   ├── logwatch.exe          # Standalone executable (from dist/)
   ├── logwatch.json         # Auto-generated on first run next to the exe
   └── logwatch.log          # Created only if "enable_log": true
   ```
3. Run `logwatch.exe` once without a config file — it will automatically generate a default `logwatch.json` template next to the executable and exit.
4. Edit `logwatch.json` with your SMTP server and recipient addresses.
5. Test locally without sending an email:
   ```cmd
   logwatch.exe --test
   ```
   This creates `logwatch-report.html` in the same folder.
6. Schedule daily execution via Windows Task Scheduler (running as `SYSTEM` to access the `Security` event log):
   ```cmd
   schtasks /create /tn "Logwatch Daily Report" /tr "\"C:\Program Files\Logwatch\logwatch.exe\"" /sc daily /st 08:00 /ru "NT AUTHORITY\SYSTEM" /rl HIGHEST /f
   ```

---

## CLI Arguments

| Flag | Description |
| :--- | :--- |
| `--config <path>` | Path to a custom JSON configuration file. |
| `--test` | Generate `logwatch-report.html` locally without sending an email. |
| `--no-log` | Disable writing to `logwatch.log` for this run. |
| `--lang <auto\|ru\|en>` | Override report language (`auto`, `ru`, or `en`). |

---

## Configuration (`logwatch.json`)

```json
{
  "mail_from": "logwatch@example.com",
  "mail_to": "admin@example.com",
  "smtp_server": "192.0.2.10",
  "smtp_port": 25,
  "smtp_user": "",
  "smtp_password": "",
  "smtp_use_starttls": false,
  "smtp_use_ssl": false,
  "subject_prefix": "logwatch for",
  "lookback_days": 1,
  "threshold_percent": 15,
  "threshold_gb": 25,
  "ignored_services": [
    "sppsvc",
    "MapsBroker",
    "RemoteRegistry",
    "WbioSrvc",
    "CDPSvc",
    "edgeupdate",
    "gupdate",
    "TrustedInstaller"
  ],
  "ignored_tasks": [
    "Kontur",
    "EdgeUpdate",
    "Yandex",
    "Customer Experience Improvement"
  ],
  "log_file": "",
  "enable_log": false,
  "language": "auto",
  "max_rows": 50
}
```

* Set `"max_rows": 0` to disable table truncation and include all matching records.
* `ignored_tasks` filters out Task Scheduler errors by case-insensitive substring match against the task name.
* **Permissions:** Reading the Windows `Security` event log requires **Administrator** or **`NT AUTHORITY\SYSTEM`** privileges.


---

## Building from Source

The compiled binary (`dist/logwatch.exe`) is completely standalone and does not require Python on target servers.

### Requirements
* Python 3.10+ (Windows or Linux with Wine)
* Dependencies via `requirements.txt` (`pywin32`, `wmi`, `pyinstaller`)

### Build on Windows
```cmd
pip install -r requirements.txt
pyinstaller --onefile --noconsole --clean --name logwatch --icon logwatch.ico --version-file version_info.txt --hidden-import win32timezone --hidden-import win32com.client --hidden-import wmi --hidden-import pythoncom --hidden-import pywintypes logwatch.py
```


### Cross-Build on Linux (via Wine)
```bash
wine "C:\\Python312\\python.exe" -m PyInstaller \
    --onefile \
    --noconsole \
    --clean \
    --name logwatch \
    --icon logwatch.ico \
    --version-file version_info.txt \
    --hidden-import win32timezone \
    --hidden-import win32com.client \
    --hidden-import wmi \
    --hidden-import pythoncom \
    --hidden-import pywintypes \
    logwatch.py
```

The resulting standalone binary is created at **`dist/logwatch.exe`**.

---
---

# Logwatch для Windows (Русская версия)

> *Независимая реализация для Windows, вдохновлённая классической утилитой Linux `logwatch`.*

Автономная утилита ежедневного мониторинга и отправки HTML-отчётов по событиям Windows Server...


## Основные возможности

1. **Сводка по системе и Топ-5 процессов по RAM:**
   * Редакция ОС, время непрерывной работы (подсветка при `< 24 ч`), загрузка ОЗУ (подсветка при `> 90%`), проверка флагов ожидания перезагрузки в реестре.
   * **Топ-5 процессов по потреблению RAM:** автоматически выводится под общей сводкой, если занято более `85%` оперативной памяти.
2. **Состояние и аппаратное здоровье дисков:**
   * Проверка свободного места на локальных дисках с подсветкой при падении ниже порога (`%` или `ГБ`).
   * **SMART и ошибки ввода-вывода:** контроль прогноза отказа дисков через WMI (`MSStorageDriver_FailurePredictStatus`) и ошибок драйверов/ФС (`disk`, `Ntfs`, `storport`, `storahci`, `nvme`) в журнале `System`.
3. **Службы, сбои и Планировщик заданий:**
   * Контроль служб с типом запуска `Auto`, которые не находятся в состоянии `Running` (с поддержкой списка исключений).
   * **Новые зарегистрированные службы:** отслеживание установки новых служб и драйверов (`SCM EventID 7045`) с путём запуска, типом старта и учётной записью.
   * **Внезапные перезагрузки и падения служб:** отслеживание аварийного отключения питания (`Kernel-Power 41`), некорректного завершения работы (`EventLog 6008`), BSOD (`BugCheck 1001`) и падений служб (`SCM 7031 / 7034`).
   * **Ошибки Планировщика заданий:** контроль журнала `TaskScheduler/Operational` на ошибки запуска (`ID 101`), сбои действий (`ID 103`) и ненулевые коды возврата (`ID 201`).
4. **Контроль доступа и учётных записей (журнал `Security`):**
   * Успешные входы (консоль, RDP, разблокировка, кэшированный вход — ID `4624`, `4801`).
   * Сводка неудачных попыток входа (`ID 4625`) с группировкой по учётной записи, IP-адресу и типу входа.
   * Изменения учётных записей и аудит: создание пользователей (`4720`), удаление пользователей (`4726`), добавление в локальные группы (`4732`), блокировки УЗ (`4740`) и очистка журнала безопасности (`1102`).
5. **Установленное ПО и обновления ОС:** отслеживание MSI-установок (`MsiInstaller 11707`), записей реестра `Uninstall` и обновлений Windows (`WindowsUpdateClient 19`).
6. **Сводка критических ошибок:** группировка событий уровня `Critical` и `Error` из журналов `System` и `Application`.
7. **Ограничение объёма и футер метрик:** настраиваемый лимит строк в таблицах (`max_rows`) для защиты от слишком крупных писем и вывод времени генерации отчёта в футере.

## Быстрый старт и структура файлов на сервере

1. Скачайте готовый **`logwatch-win-x64.exe`** со страницы [Releases](https://github.com/tehnomagistr/logwatch-for-windows/releases) (или соберите его самостоятельно по инструкции ниже), переименуйте в **`logwatch.exe`** и скопируйте в рабочую папку на сервере (например, `C:\Program Files\Logwatch\`):
2. Скопируйте `logwatch.exe` в рабочую папку на сервере (например, `C:\Program Files\Logwatch\`):
   ```text
   C:\Program Files\Logwatch\
   ├── logwatch.exe          # Исполняемый файл (из папки dist/)
   ├── logwatch.json         # Создаётся автоматически при первом запуске (рядом с exe)
   └── logwatch.log          # Создаётся только если в конфиге включено "enable_log": true
   ```
3. Запустите `logwatch.exe` первый раз — программа сама создаст шаблон `logwatch.json` с настройками по умолчанию и завершит работу.
4. Укажите адрес SMTP-сервера и получателей в `logwatch.json`.
5. Для локальной проверки без отправки почты выполните:
   ```cmd
   logwatch.exe --test
   ```
6. Зарегистрируйте ежедневную задачу в Планировщике Windows (от имени `SYSTEM` для доступа к журналу `Security`):
   ```cmd
   schtasks /create /tn "Logwatch Daily Report" /tr "\"C:\Program Files\Logwatch\logwatch.exe\"" /sc daily /st 08:00 /ru "NT AUTHORITY\SYSTEM" /rl HIGHEST /f
   ```

## Сборка из исходников

Собранный исполняемый файл (`dist/logwatch.exe`) полностью автономен и не требует установленного Python на целевых серверах.

### Зависимости
* Python 3.10+ (на Windows или на Linux через Wine)
* Зависимости из `requirements.txt` (`pywin32`, `wmi`, `pyinstaller`)

### Сборка на Windows
```cmd
pip install -r requirements.txt
pyinstaller --onefile --noconsole --clean --name logwatch --icon logwatch.ico --version-file version_info.txt --hidden-import win32timezone --hidden-import win32com.client --hidden-import wmi --hidden-import pythoncom --hidden-import pywintypes logwatch.py
```

### Кросс-сборка на Linux (через Wine)
```bash
wine "C:\\Python312\\python.exe" -m PyInstaller \
    --onefile \
    --noconsole \
    --clean \
    --name logwatch \
    --icon logwatch.ico \
    --version-file version_info.txt \
    --hidden-import win32timezone \
    --hidden-import win32com.client \
    --hidden-import wmi \
    --hidden-import pythoncom \
    --hidden-import pywintypes \
    logwatch.py
```

Готовый файл появится по пути **`dist/logwatch.exe`**.

---

## License

This project is licensed under the [MIT License](LICENSE).

