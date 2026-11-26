#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import sys
import traceback
from pathlib import Path

base_dir = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).parent
startup_log = base_dir / "startup.log"


def log_startup(msg):
    pass


try:
    import time
    import argparse
    import ctypes
    import html
    import json
    import locale
    import logging
    import os
    import smtplib
    import socket
    import xml.etree.ElementTree as ET
    from collections import defaultdict
    from datetime import datetime, timedelta, timezone
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText

    import win32evtlog
    import winreg
    import wmi
except BaseException as e:
    with open(startup_log, "a", encoding="utf-8") as f:
        f.write(f"CRITICAL IMPORT ERROR: {e}\n{traceback.format_exc()}\n")
    sys.exit(1)

NS = {"e": "http://schemas.microsoft." + "com/win/2004/08/events/event"}


DEFAULT_CONFIG = {
    "mail_from": "logwatch@localhost",
    "mail_to": "admin@localhost",
    "smtp_server": "localhost",
    "smtp_port": 25,
    "smtp_user": "",
    "smtp_password": "",
    "smtp_use_starttls": False,
    "smtp_use_ssl": False,
    "subject_prefix": "logwatch for",
    "lookback_days": 1,
    "threshold_percent": 15,
    "threshold_gb": 25,
    "ignored_services": [
        "sppsvc", "MapsBroker", "RemoteRegistry", "WbioSrvc", "CDPSvc",
        "edgeupdate", "gupdate", "TrustedInstaller",
    ],
    "ignored_tasks": [
        "Kontur",
        "EdgeUpdate",
        "Yandex",
        "Customer Experience Improvement",
    ],
    "log_file": "",
    "enable_log": True,
    "language": "auto",
    "max_rows": 50,
}

STRINGS = {
    "ru": {
        "html_lang": "ru",
        "report_title": "Ежедневный отчёт сервера",
        "os": "ОС",
        "uptime": "Uptime",
        "uptime_fmt": "{d} дн. {h} ч. {m} мин. (с {boot})",
        "ram": "RAM",
        "ram_fmt": "Занято {used}% (Свободно {free} из {total} ГБ)",
        "reboot": "Перезагрузка",
        "reboot_req": "ТРЕБУЕТСЯ ПЕРЕЗАГРУЗКА",
        "reboot_ok": "Не требуется",
        "sec_disks": "1. Состояние дисков",
        "th_disk": "Диск",
        "th_label": "Метка",
        "th_total_gb": "Всего (ГБ)",
        "th_free_gb": "Свободно (ГБ)",
        "th_free_pct": "Свободно (%)",
        "sec_services": "2. Остановленные авто-службы",
        "th_service": "Служба",
        "th_name": "Имя",
        "th_status": "Статус",
        "services_ok": "Все автоматические службы работают.",
        "sec_security": "3. Контроль доступа и учетных записей",
        "sub_success": "Успешные входы и разблокировки",
        "th_time": "Время",
        "th_event": "Событие",
        "th_account": "Учетная запись",
        "th_ip": "IP-адрес",
        "no_success": "Событий не зафиксировано.",
        "sub_failed": "Неудачные попытки входа (ID 4625)",
        "th_attempts": "Попыток",
        "th_logon_type": "Тип входа",
        "no_failed": "Неудачных попыток не зафиксировано.",
        "sub_accounts": "Изменения учетных записей и групп",
        "th_code": "Код",
        "th_details": "Детали",
        "no_accounts": "Изменений нет.",
        "sec_software": "4. Установленное ПО и обновления ОС",
        "th_type": "Тип",
        "th_desc": "Описание",
        "no_software": "Новых программ и обновлений нет.",
        "sec_critical": "5. Сводка критических ошибок",
        "th_count": "Кол-во",
        "th_last_fail": "Последний сбой",
        "th_log": "Журнал",
        "th_source": "Источник",
        "th_msg": "Сообщение",
        "no_critical": "Критических ошибок не зафиксировано.",
        "logon_2": "Вход в консоль",
        "logon_7": "Разблокировка сессии",
        "logon_10": "Вход по RDP",
        "logon_11": "Кэшированный вход",
        "unlock_4801": "Разблокировка (ID 4801)",
        "local": "Локально",
        "fail_2": "Консоль",
        "fail_3": "SMB",
        "fail_7": "Разблокировка",
        "fail_10": "RDP",
        "fail_11": "Кэш",
        "fail_type": "Тип {lt}",
        "not_specified": "Не указан",
        "acc_created": "Создан пользователь",
        "acc_group": "Добавление в локальную группу",
        "acc_fallback": "Учетная запись: {user}",
        "sw_msi": "ПО (MSI)",
        "sw_msi_fallback": "Установка MSI-пакета",
        "sw_reg": "ПО (Реестр)",
        "sw_wu": "Обновление ОС",
        "sw_wu_fallback": "Установлено обновление Windows",
        "no_msg_text": "(Нет текста сообщения)",
        "sec_crashes": "2.1. Внезапные перезагрузки и падения служб",
        "no_crashes": "Внезапных перезагрузок и падений служб не зафиксировано.",
        "crash_41": "Внезапное отключение питания / зависание (Kernel-Power 41)",
        "crash_6008": "Неожиданное завершение работы (EventLog 6008)",
        "crash_1001": "Синий экран смерти / BSOD (BugCheck 1001)",
        "crash_svc": "Аварийное падение службы (SCM {id})",
        "sec_disk_health": "1.1. Аппаратное здоровье дисков (SMART и ошибки I/O)",
        "no_disk_health": "Ошибок дисков, контроллеров и предупреждений SMART не зафиксировано.",
        "smart_warn": "SMART: Прогнозируется отказ диска!",
        "disk_io_err": "Ошибка диска / ФС ({provider} ID {id})",
        "sec_tasks": "2.2. Ошибки Планировщика заданий",
        "no_tasks": "Ошибок выполнения задач в Планировщике не зафиксировано.",
        "th_task_name": "Задача",
        "task_err_101": "Ошибка запуска (ID 101)",
        "task_err_103": "Сбой действия (ID 103)",
        "task_err_201": "Код возврата {code} (ID 201)",
        "acc_deleted": "Удален пользователь",
        "acc_locked": "Блокировка учетной записи (Lockout)",
        "sec_cleared": "КРИТИЧНО: Очищен журнал безопасности!",
        "sub_new_services": "Новые зарегистрированные службы (ID 7045)",
        "no_new_services": "Новых служб не зарегистрировано.",
        "th_bin_path": "Путь / Команда запуска",
        "th_start_type": "Тип запуска",
        "sub_top_ram": "Топ-5 процессов по потреблению RAM",
        "th_proc_name": "Процесс",
        "th_pid": "PID",
        "th_ram_mb": "Память (МБ)",
        "more_rows": "... и ещё {count} записей",
        "footer_fmt": "Сгенерировано за {sec} сек. | Период: {days} дн. | Logwatch for Windows v1.0",
    },
    "en": {
        "html_lang": "en",
        "report_title": "Daily Server Report",
        "os": "OS",
        "uptime": "Uptime",
        "uptime_fmt": "{d}d {h}h {m}m (since {boot})",
        "ram": "RAM",
        "ram_fmt": "Used {used}% (Free {free} of {total} GB)",
        "reboot": "Reboot",
        "reboot_req": "REBOOT REQUIRED",
        "reboot_ok": "Not required",
        "sec_disks": "1. Disk Status",
        "th_disk": "Drive",
        "th_label": "Label",
        "th_total_gb": "Total (GB)",
        "th_free_gb": "Free (GB)",
        "th_free_pct": "Free (%)",
        "sec_services": "2. Stopped Automatic Services",
        "th_service": "Service",
        "th_name": "Display Name",
        "th_status": "Status",
        "services_ok": "All automatic services are running.",
        "sec_security": "3. Access and Account Control",
        "sub_success": "Successful Logons and Unlocks",
        "th_time": "Time",
        "th_event": "Event",
        "th_account": "Account",
        "th_ip": "IP Address",
        "no_success": "No events recorded.",
        "sub_failed": "Failed Logon Attempts (ID 4625)",
        "th_attempts": "Attempts",
        "th_logon_type": "Logon Type",
        "no_failed": "No failed attempts recorded.",
        "sub_accounts": "Account and Group Changes",
        "th_code": "Code",
        "th_details": "Details",
        "no_accounts": "No changes recorded.",
        "sec_software": "4. Installed Software and OS Updates",
        "th_type": "Type",
        "th_desc": "Description",
        "no_software": "No new software or updates.",
        "sec_critical": "5. Critical Errors Summary",
        "th_count": "Count",
        "th_last_fail": "Last Failure",
        "th_log": "Log",
        "th_source": "Source",
        "th_msg": "Message",
        "no_critical": "No critical errors recorded.",
        "logon_2": "Console Logon",
        "logon_7": "Session Unlock",
        "logon_10": "RDP Logon",
        "logon_11": "Cached Logon",
        "unlock_4801": "Unlock (ID 4801)",
        "local": "Local",
        "fail_2": "Console",
        "fail_3": "SMB",
        "fail_7": "Unlock",
        "fail_10": "RDP",
        "fail_11": "Cached",
        "fail_type": "Type {lt}",
        "not_specified": "Not specified",
        "acc_created": "User Created",
        "acc_group": "Added to Local Group",
        "acc_fallback": "Account: {user}",
        "sw_msi": "Software (MSI)",
        "sw_msi_fallback": "MSI package installation",
        "sw_reg": "Software (Registry)",
        "sw_wu": "OS Update",
        "sw_wu_fallback": "Windows update installed",
        "no_msg_text": "(No message text)",
        "sec_crashes": "2.1. Unexpected Reboots and Service Crashes",
        "no_crashes": "No unexpected reboots or service crashes recorded.",
        "crash_41": "Unexpected Power Loss / Hang (Kernel-Power 41)",
        "crash_6008": "Unexpected Shutdown (EventLog 6008)",
        "crash_1001": "Blue Screen of Death / BSOD (BugCheck 1001)",
        "crash_svc": "Service Crash (SCM {id})",
        "sec_disk_health": "1.1. Disk Hardware Health (SMART & I/O Errors)",
        "no_disk_health": "No disk, controller errors or SMART warnings recorded.",
        "smart_warn": "SMART: Disk failure predicted!",
        "disk_io_err": "Disk / FS Error ({provider} ID {id})",
        "sec_tasks": "2.2. Task Scheduler Errors",
        "no_tasks": "No scheduled task errors recorded.",
        "th_task_name": "Task Name",
        "task_err_101": "Launch Failed (ID 101)",
        "task_err_103": "Action Failed (ID 103)",
        "task_err_201": "Exit Code {code} (ID 201)",
        "acc_deleted": "User Deleted",
        "acc_locked": "Account Locked Out",
        "sec_cleared": "CRITICAL: Security Log Cleared!",
        "sub_new_services": "Newly Installed Services (ID 7045)",
        "no_new_services": "No new services registered.",
        "th_bin_path": "Image Path / Command",
        "th_start_type": "Start Type",
        "sub_top_ram": "Top 5 Processes by RAM Usage",
        "th_proc_name": "Process",
        "th_pid": "PID",
        "th_ram_mb": "Memory (MB)",
        "more_rows": "... and {count} more records",
        "footer_fmt": "Generated in {sec}s | Period: {days}d | Logwatch for Windows v1.0",
    },
}


def tr(lang: str, key: str, **kwargs) -> str:
    table = STRINGS.get(lang, STRINGS["en"])
    text = table.get(key, STRINGS["en"].get(key, key))
    return text.format(**kwargs) if kwargs else text


def app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parent


def load_config(cli_path: str | None) -> tuple[dict, bool]:
    cfg = dict(DEFAULT_CONFIG)
    path = Path(cli_path) if cli_path else app_dir() / "logwatch.json"
    if path.exists():
        with open(path, "r", encoding="utf-8") as f:
            cfg.update(json.load(f))
        return cfg, False

    if not cli_path:
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(DEFAULT_CONFIG, f, indent=2, ensure_ascii=False)
        except Exception as e:
            with open(startup_log, "a", encoding="utf-8") as f:
                f.write(f"Failed to create default config: {e}\n")
    return cfg, True



def detect_language(cfg_lang: str) -> str:
    if cfg_lang and cfg_lang.lower() in ("ru", "en"):
        return cfg_lang.lower()
    try:
        lang_id = ctypes.windll.kernel32.GetUserDefaultUILanguage()
        if (lang_id & 0xFF) == 0x19:
            return "ru"
    except Exception:
        pass
    loc = (locale.getlocale()[0] or "").lower()
    return "ru" if loc.startswith("ru") else "en"


def setup_logging(cfg: dict) -> None:
    if not cfg.get("enable_log", True):
        logging.basicConfig(handlers=[logging.NullHandler()], force=True)
        logging.disable(logging.CRITICAL)
        return

    log_path = Path(cfg["log_file"]) if cfg.get("log_file") else app_dir() / "logwatch.log"
    try:
        log_path.parent.mkdir(parents=True, exist_ok=True)
        logging.basicConfig(
            filename=str(log_path),
            level=logging.INFO,
            format="%(asctime)s %(levelname)s %(message)s",
            encoding="utf-8",
            force=True,
        )
    except Exception:
        logging.basicConfig(level=logging.INFO, force=True)


_publisher_cache: dict[str, object] = {}


def _format_message(ev_handle, provider: str) -> str:
    if not provider:
        return ""
    try:
        pub = _publisher_cache.get(provider)
        if pub is None:
            pub = win32evtlog.EvtOpenPublisherMetadata(provider)
            _publisher_cache[provider] = pub
        return win32evtlog.EvtFormatMessage(
            pub, ev_handle, win32evtlog.EvtFormatMessageEvent
        )
    except Exception:
        return ""


def _parse_event_xml(xml_str: str, log_name: str) -> dict:
    root = ET.fromstring(xml_str)
    sys_el = root.find("e:System", NS)
    provider = ""
    p_el = sys_el.find("e:Provider", NS)
    if p_el is not None:
        provider = p_el.get("Name", "")
    level = 0
    l_el = sys_el.find("e:Level", NS)
    if l_el is not None and l_el.text:
        try:
            level = int(l_el.text)
        except ValueError:
            pass
    tc = sys_el.find("e:TimeCreated", NS)
    dt = datetime.min
    if tc is not None and tc.get("SystemTime"):
        dt_utc = datetime.strptime(tc.get("SystemTime")[:19], "%Y-%m-%dT%H:%M:%S")
        dt = dt_utc.replace(tzinfo=timezone.utc).astimezone().replace(tzinfo=None)
    data = {}
    for d in root.findall(".//e:EventData/e:Data", NS):
        name = d.get("Name")
        if name:
            data[name] = d.text or ""
    return {
        "id": int(sys_el.find("e:EventID", NS).text),
        "time": dt,
        "provider": provider,
        "log": log_name,
        "level": level,
        "data": data,
    }


STYLE = """
<style>
body { font-family: Segoe UI, Arial, sans-serif; font-size: 13px; color: #222; }
h2 { color: #1a237e; border-bottom: 2px solid #1a237e; padding-bottom: 5px; }
h3 { color: #283593; margin-top: 24px; margin-bottom: 8px; }
h4 { color: #455a64; margin-top: 14px; margin-bottom: 6px; }
p  { margin: 6px 0 12px 0; }
table { border-collapse: collapse; width: 100%; margin-bottom: 16px; }
th, td { border: 1px solid #cccccc; padding: 6px 10px; text-align: left; white-space: nowrap; }
td:last-child { white-space: normal; }
th { background-color: #f0f2f5; }
.summary-table th { width: 18%; background-color: #e8eaf6; }
.summary-table td { width: 32%; white-space: normal; }
.ok { color: #2e7d32; }
.alert { background-color: #ffe6e6; color: #b30000; font-weight: bold; }
.red   { color: #b30000; font-weight: bold; }
.green { color: #2e7d32; }
.footer { margin-top: 24px; padding-top: 8px; border-top: 1px solid #cccccc; font-size: 11px; color: #666666; }
</style>
"""


def esc(val) -> str:
    return html.escape(str(val or ""))


def query_events(log_name: str, xpath: str, with_message=False, max_events=50000) -> list[dict]:
    flags = win32evtlog.EvtQueryChannelPath | win32evtlog.EvtQueryReverseDirection
    try:
        h = win32evtlog.EvtQuery(log_name, flags, xpath)
    except Exception as e:
        logging.warning("EvtQuery failed [%s]: %s", log_name, e)
        return []
    out: list[dict] = []
    while len(out) < max_events:
        try:
            batch = win32evtlog.EvtNext(h, 100)
        except Exception:
            break
        if not batch:
            break
        for ev_handle in batch:
            try:
                xml_str = win32evtlog.EvtRender(ev_handle, win32evtlog.EvtRenderEventXml)
                rec = _parse_event_xml(xml_str, log_name)
                if with_message:
                    rec["message"] = _format_message(ev_handle, rec["provider"])
                out.append(rec)
            except Exception:
                continue
    return out


def xpath_since(dt: datetime) -> str:
    return dt.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S.000Z")
def collect_telemetry(c: wmi.WMI) -> dict:
    os_info = c.Win32_OperatingSystem()[0]
    last_boot = os_info.LastBootUpTime
    if isinstance(last_boot, str):
        last_boot = datetime.strptime(last_boot.split(".")[0], "%Y%m%d%H%M%S")
    uptime = datetime.now() - last_boot

    total_mb = int(os_info.TotalVisibleMemorySize) / 1024
    free_mb = int(os_info.FreePhysicalMemory) / 1024
    used_pct = round((total_mb - free_mb) / total_mb * 100, 1)

    reboot = False
    for path in (
        r"SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending",
        r"SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired",
    ):
        try:
            winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, path)
            reboot = True
            break
        except FileNotFoundError:
            pass

    return {
        "caption": os_info.Caption,
        "last_boot": last_boot,
        "uptime": uptime,
        "total_gb": round(total_mb / 1024, 2),
        "free_gb": round(free_mb / 1024, 2),
        "used_pct": used_pct,
        "reboot_pending": reboot,
    }

def collect_top_ram_processes(c: wmi.WMI, limit: int = 5) -> list[dict]:
    procs = []
    for p in c.Win32_Process():
        try:
            ws = int(p.WorkingSetSize or 0)
            if ws > 0:
                procs.append({
                    "name": p.Name or "Unknown",
                    "pid": p.ProcessId,
                    "ram_mb": round(ws / (1024 ** 2), 1),
                })
        except Exception:
            continue
    procs.sort(key=lambda x: -x["ram_mb"])
    return procs[:limit]


def collect_disks(c: wmi.WMI, threshold_percent: float, threshold_gb: float) -> list[dict]:
    rows = []
    for d in c.Win32_LogicalDisk(DriveType=3):
        if not d.Size:
            continue
        size = int(d.Size)
        free_space = int(d.FreeSpace)
        total_gb = round(size / (1024 ** 3), 2)
        free_gb = round(free_space / (1024 ** 3), 2)
        free_pct = round(free_space / size * 100, 1)
        alert = (free_pct < threshold_percent) or (free_gb < threshold_gb)
        rows.append({
            "device": d.DeviceID,
            "label": d.VolumeName or "",
            "total_gb": total_gb,
            "free_gb": free_gb,
            "free_pct": free_pct,
            "alert": alert,
        })
    return rows

def collect_disk_health(since: datetime, lang: str) -> list[dict]:
    rows = []

    try:
        wmi_storage = wmi.WMI(namespace="root\\wmi")
        for item in wmi_storage.MSStorageDriver_FailurePredictStatus():
            if getattr(item, "PredictFailure", False):
                inst = getattr(item, "InstanceName", "Unknown Disk")
                reason = getattr(item, "Reason", 0)
                rows.append({
                    "time": datetime.now().strftime("%d.%m.%Y %H:%M"),
                    "type": tr(lang, "smart_warn"),
                    "details": f"{inst} (Reason: {reason})",
                })
    except Exception:
        pass

    xpath = (
        f"*[System["
        f"(Provider[@Name='disk' or @Name='Ntfs' or @Name='storport' or @Name='storahci' or @Name='nvme'] "
        f"and (Level=1 or Level=2 or Level=3)) "
        f"and TimeCreated[@SystemTime >= '{xpath_since(since)}']]]"
    )
    fallback_msg = tr(lang, "no_msg_text")
    for ev in query_events("System", xpath, with_message=True):
        msg = " ".join((ev.get("message") or "").split())
        short_msg = (msg[:220] + "...") if len(msg) > 220 else msg
        rows.append({
            "time": ev["time"].strftime("%d.%m.%Y %H:%M:%S"),
            "type": tr(lang, "disk_io_err", provider=ev["provider"], id=ev["id"]),
            "details": short_msg or fallback_msg,
        })

    return rows


def collect_stopped_services(c: wmi.WMI, ignored: list[str]) -> list[dict]:
    return [
        {"name": s.Name, "display": s.DisplayName, "state": s.State}
        for s in c.Win32_Service(StartMode="Auto")
        if s.State != "Running" and s.Name not in ignored
    ]

def collect_new_services(since: datetime) -> list[dict]:
    xpath = (
        f"*[System[Provider[@Name='Service Control Manager'] and (EventID=7045) "
        f"and TimeCreated[@SystemTime >= '{xpath_since(since)}']]]"
    )
    rows = []
    for ev in query_events("System", xpath):
        d = ev["data"]
        rows.append({
            "time": ev["time"].strftime("%d.%m.%Y %H:%M:%S"),
            "name": d.get("ServiceName", ""),
            "path": d.get("ImagePath", ""),
            "start_type": d.get("StartType", ""),
            "account": d.get("AccountName", ""),
        })
    return rows


def collect_crashes_and_reboots(since: datetime, lang: str) -> list[dict]:
    xpath = (
        f"*[System["
        f"((Provider[@Name='Microsoft-Windows-Kernel-Power'] and EventID=41) or "
        f"(Provider[@Name='EventLog'] and EventID=6008) or "
        f"(Provider[@Name='Microsoft-Windows-WER-SystemErrorReporting' or @Name='BugCheck'] and EventID=1001) or "
        f"(Provider[@Name='Service Control Manager'] and (EventID=7031 or EventID=7034))) "
        f"and TimeCreated[@SystemTime >= '{xpath_since(since)}']]]"
    )
    rows = []
    fallback_msg = tr(lang, "no_msg_text")
    for ev in query_events("System", xpath, with_message=True):
        eid = ev["id"]
        if eid == 41:
            ev_type = tr(lang, "crash_41")
        elif eid == 6008:
            ev_type = tr(lang, "crash_6008")
        elif eid == 1001:
            ev_type = tr(lang, "crash_1001")
        else:
            ev_type = tr(lang, "crash_svc", id=eid)

        msg = " ".join((ev.get("message") or "").split())
        short_msg = (msg[:220] + "...") if len(msg) > 220 else msg
        rows.append({
            "time": ev["time"].strftime("%d.%m.%Y %H:%M:%S"),
            "type": ev_type,
            "details": short_msg or fallback_msg,
        })
    return rows

def collect_task_scheduler_errors(since: datetime, lang: str, ignored_tasks: list[str]) -> list[dict]:
    xpath = (
        f"*[System[(EventID=101 or EventID=103 or EventID=201) "
        f"and TimeCreated[@SystemTime >= '{xpath_since(since)}']]]"
    )
    rows = []
    fallback_msg = tr(lang, "no_msg_text")
    ignored_lower = [t.lower() for t in ignored_tasks if t]

    for ev in query_events("Microsoft-Windows-TaskScheduler/Operational", xpath, with_message=True):
        d = ev["data"]
        eid = ev["id"]
        task_name = d.get("TaskName", "")

        if task_name and any(ign in task_name.lower() for ign in ignored_lower):
            continue

        if eid == 201:
            rc = d.get("ResultCode", "0")
            if str(rc) in ("0", "0x0"):
                continue
            status = tr(lang, "task_err_201", code=rc)
        elif eid == 101:
            status = tr(lang, "task_err_101")
        else:
            status = tr(lang, "task_err_103")

        msg = " ".join((ev.get("message") or "").split())
        short_msg = (msg[:200] + "...") if len(msg) > 200 else msg
        rows.append({
            "time": ev["time"].strftime("%d.%m.%Y %H:%M:%S"),
            "task": task_name or "(Unknown)",
            "status": status,
            "details": short_msg or fallback_msg,
        })
    return rows


def collect_success_logons(since: datetime, lang: str) -> list[dict]:
    xpath = (
        f"*[System[((EventID=4624) or (EventID=4801)) "
        f"and TimeCreated[@SystemTime >= '{xpath_since(since)}']]]"
    )
    type_names = {
        "2": tr(lang, "logon_2"),
        "7": tr(lang, "logon_7"),
        "10": tr(lang, "logon_10"),
        "11": tr(lang, "logon_11"),
    }
    local_str = tr(lang, "local")
    rows = []
    for ev in query_events("Security", xpath):
        d = ev["data"]
        if ev["id"] == 4624:
            lt = d.get("LogonType", "")
            if lt not in ("2", "7", "10", "11"):
                continue
            domain = d.get("TargetDomainName", "")
            if domain in ("Window Manager", "Font Driver Host"):
                continue
            ip = d.get("IpAddress", "")
            rows.append({
                "time": ev["time"],
                "event": type_names.get(lt, lt),
                "account": f"{domain}\\{d.get('TargetUserName', '')}",
                "ip": ip if ip and ip != "-" else local_str,
            })
        elif ev["id"] == 4801:
            rows.append({
                "time": ev["time"],
                "event": tr(lang, "unlock_4801"),
                "account": f"{d.get('TargetDomainName', '')}\\{d.get('TargetUserName', '')}",
                "ip": local_str,
            })
    return rows


def collect_failed_logons(since: datetime, lang: str) -> list[dict]:
    xpath = f"*[System[(EventID=4625) and TimeCreated[@SystemTime >= '{xpath_since(since)}']]]"
    type_map = {
        "2": tr(lang, "fail_2"),
        "3": tr(lang, "fail_3"),
        "7": tr(lang, "fail_7"),
        "10": tr(lang, "fail_10"),
        "11": tr(lang, "fail_11"),
    }
    local_str = tr(lang, "local")
    not_spec = tr(lang, "not_specified")
    groups: dict[tuple, int] = defaultdict(int)
    for ev in query_events("Security", xpath):
        d = ev["data"]
        td, tu = d.get("TargetDomainName", ""), d.get("TargetUserName", "")
        user = f"{td}\\{tu}" if td and tu else (tu or not_spec)
        ip = d.get("IpAddress", "") or local_str
        if ip == "-":
            ip = local_str
        lt = d.get("LogonType", "")
        groups[(user, ip, type_map.get(lt, tr(lang, "fail_type", lt=lt)))] += 1
    return [
        {"user": u, "ip": i, "type": t, "count": n}
        for (u, i, t), n in sorted(groups.items(), key=lambda kv: -kv[1])
    ]


def collect_account_changes(since: datetime, lang: str) -> list[dict]:
    xpath = (
        f"*[System[(EventID=4720 or EventID=4726 or EventID=4732 or EventID=4740 or EventID=1102) "
        f"and TimeCreated[@SystemTime >= '{xpath_since(since)}']]]"
    )
    rows = []
    for ev in query_events("Security", xpath, with_message=True):
        eid = ev["id"]
        d = ev["data"]
        alert = eid in (4726, 4740, 1102)

        if eid == 4720:
            action = tr(lang, "acc_created")
        elif eid == 4726:
            action = tr(lang, "acc_deleted")
        elif eid == 4732:
            action = tr(lang, "acc_group")
        elif eid == 4740:
            action = tr(lang, "acc_locked")
        else:
            action = tr(lang, "sec_cleared")

        first_line = (ev.get("message") or "").splitlines()[0] if ev.get("message") else ""
        caller = d.get("TargetDomainName", "")
        user = d.get("TargetUserName", "") or d.get("SubjectUserName", "")
        fallback = tr(lang, "acc_fallback", user=f"{user} ({caller})" if caller and eid == 4740 else user)
        details = first_line or fallback

        rows.append({
            "time": ev["time"],
            "code": eid,
            "event": action,
            "details": details,
            "alert": alert,
        })
    return rows


def collect_software_changes(since: datetime, valid_dates: set[str], lang: str) -> list[dict]:
    rows = []
    since_utc = xpath_since(since)

    xpath_msi = (
        f"*[System[Provider[@Name='MsiInstaller'] and (EventID=11707) "
        f"and TimeCreated[@SystemTime >= '{since_utc}']]]"
    )
    for ev in query_events("Application", xpath_msi, with_message=True):
        msg = " ".join((ev.get("message") or "").split())
        rows.append({
            "time": ev["time"].strftime("%d.%m.%Y %H:%M"),
            "type": tr(lang, "sw_msi"),
            "description": msg or tr(lang, "sw_msi_fallback"),
        })

    uninstall_paths = [
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"),
    ]
    for root_key, subkey_path in uninstall_paths:
        try:
            with winreg.OpenKey(root_key, subkey_path) as key:
                for i in range(winreg.QueryInfoKey(key)[0]):
                    try:
                        sub_name = winreg.EnumKey(key, i)
                        with winreg.OpenKey(key, sub_name) as sub_key:
                            install_date, _ = winreg.QueryValueEx(sub_key, "InstallDate")
                            if str(install_date) in valid_dates:
                                name, _ = winreg.QueryValueEx(sub_key, "DisplayName")
                                try:
                                    ver, _ = winreg.QueryValueEx(sub_key, "DisplayVersion")
                                    desc = f"{name} ({ver})"
                                except FileNotFoundError:
                                    desc = str(name)
                                rows.append({
                                    "time": f"{install_date[6:]}.{install_date[4:6]}.{install_date[:4]}",
                                    "type": tr(lang, "sw_reg"),
                                    "description": desc,
                                })
                    except (FileNotFoundError, OSError):
                        continue
        except OSError:
            continue

    xpath_wu = f"*[System[(EventID=19) and TimeCreated[@SystemTime >= '{since_utc}']]]"
    for ev in query_events("Microsoft-Windows-WindowsUpdateClient/Operational", xpath_wu, with_message=True):
        msg = " ".join((ev.get("message") or "").split())
        rows.append({
            "time": ev["time"].strftime("%d.%m.%Y %H:%M"),
            "type": tr(lang, "sw_wu"),
            "description": msg or tr(lang, "sw_wu_fallback"),
        })

    return rows


def collect_critical_events(since: datetime, lang: str) -> list[dict]:
    xpath = (
        f"*[System[(Level=1 or Level=2) "
        f"and TimeCreated[@SystemTime >= '{xpath_since(since)}']]]"
    )
    groups: dict[tuple, dict] = {}
    fallback_msg = tr(lang, "no_msg_text")
    for log_name in ("System", "Application"):
        for ev in query_events(log_name, xpath, with_message=True):
            key = (log_name, ev["provider"], ev["id"])
            if key not in groups:
                msg = " ".join((ev.get("message") or "").split())
                short_msg = (msg[:220] + "...") if len(msg) > 220 else msg
                groups[key] = {
                    "count": 0,
                    "last": ev["time"].strftime("%d.%m %H:%M:%S"),
                    "last_dt": ev["time"],
                    "log": log_name,
                    "provider": ev["provider"],
                    "id": ev["id"],
                    "message": short_msg or fallback_msg,
                }
            groups[key]["count"] += 1
            if ev["time"] > groups[key]["last_dt"]:
                groups[key]["last_dt"] = ev["time"]
                groups[key]["last"] = ev["time"].strftime("%d.%m %H:%M:%S")

    return sorted(groups.values(), key=lambda x: (-x["count"], -x["last_dt"].timestamp()))

def render_table(headers: list[str], items: list, row_builder, max_rows: int, lang: str) -> str:
    th_html = "".join(f"<th>{esc(h)}</th>" for h in headers)
    shown = items[:max_rows] if max_rows > 0 else items
    rows_html = "".join(row_builder(item) for item in shown)
    
    if max_rows > 0 and len(items) > max_rows:
        hidden = len(items) - max_rows
        msg = esc(tr(lang, "more_rows", count=hidden))
        rows_html += f"<tr><td colspan='{len(headers)}' style='color:#666; font-style:italic; text-align:center;'>{msg}</td></tr>"
        
    return f"<table><tr>{th_html}</tr>{rows_html}</table>"

def build_report(cfg: dict, lang: str) -> str:
    t_start = time.perf_counter()
    since = datetime.now() - timedelta(days=cfg["lookback_days"])
    c = wmi.WMI()
    max_rows = int(cfg.get("max_rows", 50))
    tel = collect_telemetry(c)
    uptime_str = tr(
        lang,
        "uptime_fmt",
        d=tel["uptime"].days,
        h=tel["uptime"].seconds // 3600,
        m=(tel["uptime"].seconds // 60) % 60,
        boot=tel["last_boot"].strftime("%d.%m %H:%M"),
    )
    up_style = "red" if tel["uptime"].total_seconds() < 86400 else ""
    ram_style = "red" if tel["used_pct"] > 90 else ""
    ram_str = tr(lang, "ram_fmt", used=tel["used_pct"], free=tel["free_gb"], total=tel["total_gb"])
    reboot_str = tr(lang, "reboot_req") if tel["reboot_pending"] else tr(lang, "reboot_ok")
    reboot_style = "red" if tel["reboot_pending"] else "green"

    telemetry_html = f"""
    <table class="summary-table">
      <tr>
        <th>{esc(tr(lang, 'os'))}</th><td>{esc(tel['caption'])}</td>
        <th>{esc(tr(lang, 'uptime'))}</th><td class="{up_style}">{esc(uptime_str)}</td>
      </tr>
      <tr>
        <th>{esc(tr(lang, 'ram'))}</th><td class="{ram_style}">{esc(ram_str)}</td>
        <th>{esc(tr(lang, 'reboot'))}</th><td class="{reboot_style}">{esc(reboot_str)}</td>
      </tr>
    </table>"""

    # Top 5 RAM processes (показываем при загрузке RAM > 85%)
    top_ram_html = ""
    if tel["used_pct"] > 85:
        top_procs = collect_top_ram_processes(c, limit=5)
        if top_procs:
            rows = "".join(
                f"<tr><td>{esc(p['name'])}</td><td>{p['pid']}</td><td>{p['ram_mb']}</td></tr>"
                for p in top_procs
            )
            top_ram_html = f"""
            <h4>{esc(tr(lang, 'sub_top_ram'))}</h4>
            <table>
              <tr>
                <th>{esc(tr(lang, 'th_proc_name'))}</th>
                <th>{esc(tr(lang, 'th_pid'))}</th>
                <th>{esc(tr(lang, 'th_ram_mb'))}</th>
              </tr>
              {rows}
            </table>"""

    disks = collect_disks(c, cfg["threshold_percent"], cfg["threshold_gb"])
    disk_rows = "".join(
        f"<tr{' class=\"alert\"' if d['alert'] else ''}>"
        f"<td>{esc(d['device'])}</td><td>{esc(d['label'])}</td>"
        f"<td>{d['total_gb']}</td><td>{d['free_gb']}</td><td>{d['free_pct']} %</td></tr>"
        for d in disks
    )
    disks_html = f"""
    <h3>{esc(tr(lang, 'sec_disks'))}</h3>
    <table>
      <tr>
        <th>{esc(tr(lang, 'th_disk'))}</th>
        <th>{esc(tr(lang, 'th_label'))}</th>
        <th>{esc(tr(lang, 'th_total_gb'))}</th>
        <th>{esc(tr(lang, 'th_free_gb'))}</th>
        <th>{esc(tr(lang, 'th_free_pct'))}</th>
      </tr>
      {disk_rows}
    </table>"""


    disk_health = collect_disk_health(since, lang)
    if disk_health:
        table = render_table(
            [tr(lang, 'th_time'), tr(lang, 'th_type'), tr(lang, 'th_details')],
            disk_health,
            lambda r: f"<tr class='alert'><td>{esc(r['time'])}</td><td>{esc(r['type'])}</td><td>{esc(r['details'])}</td></tr>",
            max_rows, lang
        )
        disk_health_html = f"<h3>{esc(tr(lang, 'sec_disk_health'))}</h3>{table}"
    else:
        disk_health_html = f"<h3>{esc(tr(lang, 'sec_disk_health'))}</h3><p class='ok'>{esc(tr(lang, 'no_disk_health'))}</p>"


    stopped = collect_stopped_services(c, cfg["ignored_services"])
    if stopped:
        table = render_table(
            [tr(lang, 'th_service'), tr(lang, 'th_name'), tr(lang, 'th_status')],
            stopped,
            lambda s: f"<tr><td>{esc(s['name'])}</td><td>{esc(s['display'])}</td><td>{esc(s['state'])}</td></tr>",
            max_rows, lang
        )
        services_html = f"<h3>{esc(tr(lang, 'sec_services'))}</h3>{table}"
    else:
        services_html = f"<h3>{esc(tr(lang, 'sec_services'))}</h3><p class='ok'>{esc(tr(lang, 'services_ok'))}</p>"


    new_services = collect_new_services(since)
    if new_services:
        table = render_table(
            [tr(lang, 'th_time'), tr(lang, 'th_service'), tr(lang, 'th_start_type'), tr(lang, 'th_account'), tr(lang, 'th_bin_path')],
            new_services,
            lambda r: f"<tr class='alert'><td>{esc(r['time'])}</td><td>{esc(r['name'])}</td><td>{esc(r['start_type'])}</td><td>{esc(r['account'])}</td><td>{esc(r['path'])}</td></tr>",
            max_rows, lang
        )
        new_services_html = f"<h4>{esc(tr(lang, 'sub_new_services'))}</h4>{table}"
    else:
        new_services_html = f"<h4>{esc(tr(lang, 'sub_new_services'))}</h4><p class='ok'>{esc(tr(lang, 'no_new_services'))}</p>"

    crashes = collect_crashes_and_reboots(since, lang)
    if crashes:
        table = render_table(
            [tr(lang, 'th_time'), tr(lang, 'th_type'), tr(lang, 'th_details')],
            crashes,
            lambda r: f"<tr class='alert'><td>{esc(r['time'])}</td><td>{esc(r['type'])}</td><td>{esc(r['details'])}</td></tr>",
            max_rows, lang
        )
        crashes_html = f"<h3>{esc(tr(lang, 'sec_crashes'))}</h3>{table}"
    else:
        crashes_html = f"<h3>{esc(tr(lang, 'sec_crashes'))}</h3><p class='ok'>{esc(tr(lang, 'no_crashes'))}</p>"

    task_errors = collect_task_scheduler_errors(since, lang, cfg.get("ignored_tasks", []))
    if task_errors:
        table = render_table(
            [tr(lang, 'th_time'), tr(lang, 'th_task_name'), tr(lang, 'th_status'), tr(lang, 'th_details')],
            task_errors,
            lambda r: f"<tr class='alert'><td>{esc(r['time'])}</td><td>{esc(r['task'])}</td><td>{esc(r['status'])}</td><td>{esc(r['details'])}</td></tr>",
            max_rows, lang
        )
        tasks_html = f"<h3>{esc(tr(lang, 'sec_tasks'))}</h3>{table}"
    else:
        tasks_html = f"<h3>{esc(tr(lang, 'sec_tasks'))}</h3><p class='ok'>{esc(tr(lang, 'no_tasks'))}</p>"

    success = collect_success_logons(since, lang)
    if success:
        table = render_table(
            [tr(lang, 'th_time'), tr(lang, 'th_event'), tr(lang, 'th_account'), tr(lang, 'th_ip')],
            success,
            lambda r: f"<tr><td>{r['time'].strftime('%d.%m.%Y %H:%M:%S')}</td><td>{esc(r['event'])}</td><td>{esc(r['account'])}</td><td>{esc(r['ip'])}</td></tr>",
            max_rows, lang
        )
        success_html = f"<h4>{esc(tr(lang, 'sub_success'))}</h4>{table}"
    else:
        success_html = f"<h4>{esc(tr(lang, 'sub_success'))}</h4><p>{esc(tr(lang, 'no_success'))}</p>"


    failed = collect_failed_logons(since, lang)
    if failed:
        table = render_table(
            [tr(lang, 'th_attempts'), tr(lang, 'th_account'), tr(lang, 'th_ip'), tr(lang, 'th_logon_type')],
            failed,
            lambda r: f"<tr><td>{r['count']}</td><td>{esc(r['user'])}</td><td>{esc(r['ip'])}</td><td>{esc(r['type'])}</td></tr>",
            max_rows, lang
        )
        failed_html = f"<h4>{esc(tr(lang, 'sub_failed'))}</h4>{table}"
    else:
        failed_html = f"<h4>{esc(tr(lang, 'sub_failed'))}</h4><p class='ok'>{esc(tr(lang, 'no_failed'))}</p>"

    accounts = collect_account_changes(since, lang)
    if accounts:
        table = render_table(
            [tr(lang, 'th_time'), tr(lang, 'th_code'), tr(lang, 'th_event'), tr(lang, 'th_details')],
            accounts,
            lambda r: f"<tr{' class=\"alert\"' if r.get('alert') else ''}><td>{r['time'].strftime('%d.%m.%Y %H:%M:%S')}</td><td>{r['code']}</td><td>{esc(r['event'])}</td><td>{esc(r['details'])}</td></tr>",
            max_rows, lang
        )
        accounts_html = f"<h4>{esc(tr(lang, 'sub_accounts'))}</h4>{table}"
    else:
        accounts_html = f"<h4>{esc(tr(lang, 'sub_accounts'))}</h4><p class='ok'>{esc(tr(lang, 'no_accounts'))}</p>"

    security_html = f"<h3>{esc(tr(lang, 'sec_security'))}</h3>{success_html}{failed_html}{accounts_html}"


    valid_dates = {(datetime.now()).strftime("%Y%m%d"), since.strftime("%Y%m%d")}

    changes = collect_software_changes(since, valid_dates, lang)
    if changes:
        table = render_table(
            [tr(lang, 'th_time'), tr(lang, 'th_type'), tr(lang, 'th_desc')],
            changes,
            lambda r: f"<tr><td>{esc(r['time'])}</td><td>{esc(r['type'])}</td><td>{esc(r['description'])}</td></tr>",
            max_rows, lang
        )
        software_html = f"<h3>{esc(tr(lang, 'sec_software'))}</h3>{table}"
    else:
        software_html = f"<h3>{esc(tr(lang, 'sec_software'))}</h3><p>{esc(tr(lang, 'no_software'))}</p>"

    critical = collect_critical_events(since, lang)
    if critical:
        table = render_table(
            [tr(lang, 'th_count'), tr(lang, 'th_last_fail'), tr(lang, 'th_log'), tr(lang, 'th_source'), 'ID', tr(lang, 'th_msg')],
            critical,
            lambda r: f"<tr><td>{r['count']}</td><td>{esc(r['last'])}</td><td>{esc(r['log'])}</td><td>{esc(r['provider'])}</td><td>{r['id']}</td><td>{esc(r['message'])}</td></tr>",
            max_rows, lang
        )
        events_html = f"<h3>{esc(tr(lang, 'sec_critical'))}</h3>{table}"
    else:
        events_html = f"<h3>{esc(tr(lang, 'sec_critical'))}</h3><p class='ok'>{esc(tr(lang, 'no_critical'))}</p>"

    elapsed = round(time.perf_counter() - t_start, 2)
    footer_text = tr(lang, "footer_fmt", sec=elapsed, days=cfg["lookback_days"])
    footer_html = f'<div class="footer">{esc(footer_text)}</div>'

    host = os.environ.get("COMPUTERNAME") or socket.gethostname()
    return (
        f"<html lang=\"{esc(tr(lang, 'html_lang'))}\"><head><meta charset='utf-8'>{STYLE}</head><body>"
        f"<h2>{esc(tr(lang, 'report_title'))}: {esc(host)}</h2>"
        f"{telemetry_html}{top_ram_html}{disks_html}{disk_health_html}{services_html}{new_services_html}{crashes_html}{tasks_html}{security_html}{software_html}{events_html}"
        f"{footer_html}</body></html>"
    )

    host = os.environ.get("COMPUTERNAME") or socket.gethostname()
    return (
        f"<html lang=\"{esc(tr(lang, 'html_lang'))}\"><head><meta charset='utf-8'>{STYLE}</head><body>"
        f"<h2>{esc(tr(lang, 'report_title'))}: {esc(host)}</h2>"
        f"{telemetry_html}{top_ram_html}{disks_html}{disk_health_html}{services_html}{new_services_html}{crashes_html}{tasks_html}{security_html}{software_html}{events_html}"
        f"</body></html>"
    )

def send_email(cfg: dict, subject: str, html_body: str) -> None:
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = cfg["mail_from"]

    mail_to = cfg["mail_to"]
    if isinstance(mail_to, list):
        msg["To"] = ", ".join(mail_to)
        recipients = mail_to
    else:
        msg["To"] = str(mail_to)
        recipients = [addr.strip() for addr in str(mail_to).split(",") if addr.strip()]

    msg.attach(MIMEText(html_body, "html", "utf-8"))

    server_host = cfg["smtp_server"]
    server_port = int(cfg["smtp_port"])

    if cfg.get("smtp_use_ssl"):
        server = smtplib.SMTP_SSL(server_host, server_port, timeout=30)
    else:
        server = smtplib.SMTP(server_host, server_port, timeout=30)

    try:
        server.ehlo()
        if cfg.get("smtp_use_starttls") and not cfg.get("smtp_use_ssl"):
            server.starttls()
            server.ehlo()
        if cfg.get("smtp_user"):
            server.login(cfg["smtp_user"], cfg.get("smtp_password", ""))
        server.sendmail(cfg["mail_from"], recipients, msg.as_string())
    finally:
        server.quit()


def main() -> int:
    parser = argparse.ArgumentParser(description="Windows daily logwatch report")
    parser.add_argument("--config", help="Path to JSON config file")
    parser.add_argument("--test", action="store_true", help="Save HTML report locally without sending")
    parser.add_argument("--no-log", action="store_true", help="Disable writing to logwatch.log")
    parser.add_argument("--lang", choices=["auto", "ru", "en"], help="Report language (auto, ru, en)")
    args = parser.parse_args()


    cfg, created = load_config(args.config)
    if args.no_log:
        cfg["enable_log"] = False
    if args.lang:
        cfg["language"] = args.lang
    setup_logging(cfg)

    if created and not args.test:
        logging.info("Default config created at %s. Please configure SMTP settings.", app_dir() / "logwatch.json")
        return 0


    lang = detect_language(cfg.get("language", "auto"))
    logging.info("Report language selected: %s", lang)

    host = os.environ.get("COMPUTERNAME") or socket.gethostname()
    subject = f"{cfg['subject_prefix']} {host}"

    try:
        body = build_report(cfg, lang)
    except Exception as e:
        logging.exception("Failed to build report: %s", e)
        return 1

    if args.test:
        out = app_dir() / "logwatch-report.html"
        out.write_text(body, encoding="utf-8")
        return 0

    try:
        send_email(cfg, subject, body)
    except Exception as e:
        logging.exception("Failed to send email: %s", e)
        return 1

    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except SystemExit:
        raise
    except BaseException as e:
        with open(startup_log, "a", encoding="utf-8") as f:
            f.write(f"CRITICAL RUNTIME ERROR: {e}\n{traceback.format_exc()}\n")
        sys.exit(1)
