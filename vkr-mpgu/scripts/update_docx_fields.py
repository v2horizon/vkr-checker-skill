#!/usr/bin/env python3
"""Обновить оглавление и поля DOCX через установленный Microsoft Word.

Сборка ``build_vkr.py`` вставляет оглавление как поле: номера страниц в нём
появляются только после обновления в Word (Ctrl+A, F9). Этот скрипт делает то же
самое автоматически, если на компьютере есть Windows и Microsoft Word:

1. проверяет, что документ можно перезаписать: если он открыт в Word (или защищён
   от записи), скрипт сразу завершается с кодом 2 и просьбой закрыть файл — Word
   не запускается, ничего не меняется;
2. открывает копию документа в невидимом Word;
3. обновляет все поля и оглавления, пересчитывает разметку;
4. сохраняет результат в новый временный файл (сохранение поверх открытого
   файла на некоторых машинах зависает);
5. проверяет, что результат — корректный DOCX, и атомарно заменяет исходный файл
   через промежуточный ``.<имя>.fields-<uuid>.tmp`` рядом с ним; промежуточный
   файл удаляется в любом случае, в том числе при ошибке замены.

Если документ лежит в проекте ВКР, прежняя версия копируется в
``backups/checkpoints/<время>-fields/`` (при ошибке замены копия убирается), а
после успешной замены атомарно пишется ``exports/docx-fields.json``:
``{"docx", "docx_sha256", "pages", "word_opened", "updated_at",
"fingerprint_schema", "text_fingerprint"}`` — число страниц по Word для сверки
объёма в аннотации (``docx_sha256`` меняется после очистки метаданных,
``text_fingerprint`` — нет). ``word_opened`` равно ``true`` только если Word
действительно открыл документ.

«Word недоступен» и «Word не открыл файл» — разные исходы. Если Word запустился,
но ``Documents.Open`` вернул ошибку («Ошибка при попытке открытия файла»),
разметка DOCX нарушена: статус ``open_failed``, код 2 и подсказка проверить файл
``docx_integrity.py`` и пересобрать его. Отсутствие Word — по-прежнему код 1 и
ручная инструкция.

Процесс Word. Word запускается с тайм-аутом. Сразу после ``Documents.Open``
скрипт PowerShell узнаёт PID процесса, которому принадлежит окно открытого
документа: ``GetWindowThreadProcessId`` (user32, через ``Add-Type``) по
``$doc.ActiveWindow.Hwnd``. При тайм-ауте или сбое принудительно
останавливается только этот процесс. Если окно принадлежит процессу WINWORD,
который работал до запуска скрипта (COM подключился к уже открытому Word
пользователя), процесс не останавливается и не закрывается (``Quit`` не
вызывается). Запасной путь — только если PID окна получить не удалось
(например, Word завис до открытия документа): новые процессы WINWORD, которых не
было до запуска, с ключом ``/Automation`` или ``-Embedding`` в командной строке
(Word, запущенный пользователем, таких ключей не имеет). Перед ``taskkill``
проверяется, что PID всё ещё принадлежит WINWORD.EXE. Чужие процессы WINWORD
не трогаются.

Если Word недоступен (другая ОС, нет Word, запрещён запуск PowerShell), скрипт
ничего не меняет, печатает ручную инструкцию и возвращает код 1.

После обновления полей документ изменился: снова выполни очистку метаданных,
``vkr_audit.py validate`` и ``vkr_audit.py snapshot``.

Коды выхода: 0 — поля обновлены; 1 — автоматически обновить нельзя, нужен ручной
шаг (нет Windows, нет Word, Word не уложился в тайм-аут); 2 — ошибка
ввода-вывода, неверные аргументы (в том числе документ открыт в Word — закрой его
и повтори) или Word не открыл файл (``open_failed``: повреждена разметка DOCX).
"""

from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import tempfile
import time
import uuid
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

sys.dont_write_bytecode = True
SCRIPTS_DIR = Path(__file__).resolve().parent
FIELDS_RECORD_PATH = "exports/docx-fields.json"

MANUAL_STEPS = [
    "Открой DOCX в Microsoft Word (или LibreOffice Writer).",
    "Выдели весь текст (Ctrl+A) и нажми F9; для оглавления выбери «Обновить целиком».",
    "Сохрани файл под тем же именем и закрой его.",
    "Затем выполни очистку метаданных, vkr_audit.py validate и vkr_audit.py snapshot.",
]

# PID-файл: «window:<pid>» — процесс окна документа (останавливается при тайм-ауте);
# «shared:<pid>» — окно в процессе Word, запущенном до скрипта (не останавливается);
# «diff:<pid>,…» — запасной путь до получения PID окна.
POWERSHELL_SCRIPT = r"""
param([string]$Source, [string]$Saved, [string]$PidFile)
$ErrorActionPreference = 'Stop'
# Сообщения Word приходят на языке Office: без этого они доедут в кодировке консоли и станут «крякозябрами».
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}
$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
$word = New-Object -ComObject Word.Application
$fresh = @()
try {
    $fresh = @(Get-CimInstance Win32_Process -Filter "Name='WINWORD.EXE'" -ErrorAction Stop |
        Where-Object { ($before -notcontains [int]$_.ProcessId) -and ([string]$_.CommandLine -match '(?i)(/automation|-embedding)') } |
        ForEach-Object { [int]$_.ProcessId })
} catch {
    $fresh = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } | ForEach-Object { $_.Id })
}
Set-Content -LiteralPath $PidFile -Value ("diff:" + ($fresh -join ",")) -Encoding ASCII
$doc = $null
$shared = $false
$openFailed = $false
try {
    $word.Visible = $false
    $word.DisplayAlerts = 0
    try {
        $doc = $word.Documents.Open($Source, $false, $false, $false)
        Write-Output "opened=1"
    } catch {
        $openFailed = $true
        Write-Output ("open_error=" + ($_.Exception.Message -replace "`r`n", ' '))
    }
    if (-not $openFailed) {
        $windowPid = 0
        try {
            if (-not ('VkrFieldsUser32' -as [type])) {
                Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class VkrFieldsUser32 { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId); }'
            }
            [uint32]$found = 0
            [void][VkrFieldsUser32]::GetWindowThreadProcessId([IntPtr][int64]$doc.ActiveWindow.Hwnd, [ref]$found)
            $windowPid = [int]$found
        } catch {
            $windowPid = 0
        }
        if ($windowPid -gt 0) {
            if ($before -contains $windowPid) {
                $shared = $true
                Set-Content -LiteralPath $PidFile -Value ("shared:" + $windowPid) -Encoding ASCII
            } else {
                Set-Content -LiteralPath $PidFile -Value ("window:" + $windowPid) -Encoding ASCII
            }
        }
        Write-Output "word_pid=$windowPid"
        $doc.Fields.Update() | Out-Null
        foreach ($toc in $doc.TablesOfContents) { $toc.Update() }
        $doc.Repaginate()
        $pages = $doc.ComputeStatistics(2)
        $doc.SaveAs2($Saved, 16)
        Write-Output "pages=$pages"
    }
}
finally {
    if ($doc -ne $null) { $doc.Close($false) | Out-Null }
    if (-not $shared) { $word.Quit() | Out-Null }
}
if ($openFailed) { exit 3 }
"""

# Код выхода скрипта PowerShell, когда Word запустился, но отказался открыть документ.
OPEN_FAILED_EXIT = 3
OPEN_FAILED_HINT = (
    "Word запустился, но отказался открыть документ («Ошибка при попытке открытия файла»): "
    "разметка DOCX нарушена. Проверь структуру: docx_integrity.py <DOCX> и пересобери файл "
    "(build_vkr.py); правки из прежнего DOCX переносит import_docx.py --update"
)

WINDOWS_SHARING_ERRORS = (32, 33)  # ERROR_SHARING_VIOLATION, ERROR_LOCK_VIOLATION
WINDOWS_ACCESS_DENIED = 5


class FieldsError(Exception):
    """Ошибка ввода-вывода или аргументов (код 2)."""


def configure_output() -> None:
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            try:
                stream.reconfigure(encoding="utf-8", errors="replace")
            except (ValueError, OSError):
                pass


def find_powershell() -> Optional[str]:
    if os.name != "nt":
        return None
    for name in ("powershell.exe", "powershell", "pwsh.exe", "pwsh"):
        found = shutil.which(name)
        if found:
            return found
    return None


def find_project_root(path: Path) -> Optional[Path]:
    for parent in list(path.parents)[:4]:
        if (parent / "vkr-project.json").is_file():
            return parent
    return None


def is_valid_docx(path: Path) -> bool:
    try:
        with zipfile.ZipFile(path) as archive:
            names = set(archive.namelist())
            return "word/document.xml" in names and archive.testzip() is None
    except (OSError, zipfile.BadZipFile):
        return False


# ---------------------------------------------------------------------------
# Можно ли перезаписать файл (общая проверка для build_vkr.py и этого скрипта)
# ---------------------------------------------------------------------------

def _windows_open_error(path: Path) -> Optional[int]:
    """Код ошибки Windows при открытии существующего файла на запись и удаление; None — открывается.

    Файл открывается без изменения содержимого (OPEN_EXISTING) и сразу закрывается.
    Документ, открытый в Word, даёт ERROR_SHARING_VIOLATION: атомарная замена
    (os.replace) такого файла тоже не пройдёт.
    """
    import ctypes
    from ctypes import wintypes

    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    create = kernel32.CreateFileW
    create.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD, wintypes.LPVOID,
                       wintypes.DWORD, wintypes.DWORD, wintypes.HANDLE]
    create.restype = wintypes.HANDLE
    close = kernel32.CloseHandle
    close.argtypes = [wintypes.HANDLE]
    close.restype = wintypes.BOOL
    target = os.path.abspath(str(path))
    if len(target) >= 240 and not target.startswith("\\\\?\\"):
        target = "\\\\?\\UNC\\" + target[2:] if target.startswith("\\\\") else "\\\\?\\" + target
    generic_write, delete_access = 0x40000000, 0x00010000
    share_all = 0x1 | 0x2 | 0x4
    open_existing, normal = 3, 0x80
    handle = create(target, generic_write | delete_access, share_all, None, open_existing, normal, None)
    invalid = ctypes.c_void_p(-1).value
    if handle is None or handle == invalid:
        return ctypes.get_last_error() or -1
    close(handle)
    return None


def write_blocker(path: Path) -> Optional[str]:
    """Почему существующий файл нельзя перезаписать атомарной заменой; None — можно (или файла нет).

    Файл не меняется. В Windows проверяется, что файл открывается на запись и
    удаление при любом режиме совместного доступа: документ, открытый в Word,
    защищённый от записи или занятый другой программой, не открывается.
    """
    path = Path(path)
    if not path.exists():
        return None
    if not path.is_file():
        return "это не файл"
    if os.name != "nt":
        return None
    try:
        code = _windows_open_error(path)
    except (OSError, AttributeError, ValueError, ImportError):  # ctypes недоступен — проверит сама замена
        return None
    if code is None:
        return None
    if code in WINDOWS_SHARING_ERRORS:
        return "файл открыт в Word или другой программе"
    if code == WINDOWS_ACCESS_DENIED:
        return "нет доступа на запись: файл открыт в Word, помечен «только чтение» или защищён"
    return f"файл нельзя открыть на запись (код Windows {code})"


def directory_blocker(directory: Path) -> Optional[str]:
    """Почему в каталоге (или ближайшем существующем родителе) нельзя создать файл; None — можно."""
    probe_dir = Path(directory)
    while not probe_dir.exists() and probe_dir.parent != probe_dir:
        probe_dir = probe_dir.parent
    try:
        handle, name = tempfile.mkstemp(prefix=".vkr-write-probe-", suffix=".tmp", dir=str(probe_dir))
        os.close(handle)
        os.unlink(name)
    except OSError as error:
        return f"нельзя создать файл в {probe_dir}: {error}"
    return None


# ---------------------------------------------------------------------------
# Процесс Word
# ---------------------------------------------------------------------------

def recorded_word_pids(pid_file: Path) -> Tuple[str, List[int]]:
    """(источник, PID) из PID-файла скрипта PowerShell: window | shared | diff | none."""
    try:
        raw = pid_file.read_text(encoding="ascii", errors="ignore")
    except OSError:
        return "none", []
    lines = [line.strip() for line in raw.splitlines() if line.strip()]
    if not lines:
        return "none", []
    source, _, values = lines[-1].partition(":")
    if source not in ("window", "shared", "diff"):
        return "none", []
    pids = [int(token) for token in values.replace(" ", "").split(",") if token.isdigit() and int(token) > 0]
    return source, pids


def is_word_process(pid: int) -> bool:
    """PID всё ещё принадлежит WINWORD.EXE (защита от повторно выданного PID)."""
    try:
        result = subprocess.run(
            ["tasklist", "/FI", f"PID eq {pid}", "/FI", "IMAGENAME eq WINWORD.EXE", "/NH", "/FO", "CSV"],
            stdout=subprocess.PIPE, stderr=subprocess.DEVNULL, check=False,
        )
    except OSError:
        return False
    return f'"{pid}"' in result.stdout.decode("ascii", errors="ignore")


def kill_recorded_word(pid_file: Path) -> List[int]:
    """Останавливает только Word этого запуска: PID окна документа, иначе новые процессы автоматизации."""
    source, pids = recorded_word_pids(pid_file)
    if source not in ("window", "diff"):
        return []  # shared — документ открылся в уже работавшем Word: не наш процесс
    stopped: List[int] = []
    for pid in pids:
        if not is_word_process(pid):
            continue
        result = subprocess.run(
            ["taskkill", "/PID", str(pid), "/F"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            check=False,
        )
        if result.returncode == 0:
            stopped.append(pid)
    return stopped


def write_fields_record(project: Path, docx: Path, pages: Optional[int],
                        word_opened: bool = False) -> Tuple[str, Optional[int]]:
    """Атомарно пишет exports/docx-fields.json — число страниц после обновления полей в Word.

    Возвращает (путь записи, записанное число страниц).

    ``docx_sha256`` — хеш файла сразу после обновления (меняется после очистки
    метаданных); ``text_fingerprint`` — отпечаток сдаваемого текста (как в
    build-manifest.json), он переживает очистку метаданных. ``word_opened`` — Word
    действительно открыл документ (иначе число страниц взялось бы неизвестно откуда).
    """
    if str(SCRIPTS_DIR) not in sys.path:
        sys.path.insert(0, str(SCRIPTS_DIR))
    import vkr_common as common  # noqa: WPS433 — только для записи результата в проекте

    # Число страниц берётся из самого сохранённого файла (docProps/app.xml), чтобы запись и DOCX
    # не разошлись; ComputeStatistics Word — запасной путь, если статистики в файле нет.
    file_pages, _reason = common.docx_app_pages(docx)
    if file_pages:
        pages = file_pages
    fingerprint, _error = common.docx_text_fingerprint(docx)
    try:
        docx_path = Path(docx).resolve().relative_to(project.resolve()).as_posix()
    except ValueError:
        docx_path = str(docx)
    payload = {
        "schema": "vkr-docx-fields",
        "docx": docx_path,
        "docx_sha256": common.sha256_file(docx),
        "pages": pages,
        "word_opened": bool(word_opened),
        "updated_at": common.utc_now_iso(),
        "fingerprint_schema": common.FINGERPRINT_SCHEMA,
        "text_fingerprint": fingerprint,
    }
    common.atomic_write_text(project / FIELDS_RECORD_PATH, common.dump_json(payload))
    return FIELDS_RECORD_PATH, pages


def _unique_backup_dir(project: Path) -> Path:
    base = project / "backups" / "checkpoints"
    stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    folder = base / f"{stamp}-fields"
    counter = 2
    while folder.exists():
        folder = base / f"{stamp}-fields-{counter}"
        counter += 1
    return folder


def update_fields(docx: Path, timeout: int, powershell: Optional[str]) -> Dict[str, Any]:
    if not docx.is_file():
        raise FieldsError(f"файл не найден: {docx}")
    if docx.suffix.casefold() != ".docx":
        raise FieldsError("ожидается файл .docx")
    if not is_valid_docx(docx):
        raise FieldsError("файл не является корректным DOCX")
    # До запуска Word: открытый в Word документ заменить нельзя.
    blocker = write_blocker(docx)
    if blocker:
        raise FieldsError(f"{docx}: {blocker} — закрой файл в Word и повтори; ничего не изменено")
    blocker = directory_blocker(docx.parent)
    if blocker:
        raise FieldsError(f"{blocker}; ничего не изменено")
    if powershell is None:
        return {
            "status": "manual_required",
            "reason": "Microsoft Word через PowerShell недоступен (нужны Windows и установленный Word)",
            "manual_steps": MANUAL_STEPS,
            "document": str(docx),
        }

    work = Path(tempfile.mkdtemp(prefix="vkr-fields-"))
    staging: Optional[Path] = None
    backup_dir: Optional[Path] = None
    replaced = False
    try:
        source = work / "source.docx"
        saved = work / f"saved-{uuid.uuid4().hex}.docx"
        pid_file = work / "word-pid.txt"
        script = work / "update-fields.ps1"
        shutil.copyfile(docx, source)
        # PowerShell 5.1 читает UTF-8 без BOM как ANSI, поэтому пишем с BOM.
        script.write_text(POWERSHELL_SCRIPT, encoding="utf-8-sig")
        command = [
            powershell, "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
            "-File", str(script), "-Source", str(source), "-Saved", str(saved), "-PidFile", str(pid_file),
        ]
        started = time.monotonic()
        try:
            completed = subprocess.run(
                command,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                timeout=timeout,
                check=False,
            )
        except subprocess.TimeoutExpired:
            source_kind, pids = recorded_word_pids(pid_file)
            stopped = kill_recorded_word(pid_file)
            return {
                "status": "manual_required",
                "reason": f"Word не завершил обновление за {timeout} с; остановлены только процессы Word этого "
                          f"запуска ({source_kind}): {stopped}",
                "word_pid": {"source": source_kind, "pids": pids, "stopped": stopped},
                "manual_steps": MANUAL_STEPS,
                "document": str(docx),
            }
        output = completed.stdout.decode("utf-8", errors="replace") + completed.stderr.decode("utf-8", errors="replace")
        source_kind, pids = recorded_word_pids(pid_file)
        word_opened = "opened=1" in output
        if completed.returncode == OPEN_FAILED_EXIT or "open_error=" in output:
            # Word работает, но документ не открылся: это не «нет Word», а повреждённая разметка DOCX.
            stopped = kill_recorded_word(pid_file)
            detail = ""
            for line in output.splitlines():
                if line.strip().startswith("open_error="):
                    detail = " ".join(line.split("=", 1)[1].split())[:500]
                    break
            return {
                "status": "open_failed",
                "reason": OPEN_FAILED_HINT + (f". Ответ Word: {detail}" if detail else ""),
                "word_error": detail,
                "word_opened": False,
                "word_pid": {"source": source_kind, "pids": pids, "stopped": stopped},
                "next_steps": [
                    "docx_integrity.py <DOCX> --json",
                    "build_vkr.py <PROJECT_DIR>",
                ],
                "document": str(docx),
            }
        if completed.returncode != 0 or not saved.is_file() or not is_valid_docx(saved):
            stopped = kill_recorded_word(pid_file)
            return {
                "status": "manual_required",
                "reason": "Word не смог обновить поля: " + " ".join(output.split())[:500],
                "word_pid": {"source": source_kind, "pids": pids, "stopped": stopped},
                "manual_steps": MANUAL_STEPS,
                "document": str(docx),
            }
        pages = None
        for token in output.split():
            if token.startswith("pages="):
                try:
                    pages = int(token.split("=", 1)[1])
                except ValueError:
                    pages = None
        # Пока работал Word, пользователь мог открыть документ.
        blocker = write_blocker(docx)
        if blocker:
            raise FieldsError(f"{docx}: {blocker} — закрой файл в Word и повтори; ничего не изменено")
        backup = None
        project = find_project_root(docx.resolve())
        if project is not None:
            backup_dir = _unique_backup_dir(project)
            backup_dir.mkdir(parents=True)
            backup = backup_dir / docx.name
            shutil.copy2(docx, backup)
        staging = docx.with_name(f".{docx.stem}.fields-{uuid.uuid4().hex}.tmp")
        shutil.copyfile(saved, staging)
        try:
            os.replace(staging, docx)
        except OSError as error:
            raise FieldsError(f"не удалось заменить {docx} ({error}) — закрой файл в Word и повтори; "
                              "файл не изменён") from error
        replaced = True
        result: Dict[str, Any] = {
            "status": "updated",
            "document": str(docx),
            "pages": pages,
            "word_opened": word_opened,
            "backup": str(backup) if backup else None,
            "seconds": round(time.monotonic() - started, 1),
            "word_pid": {"source": source_kind, "pids": pids},
            "warnings": [],
            "next_steps": [
                "clean_docx_metadata.py <DOCX> --in-place",
                "vkr_audit.py <PROJECT_DIR> validate",
                "vkr_audit.py <PROJECT_DIR> snapshot",
            ],
        }
        if project is not None:
            try:
                result["fields_record"], recorded_pages = write_fields_record(project, docx, pages, word_opened)
            except (OSError, ValueError) as error:
                result["warnings"].append(f"не удалось записать {FIELDS_RECORD_PATH}: {error}")
            else:
                if pages is not None and recorded_pages != pages:
                    result["warnings"].append(
                        f"Word насчитал {pages} с., а в сохранённом файле записано {recorded_pages} с. "
                        f"(docProps/app.xml) — в {FIELDS_RECORD_PATH} записано число из файла"
                    )
                result["pages"] = recorded_pages
        return result
    finally:
        if staging is not None and staging.exists():
            try:
                staging.unlink()
            except OSError:
                pass
        if backup_dir is not None and not replaced:
            shutil.rmtree(backup_dir, ignore_errors=True)
        shutil.rmtree(work, ignore_errors=True)


def emit(result: Dict[str, Any], as_json: bool, output: Optional[Path]) -> None:
    payload = json.dumps(result, ensure_ascii=False, indent=2) + "\n"
    if output is not None:
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(payload, encoding="utf-8")
    if as_json:
        sys.stdout.write(payload)
        return
    if output is not None:
        return
    status = result.get("status")
    if status == "updated":
        print(f"Поля и оглавление обновлены: {result['document']} (страниц: {result.get('pages')})")
        if result.get("backup"):
            print(f"Прежняя версия: {result['backup']}")
        for warning in result.get("warnings") or []:
            print(f"ПРЕДУПРЕЖДЕНИЕ: {warning}")
        print("Дальше: очистка метаданных, vkr_audit.py validate, vkr_audit.py snapshot.")
    elif status == "manual_required":
        print(f"Автоматически обновить поля нельзя: {result.get('reason')}")
        for number, step in enumerate(result.get("manual_steps", []), 1):
            print(f"  {number}. {step}")
    elif status == "open_failed":
        print(f"Word не открыл документ: {result.get('reason')}", file=sys.stderr)
        for number, step in enumerate(result.get("next_steps", []), 1):
            print(f"  {number}. {step}", file=sys.stderr)
    else:
        print(f"Ошибка: {result.get('error')}", file=sys.stderr)


def main(argv: Optional[List[str]] = None) -> int:
    configure_output()
    parser = argparse.ArgumentParser(description="Обновить оглавление и поля DOCX через Microsoft Word")
    parser.add_argument("docx", type=Path, help="Путь к DOCX, обычно <PROJECT_DIR>/final/vkr.docx")
    parser.add_argument("--timeout", type=int, default=300, help="Тайм-аут работы Word в секундах (по умолчанию 300)")
    parser.add_argument("--json", action="store_true", dest="json_output", help="Вывести JSON-отчёт")
    parser.add_argument("-o", "--output", type=Path, help="Записать JSON-отчёт в UTF-8")
    args = parser.parse_args(argv)
    if args.timeout < 10:
        parser.error("--timeout должен быть не меньше 10 секунд")
    try:
        result = update_fields(args.docx, args.timeout, find_powershell())
    except FieldsError as error:
        result = {"status": "error", "error": str(error)}
        emit(result, args.json_output, args.output)
        return 2
    except OSError as error:
        result = {"status": "error", "error": f"ошибка ввода-вывода: {error}"}
        emit(result, args.json_output, args.output)
        return 2
    emit(result, args.json_output, args.output)
    if result.get("status") == "updated":
        return 0
    return 2 if result.get("status") == "open_failed" else 1


if __name__ == "__main__":
    raise SystemExit(main())
