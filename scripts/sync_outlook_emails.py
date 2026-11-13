#!/usr/bin/env python3
"""
Sync Outlook Emails to Project Directory
========================================
Universal tool that exports project-related emails (from a selected Outlook
folder and matching sent items / conversation threads) into structured Markdown
files with attachments.

Configuration priority (highest to lowest):
  1. CLI flags (--account, --folder, --keywords, --target-dir, --since, ...)
  2. Config JSON file (default: scripts/sync_outlook_config.json)
  3. Script defaults (no personal or project data hardcoded)
"""

from __future__ import annotations

import argparse
import datetime
import hashlib
import json
import re
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

# Base paths relative to this script: <project_root>/scripts/sync_outlook_emails.py
PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_TARGET_DIR = PROJECT_ROOT / "RESOURCES" / "emails"
DEFAULT_CONFIG_PATH = Path(__file__).resolve().parent / "sync_outlook_config.json"
SYNC_INDEX_FILENAME = ".sync_index.json"

INVALID_FILENAME_CHARS = re.compile(r'[\\/*?:"<>|]')
DEFAULT_ACCOUNT = ""
DEFAULT_INBOX_SUBFOLDER = ""
DEFAULT_SENT_KEYWORDS: List[str] = []


def sanitize_filename(text: str, max_len: int = 60) -> str:
    """
    Sanitize text to create a safe Windows/POSIX filename.
    Replaces disallowed characters and truncates to max_len.
    """
    if not text:
        return "Brak_tematu"

    clean = re.sub(r"[\r\n\t]+", " ", text)
    clean = INVALID_FILENAME_CHARS.sub("_", clean)
    clean = re.sub(r"\s+", " ", clean).strip()
    clean = re.sub(r"_+", "_", clean).strip(" ._")

    if not clean:
        clean = "Wiadomosc"

    if len(clean) > max_len:
        clean = clean[:max_len].rstrip(" ._")

    return clean or "Wiadomosc"


def sanitize_attachment_filename(filename: str, max_len: int = 80) -> str:
    """Sanitize an attachment filename while preserving its extension."""
    p = Path(filename)
    ext = p.suffix
    stem = p.stem
    clean_stem = sanitize_filename(stem, max_len=max(10, max_len - len(ext)))
    return f"{clean_stem}{ext}"


def _escape_yaml_str(value: str) -> str:
    """Escape characters that would break double-quoted YAML scalar values."""
    single_line = re.sub(r"[\r\n]+", " ", value).strip()
    return single_line.replace("\\", "\\\\").replace('"', '\\"')


def format_markdown_email(
    metadata: Dict[str, str],
    body: str,
    attachments: Optional[List[str]] = None,
    rel_attach_dir: str = "",
) -> str:
    """Format email metadata and body into a structured Markdown document."""
    attachments = attachments or []

    date_val = metadata.get("date", "")
    direction_val = metadata.get("direction", "IN")
    from_val = metadata.get("from", "")
    to_val = metadata.get("to", "")
    cc_val = metadata.get("cc", "")
    raw_subject = metadata.get("subject", "Brak tematu") or "Brak tematu"
    subject_val = re.sub(r"[\r\n]+", " ", raw_subject).strip() or "Brak tematu"
    entry_id_val = metadata.get("entry_id", "")

    lines = [
        "---",
        f'date: "{_escape_yaml_str(date_val)}"',
        f'direction: "{_escape_yaml_str(direction_val)}"',
        f'from: "{_escape_yaml_str(from_val)}"',
        f'to: "{_escape_yaml_str(to_val)}"',
        f'cc: "{_escape_yaml_str(cc_val)}"',
        f'subject: "{_escape_yaml_str(subject_val)}"',
        f'entry_id: "{_escape_yaml_str(entry_id_val)}"',
        "---",
        "",
        f"# {subject_val}",
        "",
        f"- **Data:** {date_val}",
        f"- **Kierunek:** {'Odebrane' if direction_val == 'IN' else 'Wysłane'}",
        f"- **Od:** {from_val}",
        f"- **Do:** {to_val}",
    ]

    if cc_val:
        lines.append(f"- **DW:** {cc_val}")

    lines.append("")

    if attachments:
        lines.append("## Załączniki")
        for att in attachments:
            link = f"{rel_attach_dir}/{att}" if rel_attach_dir else att
            lines.append(f"- [{att}]({link})")
        lines.append("")

    lines.append("## Treść wiadomości")
    lines.append("")
    clean_body = body.strip() if body else "*(Brak treści tekstowej)*"
    lines.append(clean_body)
    lines.append("")

    return "\n".join(lines)


class SyncIndexManager:
    """Tracks synchronized emails by Outlook EntryID to ensure idempotency."""

    def __init__(self, index_file_path: Path):
        self.index_file = Path(index_file_path)
        self.entries: Dict[str, Dict[str, Any]] = {}
        self.load()

    def load(self) -> None:
        if self.index_file.exists():
            try:
                with open(self.index_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    self.entries = data if isinstance(data, dict) else {}
            except Exception as e:
                print(f"[OSTRZEŻENIE] Nie udało się odczytać indeksu {self.index_file}: {e}")
                self.entries = {}
        else:
            self.entries = {}

    def is_synced(self, entry_id: str) -> bool:
        return entry_id in self.entries

    def get_filename(self, entry_id: str) -> Optional[str]:
        entry = self.entries.get(entry_id)
        if entry and isinstance(entry, dict):
            return entry.get("filename")
        return None

    def has_other_entry_at(self, rel_filename: str, entry_id: str) -> bool:
        """Return True if another EntryID is already mapped to rel_filename."""
        for existing_id, meta in self.entries.items():
            if existing_id != entry_id and isinstance(meta, dict):
                if meta.get("filename") == rel_filename:
                    return True
        return False

    def record(self, entry_id: str, filename: str) -> None:
        self.entries[entry_id] = {
            "filename": filename,
            "synced_at": datetime.datetime.now().isoformat(),
        }

    def save(self) -> None:
        self.index_file.parent.mkdir(parents=True, exist_ok=True)
        with open(self.index_file, "w", encoding="utf-8") as f:
            json.dump(self.entries, f, ensure_ascii=False, indent=2)


def load_config(config_path: Path) -> Dict[str, Any]:
    """Load optional JSON configuration file."""
    if not config_path.exists():
        return {}
    try:
        with open(config_path, "r", encoding="utf-8") as f:
            data = json.load(f)
            return data if isinstance(data, dict) else {}
    except Exception as e:
        print(f"[OSTRZEŻENIE] Błąd odczytu konfiguracji {config_path}: {e}")
        return {}


def normalize_keywords(raw: Any) -> List[str]:
    """Normalize keywords from a list or comma-separated string."""
    if not raw:
        return []
    if isinstance(raw, str):
        return [k.strip() for k in raw.split(",") if k.strip()]
    if isinstance(raw, list):
        return [str(k).strip() for k in raw if str(k).strip()]
    return []


def parse_since_date(since_raw: Optional[str]) -> Optional[datetime.date]:
    """Parse YYYY-MM-DD date string into a datetime.date object."""
    if not since_raw:
        return None
    try:
        return datetime.date.fromisoformat(since_raw.strip())
    except ValueError:
        print(f"[BŁĄD] Nieprawidłowy format daty --since '{since_raw}'. Użyj YYYY-MM-DD.")
        sys.exit(1)


def get_outlook_namespace():
    """Connect to Classic Microsoft Outlook MAPI namespace via COM."""
    try:
        import win32com.client
    except ImportError as e:
        raise RuntimeError(
            "Brak biblioteki 'pywin32'. Zainstaluj ją poleceniem: pip install pywin32"
        ) from e

    outlook = win32com.client.Dispatch("Outlook.Application")
    return outlook.GetNamespace("MAPI")


def find_store(namespace, account_email: str = ""):
    """
    Find the Outlook store matching account_email, or return the default store
    when account_email is empty.
    """
    if account_email:
        target_lower = account_email.strip().lower()
        available = []
        for store in namespace.Stores:
            display = getattr(store, "DisplayName", "")
            available.append(display)
            if target_lower in display.lower():
                return store.GetRootFolder()

        raise ValueError(
            f"Nie znaleziono konta '{account_email}' w Outlooku. "
            f"Dostępne konta: {', '.join(available) or '(brak)'}"
        )

    return namespace.GetDefaultFolder(6).Parent  # 6 = olFolderInbox


def _find_direct_child(parent_folder, name: str):
    target = name.strip().lower()
    for sub in parent_folder.Folders:
        if getattr(sub, "Name", "").lower() == target:
            return sub
    return None


def find_folder_by_path(parent_folder, folder_spec: str):
    """
    Locate a subfolder by slash/backslash path (e.g. 'Projekty/KlientX')
    or recursively by single folder name.
    """
    if not folder_spec:
        return None

    parts = [p.strip() for p in re.split(r"[/\\]+", folder_spec) if p.strip()]
    if not parts:
        return None

    if len(parts) > 1:
        current = parent_folder
        for part in parts:
            current = _find_direct_child(current, part)
            if current is None:
                return None
        return current

    target_name = parts[0]
    direct = _find_direct_child(parent_folder, target_name)
    if direct is not None:
        return direct

    # Breadth-first recursive search across nested subfolders
    queue = list(parent_folder.Folders)
    while queue:
        sub = queue.pop(0)
        if getattr(sub, "Name", "").lower() == target_name.lower():
            return sub
        try:
            queue.extend(list(sub.Folders))
        except Exception:
            pass
    return None


def list_available_folders(namespace, account_email: str = "", max_depth: int = 3) -> None:
    """Print available Outlook stores and folder hierarchy to help configure sync."""
    print("=" * 70)
    print("Dostępne konta i foldery w Microsoft Outlook:")
    print("=" * 70)

    def _walk(folder, prefix: str = "", depth: int = 0) -> None:
        if depth > max_depth:
            return
        try:
            for sub in folder.Folders:
                count = getattr(sub.Items, "Count", "?")
                print(f"{prefix}- {sub.Name} ({count})")
                _walk(sub, prefix + "  ", depth + 1)
        except Exception:
            pass

    for store in namespace.Stores:
        display = getattr(store, "DisplayName", "(bez nazwy)")
        if account_email and account_email.lower() not in display.lower():
            continue
        print(f"\n[Konto] {display}")
        try:
            root = store.GetRootFolder()
            _walk(root, prefix="  ", depth=0)
        except Exception as e:
            print(f"  (nie można odczytać folderów: {e})")


def extract_recipient_string(recipients) -> Tuple[str, str]:
    """Extract 'To' and 'CC' recipient string lists."""
    to_list = []
    cc_list = []
    try:
        for r in recipients:
            name = getattr(r, "Name", "") or ""
            address = getattr(r, "Address", "") or ""
            display = f"{name} <{address}>" if (address and address != name) else name
            if not display:
                continue
            r_type = getattr(r, "Type", 1)  # 1 = olTo, 2 = olCC
            if r_type == 2:
                cc_list.append(display)
            else:
                to_list.append(display)
    except Exception:
        pass
    return ", ".join(to_list), ", ".join(cc_list)


def parse_datetime(dt_raw) -> datetime.datetime:
    """Normalize datetime object from Outlook pywintypes/datetime."""
    if isinstance(dt_raw, datetime.datetime):
        return dt_raw.replace(tzinfo=None)
    if hasattr(dt_raw, "year") and hasattr(dt_raw, "month") and hasattr(dt_raw, "day"):
        try:
            return datetime.datetime(
                int(dt_raw.year),
                int(dt_raw.month),
                int(dt_raw.day),
                int(getattr(dt_raw, "hour", 0)),
                int(getattr(dt_raw, "minute", 0)),
                int(getattr(dt_raw, "second", 0)),
            )
        except Exception:
            pass
    try:
        return datetime.datetime.fromtimestamp(int(dt_raw))
    except Exception:
        return datetime.datetime.now()


def conversation_key(item) -> Optional[str]:
    """
    Return the 16-byte conversation (thread) key for an Outlook item.
    The first 32 hex characters of ConversationIndex identify the thread.
    """
    try:
        val = str(item.ConversationIndex)
        return val[:32] if len(val) >= 32 else (val or None)
    except Exception:
        return None


def unique_md_relative_path(
    target_dir: Path,
    month_dir_name: str,
    base_name: str,
    entry_id: str,
    index_mgr: SyncIndexManager,
) -> Tuple[str, str]:
    """
    Compute a collision-free relative Markdown path and attachment folder name.
    Appends a short hash of entry_id only when another email already occupies the name.
    """
    candidate_rel = f"{month_dir_name}/{base_name}.md"
    candidate_path = target_dir / candidate_rel
    if (
        index_mgr.has_other_entry_at(candidate_rel, entry_id)
        or (candidate_path.exists() and index_mgr.get_filename(entry_id) != candidate_rel)
    ):
        short_hash = hashlib.sha1(entry_id.encode("utf-8", errors="ignore")).hexdigest()[:6]
        disambiguated_base = f"{base_name}_{short_hash}"
        return f"{month_dir_name}/{disambiguated_base}.md", f"{disambiguated_base}_attachments"
    return candidate_rel, f"{base_name}_attachments"


def process_mail_item(
    mail,
    direction: str,
    target_dir: Path,
    index_mgr: SyncIndexManager,
    account_fallback: str = "",
    dry_run: bool = False,
    force: bool = False,
) -> Tuple[bool, str, datetime.datetime]:
    """
    Process a single Outlook MailItem and save Markdown + attachments.
    Returns (exported: bool, status_msg: str, item_datetime: datetime).
    """
    msg_class = getattr(mail, "Class", None)
    if msg_class != 43:  # 43 = olMail
        return False, "Pomijam element niebędący mailem", datetime.datetime.now()

    try:
        entry_id = mail.EntryID
        subject = getattr(mail, "Subject", None) or "Brak tematu"

        if direction == "IN":
            dt = parse_datetime(mail.ReceivedTime)
            sender_name = getattr(mail, "SenderName", "") or ""
            sender_email = getattr(mail, "SenderEmailAddress", "") or ""
            from_str = (
                f"{sender_name} <{sender_email}>"
                if (sender_email and sender_email != sender_name)
                else sender_name
            )
        else:
            dt = parse_datetime(mail.SentOn)
            from_str = getattr(mail, "SenderName", "") or account_fallback

        if not force and index_mgr.is_synced(entry_id):
            existing_fn = index_mgr.get_filename(entry_id)
            if existing_fn and (target_dir / existing_fn).exists():
                return False, f"Już zsynchronizowano: {existing_fn}", dt

        to_str, cc_str = extract_recipient_string(mail.Recipients)
        body = getattr(mail, "Body", "") or ""

        dt_str = dt.strftime("%Y-%m-%d %H:%M:%S")
        file_prefix = dt.strftime("%Y-%m-%d_%H%M")
        month_dir_name = dt.strftime("%Y-%m")
        sanitized_subject = sanitize_filename(subject, max_len=60)
        base_name = f"{file_prefix}_[{direction}]_{sanitized_subject}"

        md_filename, attach_folder_name = unique_md_relative_path(
            target_dir, month_dir_name, base_name, entry_id, index_mgr
        )
        md_path = target_dir / md_filename
        attach_folder_path = target_dir / month_dir_name / attach_folder_name

        metadata = {
            "date": dt_str,
            "direction": direction,
            "from": from_str,
            "to": to_str,
            "cc": cc_str,
            "subject": subject,
            "entry_id": entry_id,
        }

        attachment_names: List[str] = []
        attachments_count = getattr(mail.Attachments, "Count", 0)
        if attachments_count > 0:
            for i in range(1, attachments_count + 1):
                att = mail.Attachments.Item(i)
                att_name = getattr(att, "FileName", f"zalacznik_{i}")
                if not att_name:
                    continue
                safe_att_name = sanitize_attachment_filename(att_name, max_len=80)
                attachment_names.append(safe_att_name)

                if not dry_run:
                    attach_folder_path.mkdir(parents=True, exist_ok=True)
                    save_file_path = attach_folder_path / safe_att_name
                    try:
                        att.SaveAsFile(str(save_file_path))
                    except Exception as save_err:
                        print(f"  [BŁĄD] Nie udało się zapisać załącznika {safe_att_name}: {save_err}")

        if dry_run:
            return (
                True,
                f"[DRY-RUN] Zapisałby: {md_filename} (załączników: {len(attachment_names)})",
                dt,
            )

        rel_dir = attach_folder_name if attachment_names else ""
        md_content = format_markdown_email(metadata, body, attachment_names, rel_dir)
        md_path.parent.mkdir(parents=True, exist_ok=True)
        with open(md_path, "w", encoding="utf-8") as f:
            f.write(md_content)

        index_mgr.record(entry_id, md_filename)
        return True, f"Zapisano: {md_filename}", dt

    except Exception as e:
        return False, f"Błąd przetwarzania wiadomości: {e}", datetime.datetime.now()


def _locate_inbox_and_target(namespace, root_folder, inbox_subfolder: str):
    """Find Inbox and optionally the requested target folder inside Inbox or store root."""
    inbox_folder = (
        find_folder_by_path(root_folder, "Skrzynka odbiorcza")
        or find_folder_by_path(root_folder, "Inbox")
    )
    if not inbox_folder:
        try:
            inbox_folder = namespace.GetDefaultFolder(6)
        except Exception:
            inbox_folder = None

    if not inbox_subfolder:
        return inbox_folder, None

    target_folder = None
    if inbox_folder:
        target_folder = find_folder_by_path(inbox_folder, inbox_subfolder)
    if not target_folder:
        target_folder = find_folder_by_path(root_folder, inbox_subfolder)

    if not target_folder:
        raise ValueError(
            f"Nie znaleziono folderu '{inbox_subfolder}'. "
            "Uruchom skrypt z flagą --list-folders, aby sprawdzić dokładną nazwę."
        )

    return inbox_folder, target_folder


def _locate_sent_folders(namespace, root_folder) -> List[Any]:
    """Locate Sent Items folders in the target account store."""
    sent_folders = []
    seen_names = set()
    for sf_name in ["Elementy wysłane", "Sent", "Sent Items"]:
        sf = _find_direct_child(root_folder, sf_name)
        if sf and sf.Name not in seen_names:
            seen_names.add(sf.Name)
            sent_folders.append(sf)

    if not sent_folders:
        try:
            default_sent = namespace.GetDefaultFolder(5)  # 5 = olFolderSentMail
            if default_sent:
                sent_folders.append(default_sent)
        except Exception:
            pass
    return sent_folders


def run_sync(
    target_dir: Path = DEFAULT_TARGET_DIR,
    account_email: str = DEFAULT_ACCOUNT,
    inbox_subfolder: str = DEFAULT_INBOX_SUBFOLDER,
    sent_keywords: Optional[List[str]] = None,
    since_date: Optional[datetime.date] = None,
    dry_run: bool = False,
    force: bool = False,
    limit: Optional[int] = None,
    thread_match: bool = True,
) -> Dict[str, int]:
    """Main synchronization runner. Returns summary counters."""
    sent_keywords = normalize_keywords(sent_keywords)
    target_dir = Path(target_dir)

    if not inbox_subfolder and not sent_keywords:
        print(
            "[BŁĄD] Nie podano ani folderu odebranych (--folder), ani słów kluczowych (--keywords).\n"
            "Skonfiguruj plik scripts/sync_outlook_config.json lub podaj parametry w CLI.\n"
            "Podgląd dostępnych folderów: python scripts/sync_outlook_emails.py --list-folders"
        )
        sys.exit(1)

    print("=" * 70)
    print("Synchronizacja e-maili z Microsoft Outlook do Markdown")
    print(f"Konto:                {account_email or '(domyślne konto Outlook)'}")
    print(f"Folder odebranych:    {inbox_subfolder or '(pominięto - tylko słowa kluczowe)'}")
    print(f"Słowa kluczowe wysł.: {', '.join(sent_keywords) if sent_keywords else '(brak - tylko wątki z folderu)'}")
    print(f"Od daty:              {since_date.isoformat() if since_date else '(bez limitu daty)'}")
    print(f"Katalog docelowy:     {target_dir}")
    print(f"Tryb:                 {'DRY-RUN (testowy)' if dry_run else 'Zapis na dysk'}")
    print("=" * 70)

    try:
        namespace = get_outlook_namespace()
    except Exception as e:
        print(f"[BŁĄD KRYTYCZNY] Nie można połączyć się z aplikacją Outlook: {e}")
        print("Upewnij się, że używasz klasycznego Microsoft Outlook (Desktop) na Windows i jest uruchomiony.")
        sys.exit(1)

    try:
        root_folder = find_store(namespace, account_email)
        print(f"Magazyn główny: {root_folder.Name}")
        _, target_inbox = _locate_inbox_and_target(namespace, root_folder, inbox_subfolder)
    except ValueError as e:
        print(f"[BŁĄD] {e}")
        sys.exit(1)

    sent_folders = _locate_sent_folders(namespace, root_folder)

    index_file = target_dir / SYNC_INDEX_FILENAME
    index_mgr = SyncIndexManager(index_file)

    checked_count = 0
    synced_count = 0
    skipped_count = 0
    folder_conv_keys = set()

    # 1. Process Received emails from target folder
    if target_inbox is not None:
        print(f"\n--- Przetwarzanie folderu: {target_inbox.Name} (elementów: {target_inbox.Items.Count}) ---")
        inbox_items = target_inbox.Items
        try:
            inbox_items.Sort("[ReceivedTime]", True)
        except Exception:
            pass

        consecutive_older = 0
        for item in inbox_items:
            if limit and synced_count >= limit:
                print(f"Osiągnięto limit {limit} nowych wiadomości.")
                break

            if thread_match:
                k = conversation_key(item)
                if k:
                    folder_conv_keys.add(k)

            # Check date filter early if MailItem
            if since_date and getattr(item, "Class", None) == 43:
                item_dt = parse_datetime(getattr(item, "ReceivedTime", None))
                if item_dt.date() < since_date:
                    consecutive_older += 1
                    if consecutive_older >= 10:
                        print(f"Zatrzymano skanowanie folderu na wiadomościach starszych niż {since_date}.")
                        break
                    continue
                consecutive_older = 0

            checked_count += 1
            ok, msg, _ = process_mail_item(
                item,
                "IN",
                target_dir,
                index_mgr,
                account_fallback=account_email,
                dry_run=dry_run,
                force=force,
            )
            if ok:
                synced_count += 1
                print(f"  [+] {msg}")
            else:
                skipped_count += 1
                if "Już zsynchronizowano" not in msg:
                    print(f"  [-] {msg}")

        # Ensure full thread coverage even if loop stopped early via --limit
        if thread_match and not folder_conv_keys:
            for item in target_inbox.Items:
                k = conversation_key(item)
                if k:
                    folder_conv_keys.add(k)

    # 2. Process Sent emails (matching keywords or conversation threads from target folder)
    if sent_folders and (sent_keywords or (thread_match and folder_conv_keys)):
        match_parts = []
        if sent_keywords:
            match_parts.append("słowa kluczowe")
        if thread_match and folder_conv_keys:
            match_parts.append(f"wątki z folderu ({len(folder_conv_keys)})")
        print(f"\n--- Przetwarzanie wiadomości wysłanych ({' + '.join(match_parts)}) ---")

        keywords_lower = [k.lower() for k in sent_keywords]

        for sf in sent_folders:
            if limit and synced_count >= limit:
                break
            print(f"Przeszukiwanie folderu: {sf.Name}...")
            sent_items = sf.Items
            try:
                sent_items.Sort("[SentOn]", True)
            except Exception:
                pass

            consecutive_older = 0
            for item in sent_items:
                if limit and synced_count >= limit:
                    print(f"Osiągnięto limit {limit} nowych wiadomości.")
                    break

                if since_date and getattr(item, "Class", None) == 43:
                    item_dt = parse_datetime(getattr(item, "SentOn", None))
                    if item_dt.date() < since_date:
                        consecutive_older += 1
                        if consecutive_older >= 10:
                            print(f"Zatrzymano skanowanie wysłanych na dacie {since_date}.")
                            break
                        continue
                    consecutive_older = 0

                subject = getattr(item, "Subject", "") or ""
                subj_lower = subject.lower()

                item_conv_key = conversation_key(item)
                matches_kw = any(k in subj_lower for k in keywords_lower) if keywords_lower else False
                matches_thread = bool(thread_match) and bool(item_conv_key) and item_conv_key in folder_conv_keys

                if matches_kw or matches_thread:
                    checked_count += 1
                    ok, msg, _ = process_mail_item(
                        item,
                        "OUT",
                        target_dir,
                        index_mgr,
                        account_fallback=account_email,
                        dry_run=dry_run,
                        force=force,
                    )
                    if ok:
                        synced_count += 1
                        print(f"  [+] {msg}")
                    else:
                        skipped_count += 1
                        if "Już zsynchronizowano" not in msg:
                            print(f"  [-] {msg}")

    if not dry_run:
        index_mgr.save()
        print(f"\nZapisano indeks: {index_file}")

    print("\n" + "=" * 70)
    print("PODSUMOWANIE:")
    print(f"  Sprawdzonych dopasowań:    {checked_count}")
    print(f"  Wyeksportowanych (nowych): {synced_count}")
    print(f"  Pominiętych:               {skipped_count}")
    print("=" * 70)

    return {
        "checked": checked_count,
        "synced": synced_count,
        "skipped": skipped_count,
    }


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Uniwersalna synchronizacja maili z Microsoft Outlook do Markdown."
    )
    parser.add_argument(
        "--config",
        type=str,
        default=str(DEFAULT_CONFIG_PATH),
        help="Ścieżka do pliku konfiguracyjnego JSON (domyślnie scripts/sync_outlook_config.json)",
    )
    parser.add_argument(
        "--target-dir",
        type=str,
        default=None,
        help="Katalog docelowy na pliki Markdown i załączniki",
    )
    parser.add_argument(
        "--account",
        type=str,
        default=None,
        help="Adres e-mail lub nazwa konta w Outlook (opcjonalnie, domyślnie główne konto)",
    )
    parser.add_argument(
        "--folder",
        "--inbox-subfolder",
        dest="inbox_subfolder",
        type=str,
        default=None,
        help="Nazwa folderu lub ścieżka (np. 'Projekty/Nazwa') w Outlooku",
    )
    parser.add_argument(
        "--keywords",
        nargs="+",
        default=None,
        help="Słowa kluczowe w temacie dla wiadomości wysłanych",
    )
    parser.add_argument(
        "--since",
        type=str,
        default=None,
        help="Synchronizuj wiadomości od daty YYYY-MM-DD (przyspiesza skanowanie dużych skrzynek)",
    )
    parser.add_argument(
        "--list-folders",
        action="store_true",
        help="Wyświetl dostępne konta i foldery w Outlooku i zakończ",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Tryb testowy - pokazuje co zostałoby zapisane, bez modyfikacji dysku",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="Wymuś ponowny eksport nawet jeśli wiadomość jest już w indeksie",
    )
    parser.add_argument(
        "--limit",
        type=int,
        default=None,
        help="Maksymalna liczba nowych wiadomości do wyeksportowania",
    )
    parser.add_argument(
        "--no-thread-match",
        action="store_true",
        help="Wyłącz dopasowanie wysłanych po wątkach z folderu (tylko słowa kluczowe)",
    )
    return parser


def main() -> None:
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

    parser = build_parser()
    args = parser.parse_args()

    cfg = load_config(Path(args.config))

    account = args.account if args.account is not None else cfg.get("account", DEFAULT_ACCOUNT)

    if args.list_folders:
        try:
            namespace = get_outlook_namespace()
            list_available_folders(namespace, account_email=account)
        except Exception as e:
            print(f"[BŁĄD] Nie można odczytać folderów Outlooka: {e}")
            sys.exit(1)
        return

    raw_target = args.target_dir if args.target_dir is not None else cfg.get("target_dir")
    if raw_target:
        target_path = Path(raw_target)
        if not target_path.is_absolute():
            target_path = PROJECT_ROOT / target_path
    else:
        target_path = DEFAULT_TARGET_DIR

    inbox_subfolder = (
        args.inbox_subfolder
        if args.inbox_subfolder is not None
        else cfg.get("inbox_subfolder", DEFAULT_INBOX_SUBFOLDER)
    )
    keywords = (
        args.keywords
        if args.keywords is not None
        else normalize_keywords(cfg.get("sent_keywords", DEFAULT_SENT_KEYWORDS))
    )
    since_str = args.since if args.since is not None else cfg.get("since", "")
    since_date = parse_since_date(since_str)

    if args.no_thread_match:
        thread_match = False
    else:
        thread_match = bool(cfg.get("thread_match", True))

    run_sync(
        target_dir=target_path,
        account_email=account,
        inbox_subfolder=inbox_subfolder,
        sent_keywords=keywords,
        since_date=since_date,
        dry_run=args.dry_run,
        force=args.force,
        limit=args.limit,
        thread_match=thread_match,
    )


if __name__ == "__main__":
    main()
