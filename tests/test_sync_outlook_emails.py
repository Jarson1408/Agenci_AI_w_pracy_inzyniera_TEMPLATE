import datetime
import json
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

SCRIPTS_DIR = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS_DIR))

import sync_outlook_emails as sync  # noqa: E402


class FakeCollection:
    def __init__(self, items):
        self._items = list(items)

    @property
    def Count(self):
        return len(self._items)

    def Item(self, idx):
        # Outlook COM Collections are 1-indexed
        return self._items[idx - 1]

    def Sort(self, _field, _desc=False):
        return None

    def __iter__(self):
        return iter(self._items)


class FakeAttachment:
    def __init__(self, filename, content="dummy"):
        self.FileName = filename
        self._content = content

    def SaveAsFile(self, path):
        Path(path).write_text(self._content, encoding="utf-8")


class FakeRecipient:
    def __init__(self, name, address="", r_type=1):
        self.Name = name
        self.Address = address
        self.Type = r_type


class FakeMailItem:
    def __init__(
        self,
        entry_id,
        subject,
        body="Tresc testowa",
        dt=None,
        sender_name="Jan Kowalski",
        sender_email="jan@example.com",
        recipients=None,
        attachments=None,
        conv_index="01D9AABBCCDDEEFF00112233445566778899",
        msg_class=43,
    ):
        self.Class = msg_class
        self.EntryID = entry_id
        self.Subject = subject
        self.Body = body
        dt = dt or datetime.datetime(2026, 5, 10, 14, 30, 0)
        self.ReceivedTime = dt
        self.SentOn = dt
        self.SenderName = sender_name
        self.SenderEmailAddress = sender_email
        self.Recipients = FakeCollection(recipients or [FakeRecipient("Anna Nowak", "anna@example.com", 1)])
        self.Attachments = FakeCollection(attachments or [])
        self.ConversationIndex = conv_index


class FakeFolder:
    def __init__(self, name, items=None, subfolders=None):
        self.Name = name
        self.Parent = None
        self.Items = FakeCollection(items or [])
        subs = list(subfolders or [])
        for sub in subs:
            sub.Parent = self
        self.Folders = FakeCollection(subs)


class FakeStore:
    def __init__(self, display_name, root_folder):
        self.DisplayName = display_name
        self._root = root_folder

    def GetRootFolder(self):
        return self._root


class FakeNamespace:
    def __init__(self, stores, default_inbox, default_sent):
        self.Stores = FakeCollection(stores)
        self._default_inbox = default_inbox
        self._default_sent = default_sent

    def GetDefaultFolder(self, folder_id):
        if folder_id == 6:
            return self._default_inbox
        if folder_id == 5:
            return self._default_sent
        raise ValueError(f"Unknown folder id {folder_id}")


class TestSyncOutlookEmails(unittest.TestCase):
    def test_sanitize_filename_and_attachment(self):
        self.assertEqual(sync.sanitize_filename(""), "Brak_tematu")
        self.assertEqual(sync.sanitize_filename('FW: Raport <v1> / "test"?'), "FW_ Raport _v1_ _ _test")
        att = sync.sanitize_attachment_filename("Specyfikacja: PCB/v2?.pdf")
        self.assertTrue(att.endswith(".pdf"))
        self.assertNotIn(":", att)
        self.assertNotIn("/", att)

    def test_format_markdown_email_escapes_yaml_and_links_attachments(self):
        md = sync.format_markdown_email(
            metadata={
                "date": "2026-05-10 14:30:00",
                "direction": "IN",
                "from": 'Jan "Tester" <jan@example.com>',
                "to": "anna@example.com",
                "cc": "biuro@example.com",
                "subject": 'Temat:\n"Wazny"',
                "entry_id": "ID-1",
            },
            body="Linia 1\nLinia 2",
            attachments=["schemat.pdf"],
            rel_attach_dir="2026-05-10_1430_[IN]_Temat_attachments",
        )
        self.assertIn('subject: "Temat: \\"Wazny\\""', md)
        self.assertIn("- **DW:** biuro@example.com", md)
        self.assertIn("- [schemat.pdf](2026-05-10_1430_[IN]_Temat_attachments/schemat.pdf)", md)
        self.assertIn("Linia 1\nLinia 2", md)

    def test_find_folder_by_path_nested_and_recursive(self):
        target = FakeFolder("Projekt_XYZ")
        projects = FakeFolder("Projekty", subfolders=[target])
        inbox = FakeFolder("Skrzynka odbiorcza", subfolders=[projects])
        root = FakeFolder("Root", subfolders=[inbox])

        self.assertIs(sync.find_folder_by_path(inbox, "Projekty/Projekt_XYZ"), target)
        self.assertIs(sync.find_folder_by_path(inbox, "Projekty\\Projekt_XYZ"), target)
        self.assertIs(sync.find_folder_by_path(root, "Projekt_XYZ"), target)
        self.assertIsNone(sync.find_folder_by_path(root, "Nieistniejacy"))

    def test_load_config_and_keywords(self):
        with tempfile.TemporaryDirectory() as tmp:
            cfg_path = Path(tmp) / "cfg.json"
            self.assertEqual(sync.load_config(cfg_path), {})
            cfg_path.write_text(json.dumps({"inbox_subfolder": "Alpha", "sent_keywords": ["A", "B"]}), encoding="utf-8")
            loaded = sync.load_config(cfg_path)
            self.assertEqual(loaded["inbox_subfolder"], "Alpha")
        self.assertEqual(sync.normalize_keywords("TAG1, TAG2 , "), ["TAG1", "TAG2"])

    def test_run_sync_exports_inbox_thread_sent_and_is_idempotent(self):
        conv_thread = "01D9AABBCCDDEEFF0011223344556677"
        in_mail = FakeMailItem(
            entry_id="ENTRY_IN_1",
            subject="Zapytanie ofertowe",
            dt=datetime.datetime(2026, 5, 10, 9, 15, 0),
            attachments=[FakeAttachment("spec.pdf", "PDFDATA")],
            conv_index=conv_thread + "0011",
        )
        # Sent reply has NO keyword in subject, should still match via thread
        sent_thread_reply = FakeMailItem(
            entry_id="ENTRY_OUT_THREAD",
            subject="Re: Odpowiedz",
            dt=datetime.datetime(2026, 5, 10, 10, 0, 0),
            conv_index=conv_thread + "0022",
        )
        # Sent email with keyword in subject
        sent_kw_mail = FakeMailItem(
            entry_id="ENTRY_OUT_KW",
            subject="[PROJ-X] Wysylka plikow",
            dt=datetime.datetime(2026, 5, 11, 12, 0, 0),
            conv_index="99999999999999999999999999999999",
        )
        # Unrelated sent email (must be ignored)
        sent_unrelated = FakeMailItem(
            entry_id="ENTRY_OUT_OTHER",
            subject="Obiad w piatek",
            dt=datetime.datetime(2026, 5, 11, 13, 0, 0),
            conv_index="88888888888888888888888888888888",
        )

        proj_folder = FakeFolder("MojProjekt", items=[in_mail])
        inbox_folder = FakeFolder("Inbox", subfolders=[proj_folder])
        sent_folder = FakeFolder("Sent Items", items=[sent_thread_reply, sent_kw_mail, sent_unrelated])
        root_folder = FakeFolder("Root", subfolders=[inbox_folder, sent_folder])
        store = FakeStore("user@example.com", root_folder)
        ns = FakeNamespace([store], inbox_folder, sent_folder)

        with tempfile.TemporaryDirectory() as tmp:
            target_dir = Path(tmp) / "emails"
            with mock.patch.object(sync, "get_outlook_namespace", return_value=ns):
                # Dry run writes nothing
                dry_stats = sync.run_sync(
                    target_dir=target_dir,
                    account_email="",
                    inbox_subfolder="MojProjekt",
                    sent_keywords=["PROJ-X"],
                    dry_run=True,
                )
                self.assertEqual(dry_stats["synced"], 3)
                self.assertFalse(target_dir.exists())

                # Real run writes 3 markdown files + 1 attachment + index
                stats = sync.run_sync(
                    target_dir=target_dir,
                    account_email="",
                    inbox_subfolder="MojProjekt",
                    sent_keywords=["PROJ-X"],
                    dry_run=False,
                )
                self.assertEqual(stats["synced"], 3)
                md_files = list((target_dir / "2026-05").glob("*.md"))
                self.assertEqual(len(md_files), 3)
                att_files = list((target_dir / "2026-05").rglob("spec.pdf"))
                self.assertEqual(len(att_files), 1)
                self.assertTrue((target_dir / ".sync_index.json").exists())

                # Second run skips already synced messages
                stats_again = sync.run_sync(
                    target_dir=target_dir,
                    account_email="",
                    inbox_subfolder="MojProjekt",
                    sent_keywords=["PROJ-X"],
                    dry_run=False,
                )
                self.assertEqual(stats_again["synced"], 0)
                self.assertEqual(stats_again["skipped"], 3)

    def test_collision_same_minute_same_subject_adds_hash(self):
        dt = datetime.datetime(2026, 6, 1, 8, 0, 0)
        m1 = FakeMailItem("ENTRY_A", "Status", dt=dt)
        m2 = FakeMailItem("ENTRY_B", "Status", dt=dt)
        with tempfile.TemporaryDirectory() as tmp:
            target_dir = Path(tmp)
            idx = sync.SyncIndexManager(target_dir / ".sync_index.json")
            ok1, _, _ = sync.process_mail_item(m1, "IN", target_dir, idx)
            ok2, _, _ = sync.process_mail_item(m2, "IN", target_dir, idx)
            self.assertTrue(ok1)
            self.assertTrue(ok2)
            md_files = list((target_dir / "2026-06").glob("*.md"))
            self.assertEqual(len(md_files), 2)

    def test_since_date_skips_older_emails(self):
        old_mail = FakeMailItem("ENTRY_OLD", "Stary mail", dt=datetime.datetime(2025, 1, 10, 10, 0, 0))
        new_mail = FakeMailItem("ENTRY_NEW", "Nowy mail", dt=datetime.datetime(2026, 5, 1, 10, 0, 0))
        proj_folder = FakeFolder("MojProjekt", items=[new_mail, old_mail])
        inbox_folder = FakeFolder("Inbox", subfolders=[proj_folder])
        sent_folder = FakeFolder("Sent Items", items=[])
        root_folder = FakeFolder("Root", subfolders=[inbox_folder, sent_folder])
        ns = FakeNamespace([FakeStore("user@example.com", root_folder)], inbox_folder, sent_folder)

        with tempfile.TemporaryDirectory() as tmp:
            target_dir = Path(tmp) / "emails"
            with mock.patch.object(sync, "get_outlook_namespace", return_value=ns):
                stats = sync.run_sync(
                    target_dir=target_dir,
                    inbox_subfolder="MojProjekt",
                    since_date=datetime.date(2026, 1, 1),
                    dry_run=False,
                )
                self.assertEqual(stats["synced"], 1)


if __name__ == "__main__":
    unittest.main()
