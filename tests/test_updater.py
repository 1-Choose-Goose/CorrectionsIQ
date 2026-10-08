from __future__ import annotations

import hashlib
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest import mock

from update_helper import (
    install,
    prepare_staged_install,
    resource_path,
    safe_extract_archive,
    source_root,
    validate_install_paths,
)
from updater import UpdateInfo, UpdateState, validate_release_asset_url, verify_download


class UpdateStateTests(unittest.TestCase):
    def test_download_progress_survives_view_recreation(self) -> None:
        update = UpdateInfo(
            version="1.3",
            page_url="https://github.com/1-Choose-Goose/CorrectionsIQ/releases/tag/v1.3",
            asset_url="https://github.com/1-Choose-Goose/CorrectionsIQ/releases/download/v1.3/CorrectionsIQ-v1.3.zip",
        )
        state = UpdateState()

        state.mark_available(update)
        state.start_download()
        state.update_download(received=25, total=100)

        self.assertEqual(state.phase, "downloading")
        self.assertEqual(state.progress, 0.25)
        self.assertEqual(state.status_text, "Скачивание обновления... 25%")
        self.assertTrue(state.progress_visible)

    def test_failed_download_can_be_retried(self) -> None:
        update = UpdateInfo(version="1.3", page_url="", asset_url="https://example.invalid/update.zip")
        state = UpdateState()

        state.mark_available(update)
        state.start_download()
        state.mark_error("Сеть недоступна")

        self.assertFalse(state.busy)
        self.assertEqual(state.button_text, "Повторить")
        self.assertIn("Сеть недоступна", state.status_text)


class UpdateDownloadSafetyTests(unittest.TestCase):
    def test_release_asset_url_is_limited_to_project_releases(self) -> None:
        valid = (
            "https://github.com/1-Choose-Goose/CorrectionsIQ/releases/download/"
            "v1.3/CorrectionsIQ-v1.3.zip"
        )
        validate_release_asset_url(valid)

        with self.assertRaisesRegex(RuntimeError, "недопустимый адрес"):
            validate_release_asset_url("https://example.com/CorrectionsIQ-v1.3.zip")

    def test_download_digest_detects_modified_archive(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            archive_path = Path(temp_dir) / "update.zip"
            archive_path.write_bytes(b"expected update")
            digest = hashlib.sha256(b"expected update").hexdigest()
            update = UpdateInfo(
                version="1.3",
                page_url="",
                asset_url=(
                    "https://github.com/1-Choose-Goose/CorrectionsIQ/releases/download/"
                    "v1.3/CorrectionsIQ-v1.3.zip"
                ),
                asset_digest=f"sha256:{digest}",
                asset_size=len(b"expected update"),
            )

            verify_download(archive_path, update)
            archive_path.write_bytes(b"modified update")

            with self.assertRaisesRegex(RuntimeError, "контрольная сумма"):
                verify_download(archive_path, update)


class UpdateArchiveSafetyTests(unittest.TestCase):
    def test_installer_icon_resource_exists(self) -> None:
        self.assertTrue(resource_path("assets/app_icon.ico").is_file())

    def test_safe_extract_rejects_parent_directory_escape(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            archive_path = root / "update.zip"
            extract_dir = root / "extract"
            with zipfile.ZipFile(archive_path, "w") as archive:
                archive.writestr("../outside.txt", "unsafe")

            with self.assertRaisesRegex(RuntimeError, "недопустимый путь"):
                safe_extract_archive(archive_path, extract_dir)

            self.assertFalse((root / "outside.txt").exists())

    def test_source_root_requires_application_executable(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            extracted = Path(temp_dir)

            with self.assertRaisesRegex(RuntimeError, "CorrectionsIQ.exe"):
                source_root(extracted)

    def test_staged_install_preserves_user_files_and_replaces_internal_folder(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            target = root / "CorrectionsIQ"
            source = root / "source"
            staging = root / "staging"
            (target / "_internal").mkdir(parents=True)
            (source / "_internal").mkdir(parents=True)
            (target / "CorrectionsIQ.exe").write_text("old", encoding="utf-8")
            (target / "user-file.txt").write_text("keep", encoding="utf-8")
            (target / "_internal" / "stale.dll").write_text("old", encoding="utf-8")
            (source / "CorrectionsIQ.exe").write_text("new", encoding="utf-8")
            (source / "_internal" / "current.dll").write_text("new", encoding="utf-8")

            prepare_staged_install(source, target, staging)

            self.assertEqual((staging / "user-file.txt").read_text(encoding="utf-8"), "keep")
            self.assertFalse((staging / "_internal" / "stale.dll").exists())
            self.assertEqual(
                (staging / "_internal" / "current.dll").read_text(encoding="utf-8"),
                "new",
            )
            self.assertEqual((staging / "CorrectionsIQ.exe").read_text(encoding="utf-8"), "new")

    def test_validate_install_paths_rejects_app_outside_target(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            target = root / "CorrectionsIQ"
            target.mkdir()
            archive_path = root / "update.zip"
            archive_path.write_bytes(b"zip")
            outside_app = root / "Other" / "CorrectionsIQ.exe"

            with self.assertRaisesRegex(RuntimeError, "папке программы"):
                validate_install_paths(archive_path, target, outside_app)

    def test_install_rolls_back_when_updated_application_cannot_start(self) -> None:
        class FakeWindow:
            def update(self, text: str, percent: int) -> None:
                del text, percent
                return

        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            target = root / "CorrectionsIQ"
            (target / "_internal").mkdir(parents=True)
            app = target / "CorrectionsIQ.exe"
            app.write_text("old application", encoding="utf-8")
            (target / "user-file.txt").write_text("keep", encoding="utf-8")
            archive_path = root / "update.zip"
            with zipfile.ZipFile(archive_path, "w") as archive:
                archive.writestr("CorrectionsIQ/CorrectionsIQ.exe", "new application")
                archive.writestr("CorrectionsIQ/_internal/current.dll", "new library")

            with (
                mock.patch("update_helper.subprocess.Popen", side_effect=OSError("blocked")),
                self.assertRaisesRegex(OSError, "blocked"),
            ):
                install(archive_path, target, app, FakeWindow())

            self.assertEqual(app.read_text(encoding="utf-8"), "old application")
            self.assertEqual((target / "user-file.txt").read_text(encoding="utf-8"), "keep")


if __name__ == "__main__":
    unittest.main()
