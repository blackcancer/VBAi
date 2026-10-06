"""Verify the whole help document contract without requiring a Windows compiler."""
import hashlib
import importlib.util
import json
import tempfile
import unittest
from unittest.mock import patch
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
spec = importlib.util.spec_from_file_location("build_help", ROOT / "tools/docs/build_help.py")
help_builder = importlib.util.module_from_spec(spec)
spec.loader.exec_module(help_builder)


class HelpBuilderTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.source = Path(self.temp.name) / "source"
        self.output = Path(self.temp.name) / "output"
        shots = self.source / "screenshots"
        shots.mkdir(parents=True)
        data = b"\x89PNG\r\n\x1a\nfixture"
        (shots / "view.png").write_bytes(data)
        self.capture = {"id": "View", "file": "view.png", "caption": "Édition simulée", "synthetic": True,
                        "sha256": hashlib.sha256(data).hexdigest()}
        self.manual = {"interfaces": ["View"], "topics": [{"id": "start", "title": "Accueil français",
            "intro": "Données fictives", "interfaces": ["View"], "screenshots": ["View"],
            "keywords": ["Édition"], "sections": [], "related": []}]}
        (self.source / "manual.json").write_text(json.dumps(self.manual), encoding="utf-8")
        (shots / "manifest.json").write_text(json.dumps({"captures": [self.capture]}), encoding="utf-8")

    def test_html_toc_index_and_search_configuration_are_generated_together(self):
        page = help_builder.build(self.source, self.output, None)
        self.assertIn("Accueil français", page.read_text(encoding="utf-8"))
        self.assertIn("lang='fr'", page.read_text(encoding="utf-8"))
        self.assertIn("Édition", (self.output / "manual.hhk").read_text(encoding="cp1252"))
        project = (self.output / "manual.hhp").read_text(encoding="cp1252")
        self.assertIn("Full-text search=Yes", project)
        self.assertIn("Language=0x40c", project)
        self.assertEqual((self.source / "screenshots/view.png").read_bytes(), (self.output / "screenshots/view.png").read_bytes())

    def test_missing_interface_mapping_is_rejected(self):
        self.manual["interfaces"].append("MissingView")
        with self.assertRaises(ValueError): help_builder.validate(self.manual, self.source, {"View": self.capture})

    def test_duplicate_topic_or_interface_is_rejected(self):
        self.manual["topics"].append(self.manual["topics"][0])
        with self.assertRaises(ValueError): help_builder.validate(self.manual, self.source, {"View": self.capture})

    def test_unknown_topic_link_is_rejected(self):
        self.manual["topics"][0]["related"] = ["Missing"]
        with self.assertRaises(ValueError): help_builder.validate(self.manual, self.source, {"View": self.capture})

    def test_capture_hash_and_synthetic_provenance_are_checked(self):
        for mutation in ({"sha256": "incorrect"}, {"synthetic": False}):
            capture = dict(self.capture, **mutation)
            with self.assertRaises(ValueError): help_builder.validate(self.manual, self.source, {"View": capture})

    def test_user_text_is_escaped_and_does_not_create_markup(self):
        topic = dict(self.manual["topics"][0], intro="<script>unsafe</script>")
        html = help_builder.page(topic, {"start": "Accueil"}, {"View": self.capture})
        self.assertIn("&lt;script&gt;unsafe&lt;/script&gt;", html)
        self.assertNotIn("<script>", html)

    def test_capture_and_topic_paths_cannot_escape_the_manual(self):
        for name in ("../view.png", "C:/view.png", "nested/view.png"):
            with self.assertRaises(ValueError):
                help_builder.validate(self.manual, self.source, {"View": dict(self.capture, file=name)})
        self.manual["topics"][0]["id"] = "../start"
        with self.assertRaises(ValueError): help_builder.validate(self.manual, self.source, {"View": self.capture})

    def test_old_generated_files_are_not_included_in_a_new_archive(self):
        self.output.mkdir()
        (self.output / "obsolete.html").write_text("old", encoding="utf-8")
        help_builder.build(self.source, self.output, None)
        self.assertNotIn("obsolete.html", (self.output / "manual.hhp").read_text(encoding="cp1252"))

    def test_old_archive_cannot_hide_a_failed_new_compilation(self):
        self.output.mkdir()
        old = self.output / "VBAi.fr-FR.chm"
        old.write_bytes(b"ITSFold archive")
        with patch.object(help_builder.subprocess, "run", return_value=help_builder.subprocess.CompletedProcess([], 1, b"failed")):
            with self.assertRaises(RuntimeError): help_builder.build(self.source, self.output, Path("compiler.exe"))
        self.assertFalse(old.exists())

    def test_compiler_error_rejects_a_partial_archive_even_with_success_exit_code(self):
        def compile_partial(*args, **kwargs):
            (self.output / "VBAi.fr-FR.chm").write_bytes(b"ITSFpartial archive")
            return help_builder.subprocess.CompletedProcess([], 1,
                b"HHC6003: Error: The file Itircl.dll has not been registered correctly.")
        with patch.object(help_builder.subprocess, "run", side_effect=compile_partial):
            with self.assertRaises(RuntimeError): help_builder.build(self.source, self.output, Path("compiler.exe"))
        self.assertFalse((self.output / "VBAi.fr-FR.chm").exists())


if __name__ == "__main__": unittest.main()
