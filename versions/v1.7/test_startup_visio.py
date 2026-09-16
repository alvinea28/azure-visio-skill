"""Noninteractive first-run and repeated-run regression tests."""

import base64
import builtins
from contextlib import contextmanager, redirect_stderr, redirect_stdout
import hashlib
import io
import json
import os
from pathlib import Path
import shutil
import unittest
from unittest import mock
import uuid
import zipfile

from PIL import Image

import portable_visio as p
from test_offline_visio import restricted_runtime
from test_portable_visio import sample_model, SVG


class NoInput(io.StringIO):
    def read(self, *args):
        raise AssertionError("The renderer attempted to wait for input")

    def readline(self, *args):
        raise AssertionError("The renderer attempted to wait for input")


class StartupTests(unittest.TestCase):
    def setUp(self):
        self.root = Path(__file__).parent / (".startup-tests-" + uuid.uuid4().hex)
        self.resources = self.root / "readonly-resources"
        self.task = self.root / "task"
        self.resources.mkdir(parents=True)
        self.task.mkdir()
        self.script = self.resources / "portable_visio.py"
        self.script.write_text("# Synthetic path anchor, not executed\n", encoding="utf-8")
        self.model = self.task / "model.json"
        self.model.write_text(json.dumps(sample_model()), encoding="utf-8")
        image = io.BytesIO()
        Image.new("RGBA", (384, 192), "blue").save(image, format="PNG")
        data = image.getvalue()
        record = {"id": "azure-test", "path": "png/azure-test.png", "usable": True,
                  "name": "Synthetic Vault", "sha256": hashlib.sha256(data).hexdigest(),
                  "originalPath": "azure/test.svg", "originalSha256": hashlib.sha256(SVG).hexdigest(),
                  "sourceAspect": 2, "pixelWidth": 384, "pixelHeight": 192}
        self.archive = self.resources / "icons.zip"
        with zipfile.ZipFile(self.archive, "w") as archive:
            archive.writestr("catalog.json", json.dumps([record]))
            archive.writestr(record["path"], data)
        self.manifest = self.resources / "offline-assets.json"
        self.manifest.write_text(json.dumps({"schemaVersion": 1, "kind": "azure-visio-offline-png",
                                           "archives": [{"file": "icons.zip",
                                                         "sha256": hashlib.sha256(self.archive.read_bytes()).hexdigest()}]}),
                                 encoding="utf-8")

    def tearDown(self):
        for path in self.resources.iterdir():
            if path.is_file():
                path.chmod(0o600)
        shutil.rmtree(self.root)

    @contextmanager
    def unattended(self):
        with restricted_runtime(), mock.patch.object(p, "__file__", str(self.script)), \
                mock.patch.object(builtins, "input", side_effect=AssertionError("Interactive prompt")), \
                mock.patch.object(p.sys, "stdin", NoInput()):
            yield

    def command(self, args, expected=0):
        out, err = io.StringIO(), io.StringIO()
        with redirect_stdout(out), redirect_stderr(err):
            code = p.main(args)
        self.assertEqual(code, expected, err.getvalue())
        return json.loads(out.getvalue() if expected == 0 else err.getvalue())

    def flat_assets(self):
        with zipfile.ZipFile(self.archive) as archive:
            records = json.loads(archive.read("catalog.json"))
            raw = archive.read(records[0]["path"])
        records[0]["dataFile"] = "Icon-Data-1.json"
        catalog = self.resources / "catalog.json"
        data = self.resources / "Icon-Data-1.json"
        catalog.write_text(json.dumps(records), encoding="utf-8")
        data.write_text(json.dumps({"schemaVersion": 1, "kind": "azure-visio-png-data",
                                    "images": {"azure-test": base64.b64encode(raw).decode("ascii")}}),
                        encoding="utf-8")
        manifest = {"schemaVersion": 1, "kind": "azure-visio-offline-png-json",
                    "catalog": {"file": catalog.name, "sha256": hashlib.sha256(catalog.read_bytes()).hexdigest()},
                    "dataFiles": [{"file": data.name, "sha256": hashlib.sha256(data.read_bytes()).hexdigest()}]}
        self.manifest.write_text(json.dumps(manifest), encoding="utf-8")
        self.archive.unlink()
        return catalog, data, manifest

    def refresh_flat_hashes(self, manifest):
        for entry in [manifest["catalog"], *manifest["dataFiles"]]:
            entry["sha256"] = hashlib.sha256((self.resources / entry["file"]).read_bytes()).hexdigest()
        self.manifest.write_text(json.dumps(manifest), encoding="utf-8")

    def test_flat_json_companions_start_and_render_without_runtime_archives(self):
        self.flat_assets()
        self.assertFalse(list(self.resources.glob("*.zip")))
        with self.unattended():
            self.assertEqual(self.command(["startup"])["icons"], 1)
            result = self.command(["render", "--model", str(self.model)])
        self.assertEqual(result["pageCount"], 3)

    def test_flat_companion_hash_tampering_is_rejected(self):
        _, data, _ = self.flat_assets()
        data.write_text("{}", encoding="utf-8")
        with self.unattended():
            result = self.command(["startup"], 2)
        self.assertIn("SHA256 mismatch", result["error"])

    def test_flat_payload_rejects_invalid_base64_uncatalogued_and_missing_images(self):
        _, data, manifest = self.flat_assets()
        for images, pattern in (({"azure-test": "!"}, "base64"),
                                ({"not-in-catalog": "YWJj"}, "uncatalogued"),
                                ({}, "missing a usable")):
            data.write_text(json.dumps({"schemaVersion": 1, "kind": "azure-visio-png-data",
                                        "images": images}), encoding="utf-8")
            self.refresh_flat_hashes(manifest)
            with self.unattended():
                result = self.command(["startup"], 2)
            self.assertIn(pattern, result["error"])

    def test_flat_catalog_data_ownership_and_png_hash_are_enforced(self):
        catalog, _, manifest = self.flat_assets()
        original = json.loads(catalog.read_text(encoding="utf-8"))
        for key, value, pattern in (("dataFile", "other.json", "catalog companion"),
                                    ("sha256", "0" * 64, "Icon SHA256 mismatch")):
            records = [{**original[0], key: value}]
            catalog.write_text(json.dumps(records), encoding="utf-8")
            self.refresh_flat_hashes(manifest)
            with self.unattended():
                result = self.command(["startup"], 2)
            self.assertIn(pattern, result["error"])

    def test_flat_manifest_accepts_only_safe_sibling_json_paths(self):
        _, _, manifest = self.flat_assets()
        for name in ("../data.json", "icons.zip", "code.py"):
            manifest["dataFiles"][0]["file"] = name
            self.manifest.write_text(json.dumps(manifest), encoding="utf-8")
            with self.unattended():
                result = self.command(["startup"], 2)
            self.assertEqual(result["status"], "blocked")

    def test_manifest_root_type_is_validated(self):
        self.manifest.write_text("[]", encoding="utf-8")
        with self.unattended():
            result = self.command(["startup"], 2)
        self.assertIn("must be an object", result["error"])

    def test_cold_start_from_readonly_resources_outside_cwd_needs_no_input(self):
        originals = {path.name: path.read_bytes() for path in self.resources.iterdir()}
        for path in self.resources.iterdir():
            path.chmod(0o444)
        old = Path.cwd()
        os.chdir(self.task)
        try:
            with self.unattended():
                result = self.command(["startup"])
                entries = self.command(["catalog", "--search", "vault"])
        finally:
            os.chdir(old)
        self.assertEqual(result["status"], "ready")
        self.assertFalse(result["userInputRequired"])
        self.assertFalse(result["networkRequired"])
        self.assertEqual(entries[0]["id"], "azure-test")
        self.assertEqual(originals, {path.name: path.read_bytes() for path in self.resources.iterdir()})

    def test_two_runs_allocate_different_outputs_without_overwrite_questions(self):
        with self.unattended():
            validation = self.command(["validate", "--model", str(self.model)])
            first = self.command(["render", "--model", str(self.model), "--bundle"])
            second = self.command(["render", "--model", str(self.model), "--bundle"])
        self.assertTrue(validation["valid"])
        self.assertNotEqual(first["vsdx"], second["vsdx"])
        for result in (first, second):
            self.assertEqual(result["status"], "complete")
            self.assertEqual(result["pageCount"], 3)
            self.assertEqual(Path(result["vsdx"]).parent.parent, self.task)
            self.assertTrue(Path(result["pdf"]).read_bytes().startswith(b"%PDF-"))
            self.assertEqual(len(p.inspect_package(result["vsdx"])["pages"]), 3)

    def test_explicit_output_root_is_used_without_an_interactive_path_prompt(self):
        output_root = self.root / "results"
        output_root.mkdir()
        with self.unattended():
            result = self.command(["render", "--model", str(self.model),
                                   "--output-root", str(output_root), "--name", "pilot"])
        self.assertEqual(Path(result["vsdx"]).parent.parent, output_root)
        self.assertTrue(Path(result["vsdx"]).name.startswith("pilot"))

    def test_output_in_installed_resources_is_not_automatically_created(self):
        with self.unattended():
            result = self.command(["render", "--model", str(self.model),
                                   "--output-root", str(self.resources)], 2)
        self.assertIn("installed skill resources", result["error"])
        self.assertFalse(result["userInputPending"])

    def test_missing_manifest_fails_closed_without_manual_download_fallback(self):
        self.manifest.unlink()
        with self.unattended():
            for args in (["startup"], ["catalog"], ["render", "--model", str(self.model)]):
                result = self.command(args, 2)
                self.assertEqual(result["status"], "blocked")
                self.assertFalse(result["userInputPending"])
                self.assertIn("complete package", result["error"])
        self.assertEqual(list(self.task.iterdir()), [self.model])

    def test_missing_or_changed_asset_is_an_immediate_blocker(self):
        for contents in (b"changed", None):
            if contents is None:
                self.archive.unlink()
            else:
                self.archive.write_bytes(contents)
            with self.unattended():
                result = self.command(["startup"], 2)
            self.assertEqual(result["status"], "blocked")
            self.assertFalse(result["userInputPending"])

    def test_missing_base_library_never_prompts_for_package_installation(self):
        with self.unattended(), mock.patch.object(
                p, "capabilities", return_value={"dependencies": {"reportlab": False, "Pillow": True}}):
            result = self.command(["startup"], 2)
        self.assertIn("Unavailable rendering dependencies: reportlab", result["error"])
        self.assertFalse(result["userInputPending"])

    def test_present_but_broken_base_import_reports_blocker(self):
        original_import = builtins.__import__

        def broken_import(name, *args, **kwargs):
            if name.startswith("reportlab.pdfbase"):
                raise ImportError("Synthetic broken preinstalled ReportLab")
            return original_import(name, *args, **kwargs)

        with self.unattended(), mock.patch("builtins.__import__", side_effect=broken_import):
            result = self.command(["startup"], 2)
        self.assertIn("Synthetic broken preinstalled", result["error"])
        self.assertFalse(result["userInputPending"])

    def test_routing_budget_fails_instead_of_waiting_and_context_is_reset(self):
        with self.unattended(), mock.patch.object(p, "ROUTING_BUDGET_SECONDS", -1):
            result = self.command(["render", "--model", str(self.model)], 2)
        self.assertIn("layout budget", result["error"])
        self.assertIsNone(p.ROUTING_DEADLINE.get())
        self.assertEqual(list(self.task.iterdir()), [self.model])
        with self.unattended():
            self.assertEqual(self.command(["render", "--model", str(self.model)])["status"], "complete")

    def test_existing_explicit_output_remains_protected(self):
        output = self.task / "existing"
        output.mkdir()
        sentinel = output / "keep.txt"
        sentinel.write_text("Keep user content", encoding="utf-8")
        with self.unattended():
            result = self.command(["render", "--model", str(self.model),
                                   "--output-directory", str(output)], 2)
        self.assertIn("overwrites are forbidden", result["error"])
        self.assertEqual(sentinel.read_text(encoding="utf-8"), "Keep user content")


if __name__ == "__main__":
    unittest.main()
