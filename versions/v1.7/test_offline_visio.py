"""Offline PNG regression cases; fixtures are synthetic, not fallback product icons."""

import builtins
from contextlib import ExitStack, contextmanager, redirect_stderr, redirect_stdout
import copy
import hashlib
import io
import json
from pathlib import Path
import re
import shutil
import stat
import unittest
from unittest import mock
import uuid
import zipfile

from PIL import Image

import portable_visio as p
from test_portable_visio import sample_model, SVG


@contextmanager
def restricted_runtime():
    find_spec = p.importlib.util.find_spec
    original_import = builtins.__import__

    def find(name, *args, **kwargs):
        return None if name == "resvg_py" else find_spec(name, *args, **kwargs)

    def import_module(name, *args, **kwargs):
        if name.split(".")[0] in ("resvg_py", "pip"):
            raise AssertionError("Offline rendering attempted an unavailable/forbidden import: " + name)
        return original_import(name, *args, **kwargs)

    with ExitStack() as stack:
        for target in ("urllib.request.OpenerDirector.open", "urllib.request.urlopen",
                       "socket.create_connection", "subprocess.run", "subprocess.Popen", "os.system", "os.popen"):
            stack.enter_context(mock.patch(target, side_effect=AssertionError("Offline operation used " + target)))
        stack.enter_context(mock.patch.object(p.importlib.util, "find_spec", side_effect=find))
        stack.enter_context(mock.patch("builtins.__import__", side_effect=import_module))
        yield


class OfflineTests(unittest.TestCase):
    def setUp(self):
        self.root = Path(__file__).parent / (".offline-tests-" + uuid.uuid4().hex)
        self.root.mkdir()
        self.png = self.image_bytes()
        self.entry = {
            "id": "azure-test", "path": "png/azure-test.png", "collection": "azure",
            "name": "Synthetic Recovery-Services-Vault Test", "usable": True, "issue": "",
            "sha256": hashlib.sha256(self.png).hexdigest(),
            "originalSha256": hashlib.sha256(SVG).hexdigest(),
            "originalPath": "azure/test.svg", "sourceAspect": 2,
            "pixelWidth": 512, "pixelHeight": 256,
        }
        self.archive = self.root / "icons.zip"
        self.write_archive()

    def tearDown(self):
        for file in self.root.rglob("*"):
            if file.is_file():
                file.chmod(0o600)
        shutil.rmtree(self.root)

    @staticmethod
    def image_bytes(size=(512, 256), fill=(0, 128, 255, 255)):
        stream = io.BytesIO()
        Image.new("RGBA", size, fill).save(stream, format="PNG")
        return stream.getvalue()

    def write_archive(self, *, records=None, png=None, extra=(), path=None, include_png=True):
        records = [self.entry] if records is None else records
        with zipfile.ZipFile(path or self.archive, "w", zipfile.ZIP_DEFLATED) as archive:
            archive.writestr("catalog.json", json.dumps(records))
            archive.writestr("NOTICE.txt", "Synthetic test artwork only.")
            if include_png:
                archive.writestr("png/azure-test.png", self.png if png is None else png)
            for name, data in extra:
                archive.writestr(name, data)

    def catalog(self):
        return p.IconCatalog(archives=[self.archive])

    def test_capabilities_distinguish_png_from_svg(self):
        with restricted_runtime():
            report = p.capabilities()
            self.assertTrue(report["renderAvailable"])
            self.assertEqual(report["renderModes"], {"png": True, "svg": False})
            self.assertFalse(report["dependencies"]["resvg-py"])
            self.assertIsNone(p.dependencies()[3])
            with self.assertRaisesRegex(p.PortableError, "resvg-py"):
                p.dependencies(svg=True)

    def test_complete_offline_pack_from_read_only_archive(self):
        self.archive.chmod(0o444)
        original = self.archive.read_bytes()
        with restricted_runtime():
            result = p.render(sample_model(), None, self.root / "output",
                              bundle=True, icon_archives=[self.archive])
        self.assertEqual(self.archive.read_bytes(), original)
        self.assertEqual(len(result["inspection"]["pages"]), 3)
        self.assertIn("Packaged PNG pixels unchanged", result["glyphFormat"])
        pdf = Path(result["pdf"]).read_bytes()
        self.assertTrue(pdf.startswith(b"%PDF-"))
        self.assertEqual(len(re.findall(rb"/Type\s*/Page\b", pdf)), 3)
        with zipfile.ZipFile(result["vsdx"]) as drawing:
            artwork = [drawing.read(name) for name in drawing.namelist()
                       if name.startswith("visio/media/")]
            self.assertEqual(artwork, [self.png])
        with zipfile.ZipFile(result["bundle"]) as bundle:
            self.assertEqual(set(bundle.namelist()), {"architecture.vsdx", "architecture.pdf"})
        self.assertEqual(sorted(path.name for path in (self.root / "output").iterdir()),
                         ["architecture.pdf", "architecture.vsdx", "architecture.zip"])

    def test_png_directory_and_archive_have_identical_artwork(self):
        directory = self.root / "unpacked"
        (directory / "png").mkdir(parents=True)
        (directory / "catalog.json").write_text(json.dumps([self.entry]), encoding="utf-8")
        (directory / self.entry["path"]).write_bytes(self.png)
        with restricted_runtime():
            self.assertEqual(p.IconCatalog(directory).resolve("azure-test"),
                             self.catalog().resolve("azure-test"))

    def test_split_data_archives_need_no_extraction(self):
        self.write_archive(include_png=False)
        shard = self.root / "part2.zip"
        with zipfile.ZipFile(shard, "w") as archive:
            archive.writestr(self.entry["path"], self.png)
        with restricted_runtime():
            art = p.IconCatalog(archives=[self.archive, shard]).resolve("azure-test")
        self.assertEqual(art["png"], self.png)
        self.assertFalse((self.root / "png").exists())

    def test_catalog_search_uses_actual_ids_without_rendering(self):
        output = io.StringIO()
        with restricted_runtime(), redirect_stdout(output):
            code = p.main(["catalog", "--icon-archive", str(self.archive),
                           "--search", "recovery services"])
        self.assertEqual(code, 0)
        self.assertEqual(json.loads(output.getvalue())[0]["id"], "azure-test")
        self.assertEqual(self.catalog().search("unavailable-service"), [])

    def test_product_search_does_not_confuse_entra_and_central(self):
        catalog = self.catalog()
        catalog.entries = {
            "azure-identity": {**self.entry, "name": "Microsoft-Entra-ID"},
            "azure-iot": {**self.entry, "name": "IoT-Central"},
        }
        self.assertEqual([entry["id"] for entry in catalog.search("entra")], ["azure-identity"])

    def test_offline_render_and_validate_cli(self):
        model_path = self.root / "model.json"
        model_path.write_text(json.dumps(sample_model()), encoding="utf-8")
        for command in ("validate", "render"):
            args = [command, "--model", str(model_path), "--icon-archive", str(self.archive)]
            if command == "render":
                args += ["--output-directory", str(self.root / "cli-output")]
            output = io.StringIO()
            with restricted_runtime(), redirect_stdout(output):
                self.assertEqual(p.main(args), 0)
            result = json.loads(output.getvalue())
            self.assertTrue(result["valid"] if command == "validate" else Path(result["vsdx"]).is_file())

    def test_bundled_manifest_drives_catalog_validation_and_rendering(self):
        manifest_path = self.root / "offline-assets.json"
        manifest = {"schemaVersion": 1, "kind": "azure-visio-offline-png",
                    "archives": [{"file": self.archive.name,
                                  "sha256": hashlib.sha256(self.archive.read_bytes()).hexdigest()}]}
        manifest_path.write_text(json.dumps(manifest), encoding="utf-8")
        model_path = self.root / "model.json"
        model_path.write_text(json.dumps(sample_model()), encoding="utf-8")
        for command in ("catalog", "validate", "render"):
            args = [command, "--offline-assets", str(manifest_path)]
            if command != "catalog":
                args += ["--model", str(model_path)]
            if command == "render":
                args += ["--output-directory", str(self.root / "manifest-output")]
            with restricted_runtime(), redirect_stdout(io.StringIO()):
                self.assertEqual(p.main(args), 0)
        self.assertTrue((self.root / "manifest-output" / "architecture.vsdx").is_file())

    def test_manifest_rejects_tampering_traversal_and_duplicate_archives(self):
        manifest_path = self.root / "offline-assets.json"
        entry = {"file": self.archive.name,
                 "sha256": hashlib.sha256(self.archive.read_bytes()).hexdigest()}
        for archives, pattern in (
            ([{**entry, "sha256": "0" * 64}], "SHA256 mismatch"),
            ([{**entry, "file": "../icons.zip"}], "Unsafe"),
            ([{**entry, "file": "module.whl"}], "companion ZIP"),
            ([entry, entry], "Duplicate"),
            ([], "no archives"),
        ):
            manifest_path.write_text(json.dumps({"schemaVersion": 1, "kind": "azure-visio-offline-png",
                                                "archives": archives}), encoding="utf-8")
            with self.assertRaisesRegex(p.PortableError, pattern):
                p.IconCatalog(offline_assets=manifest_path)

    def test_svg_still_fails_explicitly_without_its_optional_renderer(self):
        directory = self.root / "svg"
        directory.mkdir()
        (directory / "test.svg").write_bytes(SVG)
        entry = {"id": "azure-test", "path": "test.svg", "usable": True,
                 "sha256": hashlib.sha256(SVG).hexdigest()}
        (directory / "catalog.json").write_text(json.dumps([entry]), encoding="utf-8")
        with restricted_runtime(), self.assertRaisesRegex(p.PortableError, "resvg-py"):
            p.render(sample_model(), directory, self.root / "no-output")
        self.assertFalse((self.root / "no-output").exists())

    def test_local_official_svg_archive_has_supported_import_command(self):
        source = self.root / "official-fixture.zip"
        with zipfile.ZipFile(source, "w") as archive:
            archive.writestr("Icons/Test.svg", SVG)
        destination = self.root / "imported-svg"
        with restricted_runtime(), redirect_stdout(io.StringIO()), mock.patch.object(
                p, "OFFICIAL_ARCHIVE_SHA256", hashlib.sha256(source.read_bytes()).hexdigest()):
            self.assertEqual(p.main(["import-icons", "--archive", str(source),
                                     "--output-directory", str(destination), "--accept-icon-terms"]), 0)
        self.assertEqual(len(p.load_json(destination / "catalog.json")), 1)

    def test_import_requires_explicit_terms_acceptance(self):
        with restricted_runtime(), redirect_stderr(io.StringIO()):
            self.assertEqual(p.main(["import-icons", "--archive", str(self.archive),
                                     "--output-directory", str(self.root / "not-created")]), 2)
        self.assertFalse((self.root / "not-created").exists())

    def test_import_does_not_mislabel_an_unverified_archive_as_official(self):
        with restricted_runtime(), redirect_stderr(io.StringIO()):
            self.assertEqual(p.main(["import-icons", "--archive", str(self.archive),
                                     "--output-directory", str(self.root / "not-created"),
                                     "--accept-icon-terms"]), 2)
        self.assertFalse((self.root / "not-created").exists())

    def test_directory_and_archive_are_mutually_exclusive(self):
        with self.assertRaisesRegex(p.PortableError, "either"):
            p.IconCatalog(self.root, archives=[self.archive])
        with redirect_stderr(io.StringIO()), self.assertRaises(SystemExit):
            p.main(["catalog", "--icon-directory", str(self.root),
                    "--icon-archive", str(self.archive)])

    def test_archive_rejects_code_wheels_native_modules_and_traversal(self):
        for name in ("plugin.py", "package.whl", "module.so", "module.dll",
                     "../escape.png", "png/../escape.png", "C:/escape.png", "png/NUL.png"):
            with self.subTest(name=name):
                self.write_archive(extra=[(name, b"not executable")])
                with self.assertRaises(p.PortableError):
                    self.catalog()

    def test_archive_rejects_case_duplicates_and_symlinks(self):
        self.write_archive(extra=[("png/AZURE-TEST.png", self.png)])
        with self.assertRaisesRegex(p.PortableError, "Duplicate"):
            self.catalog()
        self.write_archive()
        with zipfile.ZipFile(self.archive, "a") as archive:
            info = zipfile.ZipInfo("png/link.png")
            info.create_system = 3
            info.external_attr = (stat.S_IFLNK | 0o777) << 16
            archive.writestr(info, b"target.png")
        with self.assertRaisesRegex(p.PortableError, "Symlink"):
            self.catalog()

    def test_archive_requires_catalog_and_every_selected_resource(self):
        with zipfile.ZipFile(self.archive, "w") as archive:
            archive.writestr("NOTICE.txt", "No catalog.")
        with self.assertRaisesRegex(p.PortableError, "catalog.json is missing"):
            self.catalog()
        self.write_archive(include_png=False)
        with self.assertRaisesRegex(p.PortableError, "Missing offline PNG"):
            self.catalog()

    def test_unusable_source_stays_searchable_but_cannot_be_selected(self):
        entry = copy.deepcopy(self.entry)
        entry.update(usable=False, issue="Unsupported SVG text", path="azure/source.svg")
        self.write_archive(records=[entry], include_png=False)
        catalog = self.catalog()
        self.assertFalse(catalog.search()[0]["usable"])
        self.assertEqual(catalog.search()[0]["issue"], "Unsupported SVG text")
        with self.assertRaisesRegex(p.PortableError, "unusable"):
            catalog.resolve("azure-test")

    def test_archive_rejects_uncatalogued_payloads_and_expansion_bombs(self):
        self.write_archive(extra=[("png/extra.png", self.png)])
        with self.assertRaisesRegex(p.PortableError, "Uncatalogued"):
            self.catalog()
        self.write_archive(png=b"\0" * (2 * 1024 * 1024))
        with self.assertRaisesRegex(p.PortableError, "expansion"):
            self.catalog()
        self.write_archive()
        with mock.patch.object(p, "MAX_EXPANDED", 100):
            with self.assertRaisesRegex(p.PortableError, "Expanded"):
                self.catalog()

    def test_archive_json_and_duplicate_catalog_ids_fail_closed(self):
        self.write_archive(records=[self.entry, self.entry])
        with self.assertRaisesRegex(p.PortableError, "Duplicate catalog id"):
            self.catalog()
        with zipfile.ZipFile(self.archive, "w") as archive:
            archive.writestr("catalog.json", '[{"id":"a","ID":"b"}]')
        with self.assertRaisesRegex(p.PortableError, "Duplicate JSON key"):
            self.catalog()

    def test_png_catalog_hash_provenance_dimensions_and_aspect_are_checked(self):
        for key, value, pattern in (
            ("sha256", "0" * 64, "SHA256 mismatch"),
            ("originalSha256", "invalid", "provenance"),
            ("originalPath", "source.py", "provenance"),
            ("pixelWidth", 513, "pixelWidth"),
            ("pixelHeight", True, "pixelHeight"),
            ("sourceAspect", 4, "aspect"),
            ("sourceAspect", float("nan"), "Non-finite"),
        ):
            with self.subTest(key=key, value=value):
                entry = copy.deepcopy(self.entry)
                entry[key] = value
                self.write_archive(records=[entry])
                with self.assertRaisesRegex(p.PortableError, pattern):
                    self.catalog().resolve("azure-test")

    def test_malformed_empty_and_oversized_pngs_fail_closed(self):
        for data, size, pattern in (
            (b"not a PNG", (512, 256), "Invalid bounded PNG"),
            (self.png[:40], (512, 256), "Invalid PNG artwork"),
            (self.image_bytes(fill=(0, 0, 0, 0)), (512, 256), "empty"),
            (self.image_bytes(size=(2401, 1)), (2401, 1), "dimensions"),
        ):
            with self.subTest(pattern=pattern):
                entry = copy.deepcopy(self.entry)
                entry.update(sha256=hashlib.sha256(data).hexdigest(),
                             pixelWidth=size[0], pixelHeight=size[1])
                self.write_archive(records=[entry], png=data)
                with self.assertRaisesRegex(p.PortableError, pattern):
                    self.catalog().resolve("azure-test")

    def test_animated_png_is_not_accepted_as_static_artwork(self):
        stream = io.BytesIO()
        first = Image.new("RGBA", (512, 256), "blue")
        second = Image.new("RGBA", (512, 256), "green")
        first.save(stream, format="PNG", save_all=True, append_images=[second], duration=100)
        data = stream.getvalue()
        entry = copy.deepcopy(self.entry)
        entry["sha256"] = hashlib.sha256(data).hexdigest()
        self.write_archive(records=[entry], png=data)
        with self.assertRaisesRegex(p.PortableError, "static PNG"):
            self.catalog().resolve("azure-test")

    def test_opaque_black_png_is_visible_not_empty(self):
        stream = io.BytesIO()
        Image.new("RGB", (512, 256), "black").save(stream, format="PNG")
        data = stream.getvalue()
        entry = {**self.entry, "sha256": hashlib.sha256(data).hexdigest()}
        self.write_archive(records=[entry], png=data)
        self.assertEqual(self.catalog().resolve("azure-test")["png"], data)


if __name__ == "__main__":
    unittest.main()
