"""Run with python -m unittest discover -s AzureVisio -p test_portable_reference.py."""

from copy import deepcopy
import hashlib
import json
import os
from pathlib import Path
import shutil
import stat
from types import SimpleNamespace
import unittest
from unittest.mock import patch
import uuid

from portable_reference import (
    JSON_LIMIT, SOURCE_LIMIT, _check_stat, _local_path, validate_reference,
)


class ReferenceFidelityTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.root = Path.cwd() / ".copilot" / "session-state" / ("portable-reference-tests-" + uuid.uuid4().hex)
        cls.root.mkdir(parents=True)

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.root)

    def setUp(self):
        self.folder = self.root / uuid.uuid4().hex
        self.folder.mkdir()
        self.source_path = self.folder / "source.bin"
        self.source_bytes = b'Synthetic source bytes, NOT image recognition. $(throw "inert")'
        self.source_path.write_bytes(self.source_bytes)
        self.model_path = self.folder / "model.json"
        self.contract_path = self.folder / "contract.json"

        def node(source_id, kind, label, x, y, width, height, parent=""):
            return dict(id="n-" + source_id, sourceId="src-" + source_id,
                        kind=kind, label=label, x=x, y=y, width=width, height=height, parent=parent)

        self.model = {
            "schemaVersion": 1, "title": "Synthetic faithful fixture",
            "conversionMode": "faithful", "referenceContract": "contract.json",
            "pages": [{
                "name": "01 Reference architecture", "width": 20, "height": 10,
                "nodes": [
                    node("governance", "card", "Source governance\nPolicy and audit", 10, 9, 18, 1),
                    node("lane", "container", "Application boundary", 5, 5, 8, 6),
                    node("api", "card", "API tier\nOriginal processing detail", 5, 6.5, 6, 1, "n-lane"),
                    node("db", "card", "Data store\nOriginal persistence detail", 5, 3.5, 6, 1, "n-lane"),
                    node("note", "note", "Source note\nDiagram arrow is conceptual", 15, 5, 7, 2),
                    node("identity", "card", "Identity layer\nOriginal identity detail", 15, 7.5, 7, 1),
                ],
                "edges": [
                    dict(id="e-api-db", sourceId="src-api-db", source="n-api", target="n-db",
                         direction="forward", kind="association", label="Conceptual linkage only"),
                    dict(id="e-identity-api", sourceId="src-identity-api", source="n-identity",
                         target="n-api", direction="none", kind="association", label=""),
                ],
            }],
        }
        self.contract = {
            "schemaVersion": 1, "mode": "source-faithful",
            "source": {"path": "source.bin", "sha256": hashlib.sha256(self.source_bytes).hexdigest(),
                       "role": "authoritative-reference"},
            "referencePage": "01 Reference architecture", "unresolved": [],
            "components": [
                dict(id="src-governance", label="Source governance", requiredText=["Source governance", "Policy and audit"], parent=None, kind="card"),
                dict(id="src-lane", label="Application boundary", requiredText=[], parent=None, kind="container"),
                dict(id="src-api", label="API tier", requiredText=["Original processing detail"], parent="src-lane", kind="card"),
                dict(id="src-db", label="Data store", requiredText=["Original persistence detail"], parent="src-lane", kind="card"),
                dict(id="src-note", label="Source note", requiredText=["Diagram arrow is conceptual"], parent=None, kind="note"),
                dict(id="src-identity", label="Identity layer", requiredText=["Original identity detail"], parent=None, kind="card"),
            ],
            "relationships": [
                dict(id="src-api-db", source="src-api", target="src-db", direction="forward",
                     requiredText=["Conceptual linkage only"]),
                dict(id="src-identity-api", source="src-identity", target="src-api", direction="none"),
            ],
            "layout": {"leftToRight": [["src-lane", "src-note"]],
                       "topToBottom": [["src-governance", "src-lane"], ["src-api", "src-db"],
                                       ["src-identity", "src-note"]], "aspectRatio": 2},
        }
        self.write_json(self.model_path, self.model)
        self.write_json(self.contract_path, self.contract)

    @staticmethod
    def write_json(path, value):
        path.write_text(json.dumps(value, ensure_ascii=False), encoding="utf-8")

    def gate(self, model=None, contract=None):
        if contract is not None:
            self.write_json(self.contract_path, contract)
        return validate_reference(self.model if model is None else model, self.model_path)

    def reject(self, code, model=None, contract=None):
        with self.assertRaisesRegex(ValueError, code):
            self.gate(model, contract)

    def test_baseline_report_and_input_immutability(self):
        before = deepcopy(self.model)
        hashes = [hashlib.sha256(path.read_bytes()).hexdigest()
                  for path in (self.model_path, self.contract_path, self.source_path)]
        result = self.gate()
        self.assertIs(result["valid"], True)
        self.assertIs(result["source"]["verified"], True)
        self.assertEqual(result["source"]["actualSha256"], self.contract["source"]["sha256"])
        self.assertEqual(result["counts"], dict(expectedComponents=6, matchedComponents=6,
                                                expectedRelationships=2, matchedRelationships=2))
        for field in ("comUsed", "imageRecognitionPerformed", "dataFlowVerified"):
            self.assertIs(result[field], False)
        self.assertEqual(self.model, before)
        self.assertEqual(hashes, [hashlib.sha256(path.read_bytes()).hexdigest()
                                 for path in (self.model_path, self.contract_path, self.source_path)])

    def test_explicit_contract_override_and_absolute_source(self):
        self.model.pop("referenceContract")
        self.contract["source"]["path"] = str(self.source_path)
        self.write_json(self.contract_path, self.contract)
        result = validate_reference(self.model, self.model_path, self.contract_path)
        self.assertTrue(result["valid"])
        self.assertEqual(result["referencePath"], str(self.contract_path))

    def test_contained_contract_and_source(self):
        nested = self.folder / "nested"
        nested.mkdir()
        (nested / "source.bin").write_bytes(self.source_bytes)
        self.write_json(nested / "contract.json", self.contract)
        self.model["referenceContract"] = str(Path("nested") / "contract.json")
        self.assertTrue(self.gate()["valid"])

    def test_normalized_complete_canonical_lines(self):
        self.model["pages"][0]["nodes"][2]["label"] = "  api    TIER \r\n ORIGINAL   PROCESSING detail "
        self.assertTrue(self.gate()["valid"])

    def test_visible_caption_and_hidden_details_cannot_replace_canonical_text(self):
        self.model["pages"][0]["nodes"][2].update(
            label="API tier", displayLabel="API tier", details="Original processing detail")
        self.reject("MissingText")

    def test_text_drift_duplicates_and_title_order(self):
        for label, code in [
            ("Replacement service\nOriginal processing detail", "TitleMismatch"),
            ("API tier alternative\nOriginal processing detail", "TitleMismatch"),
            ("API tier", "MissingText"),
            ("API tier\nNOT Original processing detail", "UnexpectedText"),
            ("API tier\nOriginal processing detail\nNew service", "UnexpectedText"),
            ("API tier\nOriginal processing detail\nOriginal processing detail", "DuplicateText"),
            ("Original processing detail\nAPI tier", "TitleMismatch"),
        ]:
            with self.subTest(label=label):
                model = deepcopy(self.model)
                model["pages"][0]["nodes"][2]["label"] = label
                self.reject(code, model)

    def test_labels_remain_inert_unicode_text(self):
        payload = '__import__("os").system("not executed"); https://example.invalid \U0001f512'
        self.contract["components"][4]["label"] = payload
        self.model["pages"][0]["nodes"][4]["label"] = payload + "\nDiagram arrow is conceptual"
        self.assertTrue(self.gate(contract=self.contract)["valid"])

    def test_missing_changed_empty_and_oversized_source(self):
        self.source_path.unlink()
        self.reject("MissingSource")
        self.source_path.write_bytes(b"changed")
        self.reject("SourceHashMismatch")
        self.source_path.write_bytes(b"")
        self.reject("InputLimit")
        with self.source_path.open("wb") as stream:
            stream.truncate(SOURCE_LIMIT + 1)
        self.reject("InputLimit")

    def test_source_role_and_separate_artifacts(self):
        contract = deepcopy(self.contract)
        contract["source"]["role"] = "output-regression"
        self.reject("SourceRole", contract=contract)
        for source in (self.model_path, self.contract_path):
            with self.subTest(source=source.name):
                contract = deepcopy(self.contract)
                contract["source"]["path"] = str(source)
                self.reject("SourceRole", contract=contract)

    def test_hardlink_source_alias_rejected(self):
        alias = self.folder / "alias.json"
        try:
            os.link(self.model_path, alias)
        except OSError as error:
            self.skipTest(f"Hardlinks unavailable: {error}")
        self.contract["source"]["path"] = alias.name
        self.reject("SourceRole", contract=self.contract)

    def test_exact_source_ids_and_case_sensitive_identity(self):
        for variant in ("missing", "case", "duplicate", "extra-note"):
            model = deepcopy(self.model)
            nodes = model["pages"][0]["nodes"]
            if variant == "missing":
                del nodes[1]["sourceId"]
            elif variant == "case":
                nodes[1]["sourceId"] = "SRC-LANE"
            elif variant == "duplicate":
                nodes[1]["sourceId"] = nodes[0]["sourceId"]
            else:
                extra = deepcopy(nodes[4])
                extra.update(id="extra", sourceId="extra", annotation=True)
                nodes.append(extra)
            with self.subTest(variant=variant):
                self.reject("UnexpectedItem|DuplicateSourceId", model)
        extra = deepcopy(self.model["pages"][0]["nodes"][2])
        extra.update(id="different-case", sourceId="SRC-API")
        self.model["pages"][0]["nodes"].append(extra)
        expected = deepcopy(self.contract["components"][2])
        expected["id"] = "SRC-API"
        self.contract["components"].append(expected)
        self.assertTrue(self.gate(contract=self.contract)["valid"])

    def test_missing_components_edges_and_duplicate_native_ids(self):
        for category in ("nodes", "edges"):
            model = deepcopy(self.model)
            model["pages"][0][category].pop()
            with self.subTest(category=category):
                self.reject("MissingItem", model)
        for duplicate in ("n-GOVERNANCE", "n-governance"):
            model = deepcopy(self.model)
            model["pages"][0]["nodes"][1]["id"] = duplicate
            self.reject("DuplicateModelId", model)
        contract = deepcopy(self.contract)
        contract["relationships"][0]["id"] = "src-api"
        self.reject("DuplicateSourceId", contract=contract)

    def test_parent_and_component_kind_preserved(self):
        for field, value, code in [
            ("parent", "", "ParentMismatch"), ("parent", "missing", "ParentMismatch"),
            ("parent", "n-governance", "ParentMismatch"), ("kind", "note", "KindMismatch"),
        ]:
            model = deepcopy(self.model)
            model["pages"][0]["nodes"][2][field] = value
            with self.subTest(field=field, value=value):
                self.reject(code, model)

    def test_nested_source_hierarchy(self):
        parent = dict(id="n-inner", sourceId="src-inner", kind="container", label="Inner boundary",
                      x=5, y=5, width=6, height=4, parent="n-lane")
        self.model["pages"][0]["nodes"].append(parent)
        self.contract["components"].append(dict(id="src-inner", label="Inner boundary",
                                               requiredText=[], parent="src-lane", kind="container"))
        for index in (2, 3):
            self.model["pages"][0]["nodes"][index]["parent"] = "n-inner"
            self.contract["components"][index]["parent"] = "src-inner"
        self.assertTrue(self.gate(contract=self.contract)["valid"])
        self.model["pages"][0]["nodes"][2]["parent"] = "n-lane"
        self.reject("ParentMismatch")

    def test_relationship_endpoints_directions_and_labels(self):
        for field, value, code in [
            ("source", "n-db", "EndpointMismatch"), ("target", "n-api", "EndpointMismatch"),
            ("source", None, "InvalidInput"), ("direction", "backward", "DirectionMismatch"),
            ("label", "", "MissingText"), ("label", "NOT Conceptual linkage only", "UnexpectedText"),
            ("label", "Conceptual linkage only\nConceptual linkage only", "DuplicateText"),
        ]:
            model = deepcopy(self.model)
            model["pages"][0]["edges"][0][field] = value
            with self.subTest(field=field, value=value):
                self.reject(code, model)
        model = deepcopy(self.model)
        del model["pages"][0]["edges"][0]["direction"]
        self.reject("InvalidInput", model)

    def test_all_four_explicit_directions(self):
        for direction in ("forward", "backward", "both", "none"):
            self.model["pages"][0]["edges"][0]["direction"] = direction
            self.contract["relationships"][0]["direction"] = direction
            with self.subTest(direction=direction):
                self.assertTrue(self.gate(contract=self.contract)["valid"])

    def test_layout_order_overlap_touching_and_aspect(self):
        for index, field, value, code in [
            (3, "y", 6.5, "LayoutMismatch"), (3, "y", 5.5, "LayoutMismatch"),
            (4, "x", 1, "LayoutMismatch"), (2, "y", 3, "LayoutMismatch"),
        ]:
            model = deepcopy(self.model)
            model["pages"][0]["nodes"][index][field] = value
            with self.subTest(index=index, value=value):
                self.reject(code, model)
        model = deepcopy(self.model)
        model["pages"][0]["width"] = 40
        self.reject("AspectMismatch", model)
        self.contract["layout"]["aspectTolerance"] = 0.1
        model["pages"][0]["width"] = 21
        self.assertTrue(self.gate(model, self.contract)["valid"])

    def test_designated_page_additional_pages_permission(self):
        model = deepcopy(self.model)
        model["pages"][0]["name"] = "Proposal replacing source"
        self.reject("MissingReferencePage", model)
        model = deepcopy(self.model)
        extra = deepcopy(model["pages"][0])
        extra["name"] = "02 Proposal"
        model["pages"].append(extra)
        self.reject("AdditionalPages", model)
        self.contract["allowAdditionalPages"] = True
        self.assertTrue(self.gate(model, self.contract)["valid"])
        model["pages"][0]["nodes"] = []
        model["pages"][0]["edges"] = []
        self.reject("MissingItem", model)

    def pack(self):
        model = deepcopy(self.model)
        model["outputContract"] = "architecture-pack-v1.6"
        model["pages"][0]["view"] = "main"
        for view in ("hardening", "flowchart"):
            model["pages"].append(dict(name=view, view=view, role="notes", width=20, height=10,
                                       nodes=[], edges=[]))
        return model

    def test_v16_reference_page_is_first_and_extra_pages_are_authorized(self):
        model = self.pack()
        self.reject("AdditionalPages", model)
        self.contract["allowAdditionalPages"] = True
        self.assertTrue(self.gate(model, self.contract)["valid"])
        model["pages"][0], model["pages"][1] = model["pages"][1], model["pages"][0]
        self.reject("ReferencePageOrder", model)
        model["pages"][0]["view"] = "main"
        model["pages"][1]["view"] = "hardening"
        self.reject("ReferencePageOrder", model)

    def test_v16_page_count_and_unknown_contract(self):
        self.contract["allowAdditionalPages"] = True
        self.write_json(self.contract_path, self.contract)
        for count in (1, 2, 4):
            model = self.pack()
            if count == 4:
                extra = deepcopy(model["pages"][-1])
                extra["name"] = "extra"
                model["pages"].append(extra)
            else:
                model["pages"] = model["pages"][:count]
            with self.subTest(count=count):
                self.reject("ReferencePageOrder", model)
        model = self.pack()
        model["outputContract"] = "other"
        self.reject("Unknown outputContract", model)

    def test_strict_contract_schema_and_unresolved_extraction(self):
        mutations = [
            ("unknown field", lambda c: c.update(allowExtraPages=True)),
            ("unknown source field", lambda c: c["source"].update(downloadUrl="https://example.invalid")),
            ("unresolved", lambda c: c.update(unresolved=["Unreadable source text"])),
            ("mode", lambda c: c.update(mode="requirements-only")),
            ("boolean string", lambda c: c.update(allowAdditionalPages="true")),
            ("version string", lambda c: c.update(schemaVersion="1")),
            ("version bool", lambda c: c.update(schemaVersion=True)),
            ("components object", lambda c: c.update(components=c["components"][0])),
            ("details string", lambda c: c["components"][0].update(requiredText="Policy and audit")),
            ("bad parent", lambda c: c["components"][2].update(parent="src-note")),
            ("parent cycle", lambda c: c["components"][1].update(parent="src-lane")),
            ("unknown layout", lambda c: c["layout"].update(leftToRight=[["absent", "src-lane"]])),
            ("flat layout", lambda c: c["layout"].update(leftToRight=["src-lane", "src-note"])),
            ("repeated layout", lambda c: c["layout"].update(leftToRight=[["src-lane", "src-lane"]])),
            ("unknown endpoint", lambda c: c["relationships"][0].update(target="absent")),
            ("bad hash", lambda c: c["source"].update(sha256="not-a-hash")),
            ("multiline title", lambda c: c["components"][0].update(label="Source\ngovernance")),
            ("multiline detail", lambda c: c["components"][0].update(requiredText=["Policy\nand audit"])),
            ("bad tolerance", lambda c: c["layout"].update(aspectTolerance=1)),
            ("empty components", lambda c: c.update(components=[])),
        ]
        for name, mutate in mutations:
            with self.subTest(name=name):
                contract = deepcopy(self.contract)
                mutate(contract)
                self.reject("Reference fidelity failed", contract=contract)

    def test_typed_model_structure(self):
        for field, value in [
            ("width", "18"), ("height", False), ("x", float("nan")), ("y", float("inf")),
            ("label", {"text": "Source governance"}), ("id", ["n-governance"]),
            ("sourceId", None), ("kind", "unsupported"),
        ]:
            model = deepcopy(self.model)
            model["pages"][0]["nodes"][0][field] = value
            with self.subTest(field=field):
                self.reject("InvalidInput", model)
        for field, value in [("pages", {}), ("schemaVersion", True), ("conversionMode", "guess")]:
            model = deepcopy(self.model)
            model[field] = value
            self.reject("InvalidInput", model)

    def test_duplicate_json_properties_root_types_and_encoding(self):
        for raw in [
            b"null", b"1", b"true", b"[]", b'{"schemaVersion":1,"schemaVersion":1}',
            b'{"schemaVersion":1,"SCHEMAVERSION":1}',
            b'{"schemaVersion":1,"schema\\u0056ersion":1}', b'{"value":NaN}', b"{\xc3(}",
        ]:
            with self.subTest(raw=raw):
                self.contract_path.write_bytes(raw)
                self.reject("InvalidInput|DuplicateProperty")
                self.write_json(self.contract_path, self.contract)
                self.model_path.write_bytes(raw)
                self.reject("InvalidInput|DuplicateProperty")
                self.write_json(self.model_path, self.model)
        for encoding in ("utf-8-sig", "utf-16"):
            self.contract_path.write_bytes(json.dumps(self.contract).encode(encoding))
            self.assertTrue(self.gate()["valid"])

    def test_missing_contract_and_model_files(self):
        self.model.pop("referenceContract")
        self.reject("MissingContract")
        self.model["referenceContract"] = "missing.json"
        self.reject("InvalidInput")
        self.model_path.unlink()
        self.reject("InvalidInput")

    def test_unsafe_source_and_contract_paths(self):
        for path in [
            "https://example.invalid/source.png", r"\\server\share\source.png",
            r"..\source.bin", r"source-child\..\source.bin", "../source.bin",
            "source.bin:stream", "C:source.bin", r"\\?\C:\source.bin", "NUL",
            "source.bin.", "source.bin ", "source//source.bin", "source/./source.bin",
        ]:
            with self.subTest(path=path):
                contract = deepcopy(self.contract)
                contract["source"]["path"] = path
                self.reject("UnsafePath", contract=contract)
                self.write_json(self.contract_path, self.contract)
                model = deepcopy(self.model)
                model["referenceContract"] = path
                self.reject("UnsafePath", model)
        with self.assertRaisesRegex(ValueError, "UnsafePath"):
            validate_reference(self.model, Path("relative-model.json"), self.contract_path)

    def test_links_junctions_and_recall_attributes_fail_closed(self):
        for info in [
            SimpleNamespace(st_mode=stat.S_IFLNK, st_file_attributes=0),
            SimpleNamespace(st_mode=stat.S_IFDIR, st_file_attributes=0x400, st_reparse_tag=0xA0000003),
            SimpleNamespace(st_mode=stat.S_IFREG, st_file_attributes=0x1000),
            SimpleNamespace(st_mode=stat.S_IFREG, st_file_attributes=0x40000),
            SimpleNamespace(st_mode=stat.S_IFREG, st_file_attributes=0x400000),
        ]:
            with self.subTest(info=info):
                with self.assertRaisesRegex(ValueError, "UnsafePath"):
                    _check_stat(info, self.source_path)
        _check_stat(SimpleNamespace(st_mode=stat.S_IFREG, st_file_attributes=0x400,
                                    st_reparse_tag=0x9000101A), self.source_path)
        with patch.object(Path, "lstat", return_value=SimpleNamespace(st_mode=stat.S_IFLNK)):
            with self.assertRaisesRegex(ValueError, "UnsafePath"):
                _local_path(self.source_path)

    def test_json_size_depth_and_record_bounds(self):
        for path in (self.model_path, self.contract_path):
            with path.open("wb") as stream:
                stream.truncate(JSON_LIMIT + 1)
            self.reject("InputLimit")
            self.write_json(self.model_path, self.model)
            self.write_json(self.contract_path, self.contract)
        self.contract_path.write_text("[" * 65 + "0" + "]" * 65, encoding="utf-8")
        self.reject("InputLimit")
        self.contract["components"] = [self.contract["components"][0]] * 2001
        self.reject("InvalidInput", contract=self.contract)
        self.write_json(self.contract_path, {})
        model = deepcopy(self.model)
        cursor = model
        for _ in range(65):
            cursor["nested"] = {}
            cursor = cursor["nested"]
        self.reject("InputLimit", model)


if __name__ == "__main__":
    unittest.main()
