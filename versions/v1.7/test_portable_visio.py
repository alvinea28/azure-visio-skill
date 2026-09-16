"""Run: python -B -m unittest discover -s AzureVisio -p test_portable_visio.py.

Tests write exclusively into a newly created project-local directory and remove
only that directory afterwards. Synthetic test artwork is not an icon fallback.
PDF generation tests skip explicitly if declared dependencies are unavailable.
"""

import copy
import hashlib
import importlib.util
import io
import json
import math
from pathlib import Path
import re
import shutil
import types
import unittest
from unittest import mock
import uuid
import xml.etree.ElementTree as ET
import zipfile

import portable_visio as p


SVG = b"""<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 32 16">
<defs><linearGradient id="blue"><stop stop-color="#00aaff"/>
<stop offset="1" stop-color="#003399"/></linearGradient></defs>
<rect x="0" y="0" width="32" height="16" fill="url(#blue)"/></svg>"""
NARRATIVE = (
    "This illustrative architecture separates the application processing tier from durable data "
    "storage while preserving a clearly described request path. The application receives authorized "
    "requests, performs business validation, and retrieves only the required records. Operational "
    "owners must review identity permissions, network access, observability, recovery objectives, "
    "and cost assumptions before approving any production deployment. This note explains the "
    "drawing and does not certify the architecture."
)


def sample_model(icon_ref="azure-test", proposed=False):
    """Public smoke fixture; caller supplies an actual installed official iconRef."""
    main = {
        "name": "01 Main architecture", "view": "main", "role": "diagram",
        "width": 11.7, "height": 8.3, "title": "Application and durable storage",
        "subtitle": "Editable resources, semantic endpoints, and official glyphs", "furniture": True,
        "footer": "Illustrative architecture | Review before deployment",
        "nodes": [
            {"id": "system", "kind": "container", "label": "Application scope",
             "containerStyle": "boundary", "boundaryType": "system",
             "x": 5.85, "y": 4.0, "width": 10.6, "height": 5.2},
            {"id": "app", "kind": "card", "label": "Application processing service",
             "displayLabel": "Application", "details": "Canonical implementation details are retained here.",
             "cardStyle": "icon", "iconRef": icon_ref, "parent": "system",
             "x": 3.0, "y": 4.5, "width": 2.2, "height": 1.4},
            {"id": "data", "kind": "card", "label": "Durable data storage service",
             "displayLabel": "Data service", "cardStyle": "icon", "iconRef": icon_ref,
             "parent": "system", "x": 8.7, "y": 4.5, "width": 2.2, "height": 1.4}],
        "edges": [{"id": "query", "source": "app", "target": "data", "kind": "query",
                   "label": "Authorized query", "direction": "both", "sourceSide": "right",
                   "targetSide": "left", "sourcePosition": 0.5, "targetPosition": 0.5,
                   "dashed": False, "routeStyle": "orthogonal"}]}
    hardening = {
        "name": "02 Hardening review", "view": "hardening", "role": "notes",
        "width": 11.7, "height": 8.3, "title": "Hardening review",
        "subtitle": "No additional production controls are asserted by this illustrative fixture",
        "furniture": True, "footer": "Review outcome: not applicable",
        "nodes": [{"id": "review", "kind": "note", "label": NARRATIVE,
                   "x": 5.85, "y": 4.8, "width": 10.2, "height": 3.5}],
        "edges": []}
    flow = {
        "name": "03 Processing flow", "view": "flowchart", "role": "notes",
        "width": 11.7, "height": 8.3, "title": "Processing flow and narrative",
        "subtitle": "Detailed steps map to every main resource and relationship",
        "furniture": True, "footer": "Read alongside the main architecture",
        "nodes": [
            {"id": "step1", "kind": "card", "cardStyle": "detail",
             "label": "Receive and validate the authorized request before applying application business "
                      "rules and constructing a narrowly scoped data query.",
             "mainNodeIds": ["app"], "mainEdgeIds": ["query"],
             "x": 3.0, "y": 5.0, "width": 4.3, "height": 2.1},
            {"id": "step2", "kind": "card", "cardStyle": "detail",
             "label": "Retrieve the required records from durable storage and return the result to "
                      "the application without expanding the caller authorization scope.",
             "mainNodeIds": ["data"], "mainEdgeIds": [],
             "x": 8.7, "y": 5.0, "width": 4.3, "height": 2.1},
            {"id": "narrative", "kind": "note", "label": NARRATIVE,
             "x": 5.85, "y": 2.0, "width": 10.2, "height": 2.0}],
        "edges": [{"id": "step-flow", "source": "step1", "target": "step2",
                   "kind": "logical", "label": "Query", "direction": "forward",
                   "routeStyle": "straight"}]}
    status, reason = "not-applicable", "Illustrative fixture, not an existing customer workload."
    if proposed:
        hardening = copy.deepcopy(main)
        hardening.update(name="02 Hardening proposal", view="hardening",
                         title="Proposed private application access")
        hardening["nodes"][1]["displayLabel"] = "Private application"
        hardening["nodes"][1]["label"] = "Application processing service with private access"
        status = "proposed"
        reason = "Restrict application exposure while retaining the same application and storage scope."
    return {"schemaVersion": 1, "outputContract": p.CONTRACT,
            "presentationProfile": "enterprise", "conversionMode": "new-design",
            "title": "Portable native architecture smoke",
            "hardening": {"status": status, "reason": reason}, "pages": [main, hardening, flow]}


class PortableTests(unittest.TestCase):
    def setUp(self):
        self.root = Path(__file__).parent / (".portable-tests-" + uuid.uuid4().hex)
        self.root.mkdir()
        self.catalog_dir = self.root / "icons"
        icon = self.catalog_dir / "azure" / "test.svg"
        icon.parent.mkdir(parents=True)
        icon.write_bytes(SVG)
        self.entry = {"id": "azure-test", "path": "azure\\test.svg", "usable": True,
                      "sha256": hashlib.sha256(SVG).hexdigest()}
        self.write_catalog(self.entry)
        self.model = sample_model()

    def tearDown(self):
        shutil.rmtree(self.root)

    def write_catalog(self, entry):
        (self.catalog_dir / "catalog.json").write_text(json.dumps([entry]), encoding="utf-8")

    def invalid(self, model, pattern):
        with self.assertRaisesRegex(p.PortableError, pattern):
            p.validate_model(model)

    def need_render(self):
        missing = [k for k, v in p.capabilities()["dependencies"].items() if not v]
        if missing:
            self.skipTest("Unavailable declared renderer dependencies: " + ", ".join(missing))

    def assert_orthogonal(self, points, places=7):
        self.assertGreaterEqual(len(points), 2)
        for a, b in zip(points, points[1:]):
            self.assertTrue(round(a[0] - b[0], places) == 0 or round(a[1] - b[1], places) == 0,
                            f"Diagonal segment: {a} -> {b}")

    def routing_layouts(self, positions=((3, 5.4), (8.4, 3.4)), obstacle=False):
        nodes = copy.deepcopy(self.model["pages"][0]["nodes"][1:])
        for node, (x, y) in zip(nodes, positions):
            node.update(x=x, y=y)
        if obstacle:
            blocker = copy.deepcopy(nodes[0])
            blocker.update(id="blocker", x=5.6, y=4.5, displayLabel="Unrelated service")
            nodes.append(blocker)
        catalog = p.IconCatalog(self.catalog_dir)
        return {node["id"]: p.node_layout(node, catalog) for node in nodes}

    def assert_route_clear(self, info, nodes):
        self.assert_orthogonal(info["points"])
        self.assertEqual(info["points"][0], info["start"]["point"])
        self.assertEqual(info["points"][-1], info["end"]["point"])
        self.assertTrue(p.port_departure(info["start"], info["points"][1]))
        self.assertTrue(p.port_departure(info["end"], info["points"][-2]))
        self.assertTrue(p.route_is_clear(info["points"], p.routing_obstacles(nodes),
                                        p.endpoint_exemptions(info["edge"], info["start"], info["end"], nodes)))

    def test_valid_contract_and_positive_hardening_cases(self):
        for status in ("not-applicable", "already-enterprise", "proposed"):
            model = sample_model(proposed=status == "proposed")
            model["hardening"]["status"] = status
            self.assertTrue(p.validate_model(model)["valid"])

    def test_requirements_provenance_metadata_is_data_only(self):
        self.model.update(requirements=[{"id": "R01", "statement": "Keep the existing workload scope."}],
                          referenceNotes=[{"url": "https://learn.microsoft.com/", "role": "Product reference"}],
                          assumptions=["No deployment has been authorized."])
        self.assertTrue(p.validate_model(self.model)["valid"])
        self.model["assumptions"] = ["valid", {"invalid": "not a string"}]
        self.invalid(self.model, "assumption")

    def test_page_contract_and_type_failures(self):
        for key, value, pattern in (("schemaVersion", True, "schemaVersion"),
                                    ("outputContract", "old", "outputContract"),
                                    ("presentationProfile", "legacy", "presentationProfile"),
                                    ("pages", self.model["pages"][:2], "three pages")):
            with self.subTest(key=key):
                model = copy.deepcopy(self.model)
                model[key] = value
                self.invalid(model, pattern)
        model = copy.deepcopy(self.model)
        model["pages"][1]["role"] = "diagram"
        self.invalid(model, "role")
        model["pages"][1]["role"] = "notes"
        model["hardening"]["reason"] = ""
        self.invalid(model, "hardening.reason")

    def test_enums_reject_arrays_and_objects_with_explicit_errors(self):
        changes = (
            lambda m: m.update(outputContract=[p.CONTRACT]),
            lambda m: m.update(presentationProfile=["enterprise"]),
            lambda m: m["hardening"].update(status=["proposed"]),
            lambda m: m["pages"][0].update(view=["main"]),
            lambda m: m["pages"][0].update(role=["diagram"]),
            lambda m: m["pages"][0]["nodes"][0].update(boundaryType=["system"]),
            lambda m: m["pages"][0]["edges"][0].update(kind={"kind": "query"}),
            lambda m: m["pages"][0]["edges"][0].update(sourceSide=["right"]))
        for change in changes:
            model = sample_model()
            change(model)
            with self.assertRaises(p.PortableError):
                p.validate_model(model)

    def test_faithful_conversions_fail_closed(self):
        for mode in ("faithful", "reference-plus-proposal"):
            model = copy.deepcopy(self.model)
            model["conversionMode"] = mode
            self.invalid(model, "requires model_path")
        self.model["referenceContract"] = "source.json"
        self.invalid(self.model, "requires model_path")
        self.model["conversionMode"] = "unknown"
        self.invalid(self.model, "Unsupported source conversion")

    def test_reference_companion_invoked_for_modes_and_contract_property(self):
        model_path = self.root / "reference-model.json"
        model_path.write_text("{}", encoding="utf-8")
        companion = types.ModuleType("portable_reference")
        companion.validate_reference = mock.Mock(return_value={"valid": True, "source": {"verified": True}})
        for mode in ("new-design", "faithful", "reference-plus-proposal"):
            model = sample_model()
            model.update(conversionMode=mode, referenceContract="authoritative.json")
            with mock.patch.dict("sys.modules", {"portable_reference": companion}):
                result = p.validate_model(model, model_path)
            self.assertEqual(result["sourceVerification"], "verified-authoritative-contract")
            companion.validate_reference.assert_called_with(model, model_path.absolute(), None)
        companion.validate_reference.side_effect = ValueError("SourceHashMismatch")
        with mock.patch.dict("sys.modules", {"portable_reference": companion}):
            with self.assertRaisesRegex(p.PortableError, "SourceHashMismatch"):
                p.validate_model(model, model_path)
        with mock.patch.dict("sys.modules", {"portable_reference": None}):
            with self.assertRaisesRegex(p.PortableError, "verifier unavailable"):
                p.validate_model(model, model_path)

    def test_real_reference_companion_end_to_end(self):
        self.need_render()
        if importlib.util.find_spec("portable_reference") is None:
            self.skipTest("portable_reference.py companion is unavailable")
        model = sample_model()
        model.update(conversionMode="faithful", presentationProfile="reference",
                     referenceContract="source-contract.json")
        source = self.root / "authoritative-source.txt"
        source.write_text("Application scope contains an application processing service and "
                          "durable data storage service; an authorized query connects them.",
                          encoding="utf-8")
        main = model["pages"][0]
        for node in main["nodes"]:
            node["sourceId"] = "source-" + node["id"]
        for edge in main["edges"]:
            edge["sourceId"] = "source-" + edge["id"]
        contract = {
            "schemaVersion": 1, "mode": "source-faithful",
            "source": {"path": source.name, "sha256": hashlib.sha256(source.read_bytes()).hexdigest(),
                       "role": "authoritative-reference"},
            "referencePage": main["name"], "allowAdditionalPages": True,
            "components": [
                {"id": n["sourceId"], "label": n["label"], "requiredText": [],
                 "parent": "source-" + n["parent"] if n.get("parent") else None, "kind": n["kind"]}
                for n in main["nodes"]],
            "relationships": [
                {"id": e["sourceId"], "source": "source-" + e["source"], "target": "source-" + e["target"],
                 "direction": p.edge_direction(e), "requiredText": [e["label"]]}
                for e in main["edges"]],
            "layout": {"leftToRight": [["source-app", "source-data"]], "topToBottom": [],
                       "aspectRatio": main["width"] / main["height"], "aspectTolerance": 0.15},
            "unresolved": []}
        (self.root / "source-contract.json").write_text(json.dumps(contract), encoding="utf-8")
        model_path = self.root / "reference-model.json"
        model_path.write_text(json.dumps(model), encoding="utf-8")
        report = p.validate_model(model, model_path)
        self.assertTrue(report["referenceValidation"]["source"]["verified"])
        result = p.render(model, self.catalog_dir, self.root / "reference-output", model_path=model_path)
        self.assertEqual(result["sourceVerification"], "verified-authoritative-contract")
        self.assertEqual(result["inspection"]["sourceVerification"],
                         "not-revalidated-by-static-package-inspection")
        with zipfile.ZipFile(result["vsdx"]) as archive:
            page = ET.fromstring(archive.read("visio/pages/page1.xml"))
            value = page.find(f"{{{p.V}}}Shapes/{{{p.V}}}Shape[@NameU='av-query']/"
                              f"{{{p.V}}}Section[@N='Property']/{{{p.V}}}Row[@N='SourceIdRef']/"
                              f"{{{p.V}}}Cell[@N='Value']")
            self.assertEqual(value.get("V"), "source-query")
        proposal = sample_model(proposed=True)
        proposal.update(conversionMode="reference-plus-proposal", presentationProfile="reference",
                        referenceContract="source-contract.json")
        proposal["pages"][0] = copy.deepcopy(main)
        proposal_path = self.root / "proposal-model.json"
        proposal_path.write_text(json.dumps(proposal), encoding="utf-8")
        proposed_result = p.render(proposal, self.catalog_dir, self.root / "proposal-output",
                                   model_path=proposal_path)
        self.assertEqual(proposed_result["sourceVerification"], "verified-authoritative-contract")
        source.write_text("Tampered source", encoding="utf-8")
        with self.assertRaisesRegex(p.PortableError, "Reference fidelity validation failed"):
            p.render(model, self.catalog_dir, self.root / "tampered-output", model_path=model_path)
        self.assertFalse((self.root / "tampered-output").exists())

    def test_proposed_duplicate_rejected_even_after_relayout(self):
        self.model["hardening"]["status"] = "proposed"
        duplicate = copy.deepcopy(self.model["pages"][0])
        duplicate.update(name="Hardening", view="hardening", title="Different title")
        duplicate["nodes"][1]["x"] += 0.25
        self.model["pages"][1] = duplicate
        self.invalid(self.model, "cannot duplicate")

    def test_contract_matches_native_mapping_and_review_rules(self):
        model = sample_model()
        model["pages"][2]["nodes"][0]["cardStyle"] = "standard"
        model["pages"][2]["nodes"][0]["mainNodeIds"].append("system")
        self.assertTrue(p.validate_model(model)["valid"])
        mutations = (
            lambda m: m["pages"][0].update(edges=[]),
            lambda m: m["pages"][1]["nodes"].append(
                {"id": "extra", "kind": "card", "label": "Extra card",
                 "x": 3, "y": 2, "width": 2, "height": 1}),
            lambda m: m["pages"][2]["nodes"][1].update(mainNodeIds=[], mainEdgeIds=[]),
            lambda m: m["pages"][2]["nodes"][2].update(mainNodeIds=["system"]),
            lambda m: m["pages"][0]["nodes"][1].update(mainNodeIds=["app"]),
            lambda m: m["pages"][2]["edges"][0].update(direction="none"),
            lambda m: m["pages"][2]["edges"][0].update(target="narrative"))
        for change in mutations:
            with self.subTest(change=change):
                model = sample_model()
                change(model)
                with self.assertRaises(p.PortableError):
                    p.validate_model(model)

    def test_proposal_notes_or_style_changes_do_not_count_as_hardening(self):
        model = sample_model(proposed=True)
        duplicate = copy.deepcopy(model["pages"][0])
        duplicate.update(name="Hardening", view="hardening")
        duplicate["nodes"][1].update(displayLabel="New visible caption", color="RGB(1,2,3)")
        duplicate["nodes"].append({"id": "change-note", "kind": "note", "label": "A new presentation note",
                                   "x": 5.8, "y": 2, "width": 3, "height": 0.5})
        model["pages"][1] = duplicate
        self.invalid(model, "cannot duplicate")
        duplicate["nodes"][1]["purpose"] = "Enforce private access for the existing application."
        self.assertTrue(p.validate_model(model)["valid"])

    def test_missing_flow_text_mapping_and_disconnected(self):
        mutations = (
            lambda m: m["pages"][2]["nodes"][0].update(mainNodeIds=[]),
            lambda m: m["pages"][2]["nodes"][0].update(mainEdgeIds=["step-flow"]),
            lambda m: m["pages"][2]["nodes"][0].update(label="Too short"),
            lambda m: m["pages"][2]["nodes"][2].update(label="Too short"),
            lambda m: m["pages"][2].update(edges=[]))
        for change in mutations:
            with self.subTest(change=change):
                model = copy.deepcopy(self.model)
                change(model)
                self.invalid(model, "Flowchart|flowchart|page-one")

    def test_bounds_missing_endpoints_and_parent_cycles(self):
        mutations = (
            lambda m: m["pages"][0]["nodes"][1].update(x=-1),
            lambda m: m["pages"][0]["nodes"][1].update(width=True),
            lambda m: m["pages"][0]["nodes"][1].update(x=float("nan")),
            lambda m: m["pages"][0]["nodes"][1].update(parent="missing"),
            lambda m: m["pages"][0]["nodes"][0].update(parent="system"),
            lambda m: m["pages"][0]["edges"][0].update(target="missing"),
            lambda m: m["pages"][0]["edges"][0].update(sourcePosition=1.1),
            lambda m: m["pages"][0]["edges"][0].update(direction="sideways"),
            lambda m: m["pages"][0]["edges"][0].update(dashed="true"))
        for change in mutations:
            with self.subTest(change=change):
                model = copy.deepcopy(self.model)
                change(model)
                with self.assertRaises(p.PortableError):
                    p.validate_model(model)

    def test_style_font_and_no_silent_unknown_fields(self):
        self.model["pages"][0]["nodes"][1]["displayLabel"] = "one two three four five six seven eight nine"
        self.invalid(self.model, "Caption too long")
        self.model = sample_model()
        self.model["pages"][0]["nodes"][1]["displayLabel"] = "App \U0001f680"
        self.invalid(self.model, "unsupported visible text")
        self.model = sample_model()
        self.model["pages"][0]["nodes"][1]["font"] = "Nonexistent font"
        self.invalid(self.model, "Unsupported font")
        self.model = sample_model()
        self.model["pages"][0]["nodes"][1]["icon"] = "invented"
        self.invalid(self.model, "unsupported fields")

    def test_catalog_unresolved_hash_traversal_and_unsafe_svg(self):
        self.assertEqual(p.IconCatalog(self.catalog_dir).resolve("azure-test")["aspect"], 2)
        with self.assertRaisesRegex(p.PortableError, "Unresolved iconRef"):
            p.IconCatalog(self.catalog_dir).resolve("missing")
        for field, value, pattern in (("sha256", "0" * 64, "SHA256"),
                                      ("path", "..\\outside.svg", "Unsafe"),
                                      ("path", "/outside.svg", "Unsafe"),
                                      ("usable", False, "unusable")):
            with self.subTest(field=field):
                entry = dict(self.entry)
                entry[field] = value
                self.write_catalog(entry)
                with self.assertRaisesRegex(p.PortableError, pattern):
                    p.IconCatalog(self.catalog_dir).resolve("azure-test")

    def test_unsafe_svg_resources_and_unsupported_fonts(self):
        samples = (
            b'<svg xmlns="http://www.w3.org/2000/svg"><script/></svg>',
            b'<!DOCTYPE svg [<!ENTITY x "y">]><svg xmlns="http://www.w3.org/2000/svg"/>',
            b'<svg xmlns="http://www.w3.org/2000/svg"><image href="file:///secret"/></svg>',
            b'<svg xmlns="http://www.w3.org/2000/svg"><use href="https://example.com/a"/></svg>',
            b'<svg xmlns="http://www.w3.org/2000/svg"><text>Unverified font</text></svg>',
            b'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1 1"><path fill="url(#missing)"/></svg>',
            b'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1 1"><path fill="u&#114;l(https://example.com/a)"/></svg>',
            b'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1 1"><style>@import "x";</style></svg>')
        for data in samples:
            with self.subTest(data=data):
                with self.assertRaises(p.PortableError):
                    p.safe_svg(data)

    def test_json_duplicate_nan_and_path_safety(self):
        path = self.root / "bad.json"
        for data in ('{"a":1,"A":2}', '{"a":NaN}', '[' * 70 + '0' + ']' * 70):
            path.write_text(data, encoding="utf-8")
            with self.assertRaises(p.PortableError):
                p.load_json(path)
        for value in (self.root / ".." / "escape", "\\\\host\\share\\file", "https://example.com/file"):
            with self.assertRaises(p.PortableError):
                p.safe_path(value)

    def test_installer_ids_archive_safety_and_terms(self):
        relative = "azure\\Azure_Public_Service_Icons\\Icons\\ai + machine learning\\00028-icon-service-Batch-AI.svg"
        self.assertEqual(p.catalog_id(relative), "azure-5aed262dc9ab")
        self.assertEqual(p.catalog_id(relative.replace("\\", "/")), "azure-5aed262dc9ab")
        with mock.patch.object(p.urllib.request, "build_opener") as opener:
            with self.assertRaisesRegex(p.PortableError, "accept-icon-terms"):
                p.download_icons(self.root / "download")
            opener.assert_not_called()
        for names in (["../bad.svg"], ["bad.svg", "BAD.svg"]):
            buf = io.BytesIO()
            with zipfile.ZipFile(buf, "w") as archive:
                for name in names:
                    archive.writestr(name, SVG)
            with self.assertRaises(p.PortableError):
                p.extract_official_archive(buf.getvalue(), self.root / "download")
            self.assertFalse((self.root / "download").exists())

    def test_generate_native_structure_metadata_and_pdf_three_pages(self):
        self.need_render()
        result = p.render(self.model, self.catalog_dir, self.root / "output", bundle=True)
        self.assertTrue(Path(result["vsdx"]).is_file())
        self.assertTrue(Path(result["pdf"]).is_file())
        self.assertTrue(p.inspect_package(result["vsdx"])["valid"])
        ns = {"v": p.V}
        with zipfile.ZipFile(result["vsdx"]) as archive:
            root = ET.fromstring(archive.read("visio/pages/page1.xml"))
            groups = root.findall(".//v:Shape[@Type='Group']", ns)
            foreign = root.findall(".//v:Shape[@Type='Foreign']", ns)
            self.assertEqual(len(groups), 2)
            self.assertEqual(len(foreign), 2)
            connects = root.findall("v:Connects/v:Connect", ns)
            self.assertEqual(len(connects), 2)
            self.assertEqual({c.get("FromCell") for c in connects}, {"BeginX", "EndX"})
            shape = root.find("v:Shapes/v:Shape[@NameU='av-query']", ns)
            self.assertIsNotNone(shape.find("v:Cell[@N='BeginX']", ns))
            self.assertIn("PAR(PNT(", shape.find("v:Cell[@N='BeginX']", ns).get("F"))
            self.assertEqual(shape.find("v:Cell[@N='BeginArrow']", ns).get("V"), "13")
            self.assertEqual(shape.find("v:Cell[@N='EndArrow']", ns).get("V"), "13")
            for group in groups:
                self.assertIsNotNone(group.find("v:Section[@N='Connection']", ns))
                self.assertIsNotNone(group.find("v:Section[@N='Property']/v:Row[@N='FullLabel']", ns))
                self.assertIsNotNone(group.find("v:Shapes/v:Shape/v:Section[@N='Property']/v:Row[@N='AvRole']", ns))
                tags = [c.tag.rsplit("}", 1)[-1] for c in group]
                self.assertLess(tags.index("Shapes"), tags.index("Text"))
            document = ET.fromstring(archive.read("visio/document.xml"))
            contract = document.find("v:DocumentSheet/v:Section/v:Row[@N='OutputContract']/v:Cell[@N='Value']", ns)
            self.assertEqual(contract.get("V"), p.CONTRACT)
            stored = ET.fromstring(archive.read("customXml/item1.xml"))
            self.assertEqual(json.loads(stored.text), self.model)
            pngs = [n for n in archive.namelist() if n.endswith(".png")]
            self.assertEqual(len(pngs), 1)
        pdf = Path(result["pdf"]).read_bytes()
        with zipfile.ZipFile(result["bundle"]) as bundle:
            self.assertEqual(set(bundle.namelist()), {"architecture.vsdx", "architecture.pdf"})
            self.assertEqual(bundle.read("architecture.vsdx"), Path(result["vsdx"]).read_bytes())
            self.assertEqual(bundle.read("architecture.pdf"), pdf)
        self.assertIn("not a desktop Visio export", result["pdfCreation"])
        self.assertEqual(len(re.findall(rb"/Type\s*/Page\b", pdf)), 3)
        if importlib.util.find_spec("pypdf"):
            from pypdf import PdfReader
            reader = PdfReader(result["pdf"])
            self.assertEqual(len(reader.pages), 3)
            self.assertIn("Application", reader.pages[0].extract_text())
            flow = reader.pages[2].extract_text()
            self.assertIn("Receive and validate", flow)
            self.assertIn("certify the architecture", " ".join(flow.split()))
        with self.assertRaisesRegex(p.PortableError, "NEW"):
            p.render(self.model, self.catalog_dir, self.root / "output")
        self.assertEqual(Path(result["pdf"]).read_bytes(), pdf)

    def test_render_fails_before_creating_output_for_clipping(self):
        self.need_render()
        self.model["pages"][2]["nodes"][0]["height"] = 0.15
        with self.assertRaisesRegex(p.PortableError, "clip"):
            p.render(self.model, self.catalog_dir, self.root / "output")
        self.assertFalse((self.root / "output").exists())

    def test_missing_dependencies_reported_explicitly(self):
        with mock.patch.object(p, "capabilities", return_value={"dependencies": {"reportlab": False}}):
            with self.assertRaisesRegex(p.PortableError, "Unavailable rendering dependencies: reportlab"):
                p.dependencies()

    def test_relationship_defaults_match_native_controller(self):
        self.assertEqual(p.edge_direction({"kind": "governance"}), "none")
        self.assertEqual(p.edge_direction({"kind": "traffic"}), "forward")
        self.assertTrue(p.edge_dashed({"kind": "dependency"}))
        self.assertFalse(p.edge_dashed({"kind": "dependency", "dashed": False}))
        self.assertEqual(p.edge_direction({"kind": "dns", "direction": "both"}), "both")

    def test_geometry_anchor_routes_and_caption_semantics(self):
        self.need_render()
        catalog = p.IconCatalog(self.catalog_dir)
        layouts, _ = p.prepare(self.model, catalog)
        edge = layouts[0]["edges"][0]
        node = layouts[0]["nodes"]["app"]
        self.assertAlmostEqual(edge["start"]["point"][0], node["glyph"][0] + node["glyph"][2] / 2)
        self.assertAlmostEqual(node["glyph"][2] / node["glyph"][3], 2)
        self.assertEqual(node["text"].text, "Application")
        self.assertNotIn("Canonical", node["text"].text)
        self.assertEqual(layouts[0]["edges"][0]["text"].size, 10)
        self.assertEqual(dict(layouts[0]["furniture"])["subtitle"].size, 11)
        self.assertEqual(dict(layouts[0]["furniture"])["footer"].size, 10)

    def test_native_short_note_and_empty_root_parent_fit(self):
        self.need_render()
        self.model["pages"][0]["nodes"][0]["parent"] = ""
        self.model["pages"][0]["nodes"].append(
            {"id": "short-note", "kind": "note", "label": "Proposed controls need review before deployment.",
             "x": 5.85, "y": 1, "width": 10, "height": 0.35, "fontSize": 11})
        self.assertTrue(p.validate_model(self.model)["valid"])
        layout, _ = p.prepare(self.model, p.IconCatalog(self.catalog_dir))
        self.assertEqual(len(layout[0]["nodes"]["short-note"]["text"].lines), 1)

    def test_visible_label_ports_and_centered_caption(self):
        self.need_render()
        node = copy.deepcopy(self.model["pages"][0]["nodes"][1])
        node.update(cardStyle="label", label="Generic actor", displayLabel="Generic actor",
                    x=2, y=5, width=3, height=1.4)
        node.pop("iconRef")
        info = p.node_layout(node, None)
        self.assertEqual(info["text"].y, node["y"])
        self.assertLess(info["text"].height, node["height"])
        bounds = p.text_rect(info["text"])
        self.assertLess(bounds[2] - bounds[0], node["width"])
        target = self.routing_layouts()["data"]
        for side in p.SIDES:
            start, _ = p.anchors(info, target, {"sourceSide": side})
            if side in ("left", "right"):
                self.assertEqual(start["point"][0], round(bounds[0 if side == "left" else 2], 10))
                self.assertAlmostEqual(start["point"][1], node["y"])
            else:
                self.assertEqual(start["point"][1], round(bounds[1 if side == "bottom" else 3], 10))

    def test_all_side_pairs_and_positions_avoid_visible_obstacles(self):
        self.need_render()
        nodes = self.routing_layouts(obstacle=True)
        for source_side in sorted(p.SIDES):
            for target_side in sorted(p.SIDES):
                for position in (0, 0.35, 1):
                    with self.subTest(source=source_side, target=target_side, position=position):
                        edge = {"id": "route", "source": "app", "target": "data", "kind": "telemetry",
                                "sourceSide": source_side, "targetSide": target_side,
                                "sourcePosition": position, "targetPosition": 1 - position}
                        self.assert_route_clear(p.edge_layout(edge, nodes, (0.02, 0.02, 11.68, 8.28)), nodes)

    def test_straight_requests_become_orthogonal_without_diagonal_shortcuts(self):
        self.need_render()
        nodes = self.routing_layouts(obstacle=True)
        for direction in ("forward", "backward", "both", "none"):
            edge = {"id": "route", "source": "app", "target": "data", "kind": "logical",
                    "direction": direction, "routeStyle": "straight"}
            info = p.edge_layout(edge, nodes, (0.02, 0.02, 11.68, 8.28))
            self.assertGreater(len(info["points"]), 2)
            self.assert_route_clear(info, nodes)

    def test_aligned_direct_paths_remain_direct_and_clear(self):
        self.need_render()
        for positions in (((3, 5), (8, 5)), ((3, 5), (3, 2))):
            nodes = self.routing_layouts(positions)
            info = p.edge_layout({"id": "route", "source": "app", "target": "data",
                                  "kind": "logical", "routeStyle": "straight"}, nodes)
            self.assertEqual(len(info["points"]), 2)
            self.assert_route_clear(info, nodes)

    def test_intervening_service_blocks_direct_shortcut(self):
        self.need_render()
        nodes = self.routing_layouts(((3, 4.5), (8.4, 4.5)), obstacle=True)
        info = p.edge_layout({"id": "route", "source": "app", "target": "data",
                              "kind": "logical", "sourceSide": "right", "targetSide": "left",
                              "routeStyle": "straight"}, nodes)
        self.assertGreater(len(info["points"]), 2)
        self.assert_route_clear(info, nodes)
        for a, b in zip(info["points"], info["points"][1:]):
            for owner, _, box in p.routing_obstacles(nodes):
                if owner == "blocker":
                    self.assertFalse(p.segment_hits(a, b, p.inflate(box)))

    def test_bottom_icon_ports_attach_below_actual_caption(self):
        self.need_render()
        nodes = self.routing_layouts()
        for position in (0, 0.2, 0.5, 1):
            start, _ = p.anchors(nodes["app"], nodes["data"],
                                 {"sourceSide": "bottom", "sourcePosition": position})
            bounds = p.text_rect(nodes["app"]["text"])
            self.assertEqual(start["point"], (round(bounds[0] + (bounds[2] - bounds[0]) * position, 10),
                                              round(bounds[1], 10)))

    def test_explicit_diagonal_reversed_and_obstructed_routes_fail(self):
        self.need_render()
        nodes = self.routing_layouts(((3, 4.5), (8.4, 4.5)), obstacle=True)
        edge = {"id": "route", "source": "app", "target": "data", "kind": "logical",
                "sourceSide": "right", "targetSide": "left", "routeStyle": "straight"}
        start, end = p.anchors(nodes["app"], nodes["data"], edge)
        a, b = start["point"], end["point"]
        for points, message in (
                ([a, b], "intersects"),
                ([a, (5, 6), b], "diagonal"),
                ([a, (a[0] - 0.2, a[1]), (a[0] - 0.2, 7), (b[0] - 0.2, 7),
                  (b[0] - 0.2, b[1]), b], "outside")):
            with self.subTest(message=message):
                edge["points"] = [dict(x=x, y=y) for x, y in points]
                with self.assertRaisesRegex(p.PortableError, message):
                    p.edge_layout(edge, nodes, (0.02, 0.02, 11.68, 8.28))

    def test_blocked_port_and_no_path_fail_without_diagonal_fallback(self):
        self.need_render()
        nodes = self.routing_layouts(((3, 5), (8, 5)))
        edge = {"id": "route", "source": "app", "target": "data", "kind": "logical",
                "sourceSide": "right", "targetSide": "left"}
        start, _ = p.anchors(nodes["app"], nodes["data"], edge)
        x, y = start["point"]
        with self.assertRaisesRegex(p.PortableError, "port is blocked"):
            p.edge_layout(edge, nodes, (0, 0, 12, 8),
                          [("blocked-port", "body", (x + 0.02, y - 0.2, x + 0.3, y + 0.2))])
        with self.assertRaisesRegex(p.PortableError, "no obstacle-free orthogonal route"):
            p.edge_layout(edge, nodes, (0, 0, 12, 8), [("wall", "body", (5, 0, 6, 8))])

    def test_blocked_boundary_port_does_not_shortcut_through_note(self):
        self.need_render()
        nodes = self.routing_layouts()
        boundary = {"id": "scope", "kind": "container", "label": "Workload scope",
                    "containerStyle": "boundary", "boundaryType": "vnet",
                    "x": 6, "y": 4, "width": 4, "height": 4}
        note = {"id": "note", "kind": "note", "label": "An unrelated review note",
                "x": 6, "y": 1.8, "width": 3, "height": 0.6}
        nodes.update(scope=p.node_layout(boundary, None), note=p.node_layout(note, None))
        with self.assertRaisesRegex(p.PortableError, "port is blocked"):
            p.edge_layout({"id": "route", "source": "app", "target": "scope", "kind": "dns",
                           "sourceSide": "left", "targetSide": "bottom"}, nodes, (0, 0, 12, 8))

    def test_return_route_label_stays_beside_line_outside_cards(self):
        self.need_render()
        nodes = {}
        for key, y in (("app", 6), ("data", 3), ("middle", 4.5)):
            node = {"id": key, "kind": "card", "cardStyle": "detail",
                    "label": "Processing step", "x": 4.5, "y": y, "width": 8, "height": 1.2}
            nodes[key] = p.node_layout(node, None)
        info = p.edge_layout({"id": "return", "source": "app", "target": "data", "kind": "logical",
                              "sourceSide": "right", "targetSide": "right", "label": "Invalid"},
                             nodes, (0, 0, 12, 8))
        self.assert_route_clear(info, nodes)
        self.assertGreater(p.text_rect(info["text"])[0], 8.5)

    def test_label_reservations_protect_future_ports(self):
        self.need_render()
        nodes = self.routing_layouts(((3, 4.5), (8.4, 4.5)))
        edge = {"id": "route", "source": "app", "target": "data", "kind": "logical",
                "sourceSide": "right", "targetSide": "left", "label": "Integration"}
        raw = p.edge_layout(edge, nodes, (0, 0, 12, 8))
        x, y = raw["text"].x, raw["text"].y
        reserved = (x - 0.1, y - 0.24, x + 0.1, y + 0.24)
        info = p.edge_layout(edge, nodes, (0, 0, 12, 8),
                             label_obstacles=[("future-port", "port", reserved)])
        a, b, c, d = p.text_rect(info["text"])
        self.assertFalse(a < reserved[2] and c > reserved[0] and b < reserved[3] and d > reserved[1])

    def test_windows_autosize_icon_size_and_flow_only_native_properties(self):
        self.need_render()
        layouts, images = p.prepare(self.model, p.IconCatalog(self.catalog_dir))
        with zipfile.ZipFile(io.BytesIO(p.make_vsdx(self.model, layouts, images))) as archive:
            ns = {"v": p.V}
            windows = ET.fromstring(archive.read("visio/windows.xml"))
            self.assertEqual(windows.tag, f"{{{p.V}}}Windows")
            self.assertIsNotNone(windows.find("v:Window[@Page='0']", ns))
            self.assertFalse(windows.findall(".//v:ViewCenterX", ns))
            for sheet in ET.fromstring(archive.read("visio/pages/pages.xml")).findall("v:Page/v:PageSheet", ns):
                self.assertEqual(sheet.find("v:Cell[@N='DrawingResizeType']", ns).get("V"), "0")
            for page_index, layout in enumerate(layouts, 1):
                page = ET.fromstring(archive.read(f"visio/pages/page{page_index}.xml"))
                for node in layout["page"]["nodes"]:
                    shape = page.find(f"v:Shapes/v:Shape[@NameU='av-{node['id']}']", ns)
                    props = {r.get("N"): r.find("v:Cell[@N='Value']", ns).get("V")
                             for r in shape.findall("v:Section[@N='Property']/v:Row", ns)}
                    self.assertEqual("MainNodeIds" in props, "mainNodeIds" in node)
                    self.assertEqual("MainEdgeIds" in props, "mainEdgeIds" in node)
                    if layout["nodes"][node["id"]]["glyph"]:
                        self.assertGreater(float(props["IconSize"]), 0)
                        self.assertEqual(float(props["IconAspect"]), 2)

    def test_native_caption_children_match_visible_port_bounds(self):
        self.need_render()
        for style in ("icon", "label"):
            node = copy.deepcopy(self.model["pages"][0]["nodes"][1])
            node.update(cardStyle=style, x=3, y=5, width=3)
            if style == "label":
                node.pop("iconRef")
            layouts = self.routing_layouts()
            layouts["app"] = p.node_layout(node, p.IconCatalog(self.catalog_dir))
            original_box = copy.deepcopy(layouts["app"]["text"])
            bounds = p.text_rect(original_box)
            for position in (0, 0.35, 1):
                edge = {"id": "route", "source": "app", "target": "data", "kind": "logical",
                        "sourceSide": "bottom", "targetSide": "left", "sourcePosition": position}
                info = p.edge_layout(edge, layouts, (0, 0, 12, 8))
                page = {"nodes": [layout["node"] for layout in layouts.values()], "edges": [edge]}
                data, _ = p.native_page({"page": page, "nodes": layouts, "edges": [info], "furniture": []},
                                        {"azure-test": "image1.png"})
                root = ET.fromstring(data)
                group = root.find(f"{{{p.V}}}Shapes/{{{p.V}}}Shape[@NameU='av-app']")
                child = group.find(f"{{{p.V}}}Shapes/{{{p.V}}}Shape[@NameU='caption-app']")
                cells = {c.get("N"): float(c.get("V")) for c in child.findall(f"{{{p.V}}}Cell")
                         if c.get("N") in ("PinX", "PinY", "Width", "Height", "TxtWidth", "TxtHeight")}
                self.assertAlmostEqual(cells["Width"], bounds[2] - bounds[0])
                self.assertAlmostEqual(cells["Height"], bounds[3] - bounds[1])
                self.assertEqual(cells["TxtWidth"], cells["Width"])
                self.assertEqual(cells["TxtHeight"], cells["Height"])
                left = node["x"] - node["width"] / 2 + cells["PinX"] - cells["Width"] / 2
                bottom = node["y"] - node["height"] / 2 + cells["PinY"] - cells["Height"] / 2
                self.assertAlmostEqual(info["start"]["point"][0], left + cells["Width"] * position)
                self.assertAlmostEqual(info["start"]["point"][1], bottom)
                self.assertEqual(layouts["app"]["text"], original_box)

    def test_measured_captions_keep_native_font_tolerance_without_changing_card_margins(self):
        self.need_render()
        self.model["pages"][0]["nodes"][1]["displayLabel"] = "Azure Event Grid\nSystem topic"
        self.model["pages"][0]["nodes"][2]["displayLabel"] = "Blob container\nInput documents"
        layouts, images = p.prepare(self.model, p.IconCatalog(self.catalog_dir))
        with zipfile.ZipFile(io.BytesIO(p.make_vsdx(self.model, layouts, images))) as archive:
            for index in (1, 2, 3):
                root = ET.fromstring(archive.read(f"visio/pages/page{index}.xml"))
                for shape in root.iter(f"{{{p.V}}}Shape"):
                    if shape.find(f"{{{p.V}}}Text") is None:
                        continue
                    caption = shape.find(f"{{{p.V}}}Section[@N='Property']/{{{p.V}}}Row[@N='AvRole']/"
                                         f"{{{p.V}}}Cell[@N='Value']")
                    measured_caption = caption is not None and caption.get("V") == "caption"
                    for name in ("LeftMargin", "RightMargin", "TopMargin", "BottomMargin"):
                        expected = 0 if measured_caption and name in ("LeftMargin", "RightMargin") else 0.04
                        self.assertEqual(float(shape.find(f"{{{p.V}}}Cell[@N='{name}']").get("V")), expected)
                    if measured_caption:
                        width = float(shape.find(f"{{{p.V}}}Cell[@N='TxtWidth']").get("V"))
                        lines = "".join(shape.find(f"{{{p.V}}}Text").itertext()).splitlines()
                        pdfmetrics = p.dependencies()[0]
                        size = float(shape.find(f"{{{p.V}}}Section[@N='Character']/{{{p.V}}}Row/"
                                                f"{{{p.V}}}Cell[@N='Size']").get("V")) * 72
                        longest = max(pdfmetrics.stringWidth(line, "Helvetica", size) for line in lines) / 72
                        self.assertAlmostEqual(width - longest, 0.08)

    def test_native_inward_ports_oppose_outward_route_escape(self):
        self.need_render()
        nodes = self.routing_layouts()
        for side in sorted(p.SIDES):
            edge = {"id": "route", "source": "app", "target": "data", "kind": "logical",
                    "sourceSide": side, "targetSide": side}
            info = p.edge_layout(edge, nodes, (0, 0, 12, 8))
            self.assert_route_clear(info, nodes)
            page = {"nodes": [layout["node"] for layout in nodes.values()], "edges": [edge]}
            data, _ = p.native_page({"page": page, "nodes": nodes, "edges": [info], "furniture": []},
                                    {"azure-test": "image1.png"})
            root = ET.fromstring(data)
            for node_id in ("app", "data"):
                row = root.find(f"{{{p.V}}}Shapes/{{{p.V}}}Shape[@NameU='av-{node_id}']/"
                                f"{{{p.V}}}Section[@N='Connection']/{{{p.V}}}Row")
                self.assertEqual(row.find(f"{{{p.V}}}Cell[@N='Type']").get("V"), "0")
                vx, vy = p.PORT_DIRECTIONS[side]
                self.assertEqual(float(row.find(f"{{{p.V}}}Cell[@N='DirX']").get("V")), -vx)
                self.assertEqual(float(row.find(f"{{{p.V}}}Cell[@N='DirY']").get("V")), -vy)

    def test_native_reroute_disables_rounding_and_line_jumps(self):
        self.need_render()
        layouts, images = p.prepare(self.model, p.IconCatalog(self.catalog_dir))
        with zipfile.ZipFile(io.BytesIO(p.make_vsdx(self.model, layouts, images))) as archive:
            pages = ET.fromstring(archive.read("visio/pages/pages.xml"))
            for index, page in enumerate(pages, 1):
                sheet = page.find(f"{{{p.V}}}PageSheet")
                for name, value in (("RouteStyle", "1"), ("LineRouteExt", "1"), ("LineJumpCode", "0"),
                                    ("DrawingResizeType", "0")):
                    self.assertEqual(sheet.find(f"{{{p.V}}}Cell[@N='{name}']").get("V"), value)
                root = ET.fromstring(archive.read(f"visio/pages/page{index}.xml"))
                for shape in root.findall(f"{{{p.V}}}Shapes/{{{p.V}}}Shape"):
                    if shape.find(f"{{{p.V}}}Cell[@N='BeginX']") is None:
                        continue
                    for name, value in (("ShapeRouteStyle", "1"), ("ConFixedCode", "0"),
                                        ("ConLineRouteExt", "1"), ("ConLineJumpCode", "1"), ("Rounding", "0")):
                        self.assertEqual(shape.find(f"{{{p.V}}}Cell[@N='{name}']").get("V"), value)
                    for endpoint in ("Begin", "End"):
                        for axis in ("X", "Y"):
                            self.assertIn("PAR(PNT(", shape.find(f"{{{p.V}}}Cell[@N='{endpoint}{axis}']").get("F"))
                    for row in shape.findall(f"{{{p.V}}}Section[@N='Geometry']/{{{p.V}}}Row"):
                        for axis in ("X", "Y"):
                            self.assertIsNotNone(row.find(f"{{{p.V}}}Cell[@N='{axis}']").get("F"))

    def test_native_geometry_and_glue_stay_orthogonal_after_endpoint_moves(self):
        self.need_render()
        models = []
        for positions in (((3, 4.5), (8.7, 4.5)), ((3, 5), (3, 2.7)), ((3, 5.2), (8.7, 3.1))):
            model = copy.deepcopy(self.model)
            for node, (x, y) in zip(model["pages"][0]["nodes"][1:], positions):
                node.update(x=x, y=y)
            model["pages"][0]["edges"][0].pop("sourceSide")
            model["pages"][0]["edges"][0].pop("targetSide")
            models.append(model)
        model = copy.deepcopy(self.model)
        model["pages"][0]["edges"][0].update(sourceSide="bottom", targetSide="bottom")
        models.append(model)
        for model in models:
            layouts, images = p.prepare(model, p.IconCatalog(self.catalog_dir))
            data = p.make_vsdx(model, layouts, images)
            self.assertTrue(p.inspect_package(data)["valid"])
            with zipfile.ZipFile(io.BytesIO(data)) as archive:
                for index, layout in enumerate(layouts, 1):
                    root = ET.fromstring(archive.read(f"visio/pages/page{index}.xml"))
                    for info in layout["edges"]:
                        element = root.find(f"{{{p.V}}}Shapes/{{{p.V}}}Shape[@NameU='av-{info['edge']['id']}']")
                        self.assertEqual(element.find(f"{{{p.V}}}Cell[@N='ShapeRouteStyle']").get("V"), "1")
                        self.assertEqual(element.find(f"{{{p.V}}}Cell[@N='ConFixedCode']").get("V"), "0")
                        for shift in ((0, 0, 0, 0), (0.3, -0.7, -0.4, 0.2), (0, 0, 0.2, 1.3)):
                            a, b = info["points"][0], info["points"][-1]
                            ax, ay, bx, by = a[0] + shift[0], a[1] + shift[1], b[0] + shift[2], b[1] + shift[3]
                            angle = math.atan2(by - ay, bx - ax)
                            env = {"BeginX": ax, "BeginY": ay, "EndX": bx, "EndY": by,
                                   "Angle": angle, "Width": math.hypot(bx - ax, by - ay),
                                   "COS": math.cos, "SIN": math.sin}
                            points = []
                            for row in element.findall(f"{{{p.V}}}Section[@N='Geometry']/{{{p.V}}}Row"):
                                local = [eval(row.find(f"{{{p.V}}}Cell[@N='{axis}']").get("F"),
                                              {"__builtins__": {}}, env) for axis in ("X", "Y")]
                                points.append((ax + local[0] * math.cos(angle) - local[1] * math.sin(angle),
                                               ay + local[0] * math.sin(angle) + local[1] * math.cos(angle)))
                            self.assert_orthogonal(points)
                            self.assertAlmostEqual(math.dist(points[0], (ax, ay)), 0)
                            self.assertAlmostEqual(math.dist(points[-1], (bx, by)), 0)

    def test_pdf_uses_exact_prepared_connector_geometry(self):
        self.need_render()
        dependencies = p.dependencies()
        actual_canvas = dependencies[1]
        recorded = []

        class RecordingCanvas(actual_canvas):
            def drawPath(self, path, stroke=1, fill=0, fillMode=None):
                if stroke and not fill:
                    recorded.append(path.getCode())
                return super().drawPath(path, stroke, fill, fillMode)

        self.model["pages"][0]["nodes"][2]["y"] = 3.4
        layouts, images = p.prepare(self.model, p.IconCatalog(self.catalog_dir))
        with mock.patch.object(p, "dependencies", return_value=(dependencies[0], RecordingCanvas, *dependencies[2:])):
            data = p.make_pdf(self.model, layouts, images)
        self.assertTrue(data.startswith(b"%PDF-"))
        expected = []
        from reportlab.pdfgen.pathobject import PDFPathObject
        for layout in layouts:
            for info in layout["edges"]:
                path = PDFPathObject()
                path.moveTo(*(v * 72 for v in info["points"][0]))
                for point in info["points"][1:]:
                    path.lineTo(*(v * 72 for v in point))
                expected.append(path.getCode())
                self.assert_orthogonal(info["points"])
        self.assertEqual(recorded, expected)

    def test_inspection_rejects_missing_windows_and_tampered_diagonal(self):
        self.need_render()
        layouts, images = p.prepare(self.model, p.IconCatalog(self.catalog_dir))
        with zipfile.ZipFile(io.BytesIO(p.make_vsdx(self.model, layouts, images))) as archive:
            original = {name: archive.read(name) for name in archive.namelist()}
        for issue in ("windows", "diagonal", "target"):
            parts = dict(original)
            if issue == "windows":
                parts.pop("visio/windows.xml")
            else:
                root = ET.fromstring(parts["visio/pages/page1.xml"])
                if issue == "diagonal":
                    shape = root.find(f"{{{p.V}}}Shapes/{{{p.V}}}Shape[@NameU='av-query']")
                    shape.find(f"{{{p.V}}}Section[@N='Geometry']/{{{p.V}}}Row[@IX='2']/"
                               f"{{{p.V}}}Cell[@N='Y']").set("V", "0.15")
                else:
                    root.find(f"{{{p.V}}}Connects/{{{p.V}}}Connect").set("ToSheet", "3")
                parts["visio/pages/page1.xml"] = p.xml_bytes(root)
            output = io.BytesIO()
            with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as archive:
                for name, value in parts.items():
                    archive.writestr(name, value)
            with self.assertRaises(p.PortableError):
                p.inspect_package(output.getvalue())


if __name__ == "__main__":
    unittest.main()
