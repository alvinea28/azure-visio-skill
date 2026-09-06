#!/usr/bin/env python3
"""Azure Visio 1.6: local, cross-platform NEW VSDX/PDF creation without Office.

Python 3.10+ is required, not presumed available in any Copilot product. Run
`python portable_visio.py capabilities` in the actual execution environment.
Install requirements-portable.txt only if package installation is permitted.
No Scout handoff, shell execution, Office automation, upload, or cloud-file API.

Coordinates are center-based inches, Y upward. PDF and editable native shapes
share one layout. Official SVG artwork is rendered by resvg at 600 dpi to
transparent PNG, preserving gradients and aspect ratio, and embedded as Foreign
children of native groups. Captions remain native/extractable text. No template
or third-party VSDX implementation is used. The OPC/ShapeSheet implementation
follows Microsoft's published Visio 2012 XML schema documentation:
https://learn.microsoft.com/office/client-developer/visio/introduction-to-the-visio-file-formatvsdx
https://learn.microsoft.com/office/client-developer/visio/shape-element-shapes_type-complextypevisio-xml

Supported: architecture-pack-v1.6 models. Faithful/reference-plus-proposal models
and any model naming referenceContract must pass the bundled portable_reference
source-contract verifier before rendering. Missing verifier/source fails closed.
Screenshot extraction remains the caller's responsibility: persist the source,
extract an authoritative contract, and resolve ambiguities before conversion.
Metadata is preserved as Prop.* Shape Data, document/page
ShapeSheets, and a related custom XML part containing the complete JSON model.
Containers use semantic ParentId membership, not Visio ContainerProperties.

Visible text supports the Windows-1252 repertoire with Arial (native) and
metrically compatible Helvetica (PDF); other fonts/scripts fail explicitly.
No system font lookup or font substitution is performed by the PDF renderer.
This limitation does not remove Unicode retained solely in semantic metadata.
SVG text/fonts and active/external resources are rejected, not approximated.
Native Visio reopening/visual review is an independent acceptance step; static
inspection is not a claim of native acceptance or architecture correctness.
"""

from __future__ import annotations

import argparse
import hashlib
import heapq
import importlib.util
import io
import json
import math
import os
from pathlib import Path, PurePosixPath
import re
import stat
import sys
import urllib.error
import urllib.request
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass, field, replace


OFFICIAL_URL = "https://arch-center.azureedge.net/icons/Azure_Public_Service_Icons_V24.zip"
ICON_TERMS = "https://learn.microsoft.com/azure/architecture/icons/"
CONTRACT = "architecture-pack-v1.6"
V = "http://schemas.microsoft.com/office/visio/2012/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
VR = "http://schemas.microsoft.com/visio/2010/relationships/"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
ET.register_namespace("", V)
ET.register_namespace("r", R)
MAX_JSON = 10 * 1024 * 1024
MAX_SVG = 4 * 1024 * 1024
MAX_ZIP = 100 * 1024 * 1024
MAX_EXPANDED = 500 * 1024 * 1024
SIDES = {"left", "right", "top", "bottom"}
KINDS = {"logical", "query", "traffic", "ingestion", "dependency", "association",
         "telemetry", "governance", "dns", "peering"}
BOUNDARIES = {"cloud", "tenant", "region", "subscription", "resource-group",
              "vnet", "subnet", "availability-zone", "cluster", "environment",
              "system", "trust", "control-plane", "functional"}
COLORS = {"logical": "#475569", "query": "#6d28d9", "traffic": "#c2630d",
          "ingestion": "#0891b2", "dependency": "#64748b", "association": "#64748b",
          "telemetry": "#05825e", "governance": "#64748b", "dns": "#7c3aed",
          "peering": "#2563eb"}
PORT_DIRECTIONS = {"left": (-1, 0), "right": (1, 0), "top": (0, 1), "bottom": (0, -1)}
ROUTE_CLEARANCE = 0.10
PORT_ESCAPE = 0.14
EPSILON = 1e-8
IDENT = re.compile(r"[A-Za-z0-9][A-Za-z0-9_.-]{0,127}\Z")
RGB = re.compile(r"RGB\(\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})\s*\)\Z")


class PortableError(ValueError):
    """An actionable validation, capability, or resource error."""


def require(condition, message):
    if not condition:
        raise PortableError(message)


def object_fields(value, required, optional=(), context="object"):
    require(isinstance(value, dict), f"{context} must be an object")
    require(set(required) <= value.keys(), f"{context}: missing {set(required) - value.keys()}")
    require(value.keys() <= set(required) | set(optional),
            f"{context}: unsupported fields {value.keys() - set(required) - set(optional)}")


def text(value, context, empty=False):
    require(isinstance(value, str) and len(value) <= 32768
            and (empty or value.strip()), f"{context} must be a bounded string")
    require(not any(ord(c) < 32 and c not in "\n\r\t" or 0xD800 <= ord(c) <= 0xDFFF
                    or ord(c) in (0xFFFE, 0xFFFF) for c in value),
            f"{context}: unsupported XML control character")
    return value


def number(value, context, minimum=0, maximum=200):
    require(type(value) in (int, float) and math.isfinite(value)
            and minimum <= value <= maximum,
            f"{context} must be finite, within {minimum}..{maximum}")
    return value


def array(value, context, maximum=2000):
    require(isinstance(value, list) and len(value) <= maximum,
            f"{context} must be an array of at most {maximum} entries")
    return value


def identifier(value, context="id"):
    require(isinstance(value, str) and IDENT.fullmatch(value), f"Invalid {context}: {value!r}")
    return value


def words(value):
    return len(re.findall(r"[^\W_]+(?:[-/][^\W_]+)*", value, re.UNICODE))


def visible_text(value, context):
    text(value, context, empty=True)
    try:
        value.encode("cp1252")
    except UnicodeEncodeError as exc:
        raise PortableError(f"{context}: unsupported visible text U+{ord(value[exc.start]):04X}; "
                            "only Windows-1252 text is supported; no font fallback") from exc
    require("\t" not in value, f"{context}: tabs are unsupported; use explicit spaces")
    return value.replace("\r\n", "\n").replace("\r", "\n")


def safe_path(value, *, must_exist=False):
    raw = os.fspath(value)
    require("\x00" not in raw and not raw.startswith(("\\\\", "//"))
            and "://" not in raw, "Only regular local filesystem paths are supported")
    path = Path(raw)
    require(".." not in path.parts, "Path traversal is forbidden")
    for part in path.parts[1:] if path.anchor else path.parts:
        require(":" not in part and not part.endswith((".", " "))
                and not re.match(r"(?i)^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\.|$)", part),
                f"Unsafe path segment: {part!r}")
    path = path.absolute()
    for cursor in reversed([path, *path.parents]):
        if cursor.exists() or cursor.is_symlink():
            info = cursor.lstat()
            require(not stat.S_ISLNK(info.st_mode), f"Symbolic links are forbidden: {cursor}")
            tag = getattr(info, "st_reparse_tag", 0)
            # Locally hydrated OneDrive CLOUD tags are not links. Never hydrate.
            require(not tag or tag & 0xFFFF0FFF == 0x9000001A,
                    f"Non-cloud reparse point/junction forbidden: {cursor}")
            attrs = getattr(info, "st_file_attributes", 0)
            require(not attrs & 0x441000, f"Input must be locally hydrated: {cursor}")
    if must_exist:
        require(path.is_file(), f"Input file unavailable: {path}")
    return path


def relative_parts(value):
    require(isinstance(value, str), "Relative resource path must be a string")
    normalized = value.replace("\\", "/")
    parts = normalized.rstrip("/").split("/")
    require(bool(normalized) and not normalized.startswith("/")
            and all(p and p not in (".", "..") and ":" not in p
                    and not p.endswith((".", " ")) for p in parts),
            f"Unsafe archive/catalog path: {value!r}")
    return parts


def read_bytes(path, limit):
    path = safe_path(path, must_exist=True)
    require(path.stat().st_size <= limit, f"File exceeds {limit} bytes: {path}")
    with path.open("rb") as stream:
        data = stream.read(limit + 1)
    require(len(data) <= limit, f"File grew beyond byte limit: {path}")
    return data


def _pairs(pairs):
    result = {}
    seen = set()
    for key, value in pairs:
        require(key.casefold() not in seen, f"Duplicate JSON key: {key}")
        seen.add(key.casefold())
        result[key] = value
    return result


def _json_depth(value, depth=0):
    require(depth <= 64, "JSON exceeds 64 nesting levels")
    if isinstance(value, dict):
        for item in value.values():
            _json_depth(item, depth + 1)
    elif isinstance(value, list):
        for item in value:
            _json_depth(item, depth + 1)
    elif isinstance(value, float):
        require(math.isfinite(value), "Non-finite JSON numbers are forbidden")


def load_json(path):
    try:
        value = json.loads(read_bytes(path, MAX_JSON).decode("utf-8-sig"), object_pairs_hook=_pairs)
    except (UnicodeDecodeError, json.JSONDecodeError, RecursionError) as exc:
        raise PortableError(f"Invalid bounded UTF-8 JSON: {exc}") from exc
    _json_depth(value)
    return value


def color(value, default="#475569"):
    if value is None:
        return default
    match = RGB.fullmatch(value) if isinstance(value, str) else None
    require(match is not None and all(int(c) <= 255 for c in match.groups()),
            f"Color must be RGB(r,g,b), with channels 0..255: {value!r}")
    return "#" + "".join(f"{int(c):02x}" for c in match.groups())


def card_style(node):
    return node.get("cardStyle", "icon" if node.get("iconRef") else "label")


def default_font_size(node):
    if node["kind"] == "container":
        return 11
    if node["kind"] == "note":
        return 12
    return 10 if card_style(node) in ("icon", "label") else 12


def edge_direction(edge):
    return edge.get("direction", "none" if edge["kind"] in ("peering", "dns", "governance", "association")
                    else "forward")


def edge_dashed(edge):
    return edge.get("dashed", edge["kind"] in ("peering", "dns", "telemetry", "dependency", "association"))


def caption(node):
    if node["kind"] == "card" and card_style(node) in ("icon", "label"):
        return node.get("displayLabel", node["label"].splitlines()[0])
    return node["label"]


def rect(node):
    return (node["x"] - node["width"] / 2, node["y"] - node["height"] / 2,
            node["x"] + node["width"] / 2, node["y"] + node["height"] / 2)


def _metadata_validation(item):
    for key in ("state", "purpose", "sourceRef", "sourceId", "component", "url", "layer",
                "details", "displayLabel"):
        if key in item:
            text(item[key], key, empty=True)
    if "requirementIds" in item:
        for entry in array(item["requirementIds"], "requirementIds"):
            text(entry, "requirement id")
    if "confidence" in item:
        number(item["confidence"], "confidence", 0, 1)
    for key in ("mainNodeIds", "mainEdgeIds"):
        if key in item:
            values = array(item[key], key)
            for entry in values:
                identifier(entry, key)
            require(len(set(values)) == len(values), f"{key}: duplicate mapping")
    if "font" in item:
        require(item["font"] in ("Arial", "Helvetica"), "Unsupported font; use Arial or Helvetica")


def _validate_structure(model):
    """Validate the architecture model without following any external references."""
    _json_depth(model)
    object_fields(model, ("schemaVersion", "outputContract", "presentationProfile", "hardening", "pages"),
                  ("title", "conversionMode", "referenceContract", "scope", "description",
                   "requirements", "referenceNotes", "assumptions"), "model")
    require(type(model["schemaVersion"]) is int and model["schemaVersion"] == 1,
            "schemaVersion must be integer 1")
    require(model["outputContract"] == CONTRACT, f"outputContract must be {CONTRACT}")
    require(model["presentationProfile"] in ("enterprise", "reference"), "Invalid presentationProfile")
    mode = model.get("conversionMode", "new-design")
    require(mode in ("new-design", "faithful", "reference-plus-proposal"),
            "Unsupported source conversion mode")
    if "referenceContract" in model:
        text(model["referenceContract"], "referenceContract")
    for key in ("title", "description"):
        if key in model:
            text(model[key], key, empty=True)
    # Provenance is retained in the model part, never followed as a URL or executed.
    for key in ("requirements", "referenceNotes"):
        if key in model:
            for item in array(model[key], key):
                require(isinstance(item, dict) and item, f"{key} entries must be nonempty metadata objects")
                for name, value in item.items():
                    text(name, f"{key} field name")
                    text(value, f"{key}.{name}", empty=True)
    if "assumptions" in model:
        for item in array(model["assumptions"], "assumptions"):
            text(item, "assumption")
    hardening = model["hardening"]
    object_fields(hardening, ("status", "reason"), context="hardening")
    require(hardening["status"] in ("proposed", "already-enterprise", "not-applicable"),
            "Invalid hardening.status")
    text(hardening["reason"], "hardening.reason")
    pages = array(model["pages"], "pages", 3)
    require(len(pages) == 3, "Exactly three pages are required")
    page_names = set()
    metadata = ("state", "purpose", "sourceRef", "sourceId", "component", "url", "layer",
                "requirementIds", "confidence", "details", "mainNodeIds", "mainEdgeIds")
    for index, page in enumerate(pages):
        object_fields(page, ("name", "view", "role", "width", "height", "nodes", "edges"),
                      ("title", "subtitle", "furniture", "footer", "titleFontSize", "scope"), "page")
        text(page["name"], "page.name")
        require(len(page["name"]) <= 100 and not re.search(r"[:\\/?*\[\]]", page["name"]),
                "Invalid native page name")
        require(page["name"].casefold() not in page_names, "Duplicate page name")
        page_names.add(page["name"].casefold())
        require(page["view"] == ("main", "hardening", "flowchart")[index], "Invalid page view/order")
        role = "diagram" if index == 0 or index == 1 and hardening["status"] == "proposed" else "notes"
        require(page["role"] == role, f"Page {index + 1} role must be {role}")
        for key in ("width", "height"):
            number(page[key], f"page.{key}", 1, 200)
        for key in ("title", "subtitle"):
            visible_text(page.get(key, ""), f"page.{key}")
        require(type(page.get("furniture", True)) is bool, "furniture must be boolean")
        footer = page.get("footer")
        require(footer is None or footer is False or isinstance(footer, str), "Invalid footer")
        if isinstance(footer, str):
            visible_text(footer, "page.footer")
        number(page.get("titleFontSize", 23), "titleFontSize", 8, 72)
        nodes, edges = array(page["nodes"], "nodes", 500), array(page["edges"], "edges", 1000)
        ids, node_map = set(), {}
        for node in nodes:
            object_fields(node, ("id", "kind", "label", "x", "y", "width", "height"),
                          (*metadata, "parent", "cardStyle", "iconRef", "displayLabel",
                           "containerStyle", "boundaryType", "fontSize", "font", "fill",
                           "color", "iconSize", "linePattern", "scope"), "node")
            identifier(node["id"])
            require(node["id"] not in ids, f"Duplicate shape id: {node['id']}")
            ids.add(node["id"])
            node_map[node["id"]] = node
            require(node["kind"] in ("card", "container", "note"), "Invalid node kind")
            text(node["label"], "node.label")
            _metadata_validation(node)
            require(not any(key in node for key in ("mainNodeIds", "mainEdgeIds"))
                    or index == 2 and node["kind"] == "card",
                    "MAIN coverage metadata is only valid on flowchart cards")
            for key in ("x", "y"):
                number(node[key], key)
            for key in ("width", "height"):
                number(node[key], key, 0.05)
            number(node.get("fontSize", default_font_size(node)), "fontSize", 6, 72)
            color(node.get("fill"))
            color(node.get("color"))
            if "linePattern" in node:
                require(type(node["linePattern"]) is int and node["linePattern"] in (0, 1, 2),
                        "linePattern must be 0, 1, or 2")
            left, bottom, right, top = rect(node)
            require(left >= -1e-8 and bottom >= -1e-8 and right <= page["width"] + 1e-8
                    and top <= page["height"] + 1e-8, f"Out-of-bounds node: {node['id']}")
            if "parent" in node and node["parent"] not in (None, ""):
                identifier(node["parent"], "parent")
            if node["kind"] == "container":
                require(node.get("containerStyle") == "boundary"
                        and isinstance(node.get("boundaryType"), str)
                        and node.get("boundaryType") in BOUNDARIES,
                        "Containers require containerStyle boundary and a meaningful boundaryType")
            else:
                require("containerStyle" not in node and "boundaryType" not in node,
                        "Boundary fields apply only to containers")
            if node["kind"] == "card":
                require(card_style(node) in ("icon", "label", "standard", "detail"), "Invalid cardStyle")
                if card_style(node) == "icon":
                    require(bool(node.get("iconRef")), f"Missing iconRef on {node['id']}")
                if "iconSize" in node:
                    number(node["iconSize"], "iconSize", 0.35, 4)
                if node.get("iconRef"):
                    identifier(node["iconRef"], "iconRef")
            else:
                require(not any(k in node for k in ("cardStyle", "iconSize", "iconRef")),
                        "Card fields apply only to cards")
            visible_text(caption(node), f"caption {node['id']}")
        for node in nodes:
            seen, child = {node["id"]}, node
            while child.get("parent"):
                parent_id = child["parent"]
                require(parent_id in node_map, f"Unresolved parent: {parent_id}")
                require(parent_id not in seen, "Cyclic containment")
                seen.add(parent_id)
                parent = node_map[parent_id]
                require(parent["kind"] == "container", "Parent must be a container")
                a, b = rect(child), rect(parent)
                require(a[0] >= b[0] - 1e-8 and a[1] >= b[1] - 1e-8
                        and a[2] <= b[2] + 1e-8 and a[3] <= b[3] + 1e-8,
                        f"Child {child['id']} lies outside its parent")
                child = parent
        for edge in edges:
            object_fields(edge, ("id", "source", "target", "kind"),
                          (*metadata, "label", "direction", "dashed", "sourceSide", "targetSide",
                           "sourcePosition", "targetPosition", "routeStyle", "points", "color"),
                          "edge")
            identifier(edge["id"])
            require(edge["id"] not in ids, f"Duplicate shape id: {edge['id']}")
            ids.add(edge["id"])
            _metadata_validation(edge)
            require(not any(key in edge for key in ("mainNodeIds", "mainEdgeIds")),
                    "MAIN coverage metadata is only valid on flowchart cards")
            for endpoint in ("source", "target"):
                identifier(edge[endpoint], endpoint)
                require(edge[endpoint] in node_map, f"Unresolved {endpoint}: {edge[endpoint]}")
                if endpoint + "Side" in edge:
                    require(isinstance(edge[endpoint + "Side"], str)
                            and edge[endpoint + "Side"] in SIDES, "Invalid anchor side")
                number(edge.get(endpoint + "Position", 0.5), endpoint + "Position", 0, 1)
            require(edge["source"] != edge["target"], "Self-loop needs explicit separate ports; unsupported")
            require(isinstance(edge["kind"], str) and edge["kind"] in KINDS, "Invalid relationship kind")
            require(edge.get("direction", "forward") in ("forward", "backward", "both", "none"),
                    "Invalid direction")
            require(type(edge.get("dashed", False)) is bool, "dashed must be boolean")
            require(edge.get("routeStyle", "orthogonal") in ("straight", "orthogonal"), "Invalid routeStyle")
            color(edge.get("color"))
            visible_text(edge.get("label", ""), f"edge label {edge['id']}")
            if "points" in edge:
                points = array(edge["points"], "edge.points", 50)
                require(len(points) >= 2, "Explicit route needs at least two points")
                for point in points:
                    object_fields(point, ("x", "y"), context="point")
                    number(point["x"], "point.x", 0, page["width"])
                    number(point["y"], "point.y", 0, page["height"])
        if role == "diagram":
            cards = [n for n in nodes if n["kind"] == "card"]
            require(len(cards) >= 2, "Diagram needs at least two architectural entities")
            card_ids = {node["id"] for node in cards}
            require(any(edge["source"] in card_ids and edge["target"] in card_ids for edge in edges),
                    "Diagram needs an attached relationship between distinct card entities")
            require(sum(card_style(n) == "icon" for n in cards) / len(cards) >= 0.6,
                    "Diagram requires at least 60% icon-led cards; no fallback icons")
            require(sum(card_style(n) in ("standard", "detail") for n in cards) / len(cards) <= 0.2,
                    "Diagram permits at most 20% boxed cards")
            require(sum(n["kind"] == "note" for n in nodes) <= 3, "At most three diagram callouts")
            for node in nodes:
                cap = caption(node)
                limit = {"card": 8, "note": 18, "container": 10}[node["kind"]]
                require(words(cap) <= limit, f"Caption too long on {node['id']}: maximum {limit} words")
                if node["kind"] == "card":
                    require(len(cap) <= 80 and 1 <= len(cap.splitlines()) <= 3,
                            "Entity captions need 1-3 lines and at most 80 characters")
                    require(not re.search(r"(?m)^\s*(?:[-*\u2022]|\d+[.)])\s+", cap),
                            "Implementation checklists belong on notes pages")
            require(all(words(e.get("label", "")) <= 6 for e in edges), "Edge caption exceeds six words")
            density = sum(words(caption(n)) for n in nodes) + sum(words(e.get("label", "")) for e in edges)
            density += sum(words(page.get(k, "") or "") for k in ("title", "subtitle", "footer"))
            require(density <= 300, "Diagram exceeds 300 visible words")
    main, hard_page, flow = pages
    main_nodes = {n["id"] for n in main["nodes"]}
    required_main_nodes = {n["id"] for n in main["nodes"] if n["kind"] == "card"}
    main_edges = {e["id"] for e in main["edges"]}
    flow_cards = [n for n in flow["nodes"] if n["kind"] == "card"]
    require(len(flow_cards) >= 2 and all(n.get("cardStyle") in ("detail", "standard") and words(n["label"]) >= 10
                                       for n in flow_cards),
            "Flowchart needs >=2 explicit detail/standard cards with >=10 visible words each")
    covered_nodes, covered_edges = set(), set()
    for node in flow["nodes"]:
        for key, expected, covered in (("mainNodeIds", main_nodes, covered_nodes),
                                       ("mainEdgeIds", main_edges, covered_edges)):
            values = node.get(key, [])
            require(set(values) <= expected, f"{key} must refer only to page-one nodes/edges")
            covered.update(values)
        if node["kind"] == "card":
            require("mainNodeIds" in node and "mainEdgeIds" in node,
                    "Flowchart detail cards require mainNodeIds and mainEdgeIds arrays")
            require(node["mainNodeIds"] or node["mainEdgeIds"],
                    "Each flowchart card must reference page-one nodes or edges")
    require(required_main_nodes <= covered_nodes and covered_edges == main_edges,
            "Flowchart mapping must cover every page-one card and edge")
    adjacent = {n["id"]: set() for n in flow_cards}
    for edge in flow["edges"]:
        require(edge["source"] in adjacent and edge["target"] in adjacent
                and edge_direction(edge) in ("forward", "backward", "both"),
                "Flowchart edges must be directed between distinct flowchart cards")
        adjacent[edge["source"]].add(edge["target"])
        adjacent[edge["target"]].add(edge["source"])
    reached, pending = set(), [flow_cards[0]["id"]]
    while pending:
        current = pending.pop()
        if current not in reached:
            reached.add(current)
            pending.extend(adjacent[current] - reached)
    require(reached == set(adjacent), "All flowchart detail cards must be connected")
    require(any(n["kind"] == "note" and words(n["label"]) >= 40 for n in flow["nodes"]),
            "Flowchart requires a visible narrative note of at least 40 words")
    if hardening["status"] != "proposed":
        require(all(n["kind"] == "note" for n in hard_page["nodes"]) and not hard_page["edges"],
                "Non-proposed hardening must use notes only, without cards or edges")
        require(any(n["kind"] == "note" and words(n["label"]) >= 40 for n in hard_page["nodes"]),
                "Non-proposal hardening page needs a visible note of at least 40 words")
    else:
        def semantics(page):
            lookup = {n["id"]: (n["kind"], n["label"], n.get("details", ""),
                                n.get("purpose", n["label"]), n.get("boundaryType", ""))
                      for n in page["nodes"] if n["kind"] != "note"}
            relations = [("containment", lookup[n["parent"]], lookup[n["id"]])
                         for n in page["nodes"] if n["id"] in lookup and n.get("parent") in lookup]
            relations.extend((lookup[e["source"]], lookup[e["target"]], e["kind"],
                              e.get("label", ""), e.get("details", ""), edge_direction(e))
                             for e in page["edges"] if e["source"] in lookup and e["target"] in lookup)
            return (sorted(json.dumps(value, ensure_ascii=False) for value in lookup.values()),
                    sorted(json.dumps(value, ensure_ascii=False) for value in relations))
        require(semantics(main) != semantics(hard_page),
                "Proposed hardening cannot duplicate main with renamed IDs/repositioning")
        if "scope" in main or "scope" in hard_page:
            require(main.get("scope") == hard_page.get("scope"), "Hardening scope must remain unchanged")
    return {"valid": True, "outputContract": CONTRACT, "pages": 3,
            "layoutChecked": False, "comUsed": False}


def requires_reference(model):
    return model.get("conversionMode") in ("faithful", "reference-plus-proposal") or "referenceContract" in model


def validate_model(model, model_path=None, contract_path=None):
    """Validate structure and, when required, the authoritative source contract."""
    result = _validate_structure(model)
    if requires_reference(model):
        require(model_path is not None,
                "Reference conversion requires model_path for authoritative source-contract verification")
        path = safe_path(model_path, must_exist=True)
        try:
            from portable_reference import validate_reference
        except ImportError as exc:
            raise PortableError("Source-contract verifier unavailable: deploy portable_reference.py "
                                "alongside portable_visio.py; conversion was not performed") from exc
        try:
            report = validate_reference(model, path, contract_path)
        except ValueError as exc:
            raise PortableError(f"Reference fidelity validation failed: {exc}") from exc
        require(isinstance(report, dict) and report.get("valid") is True,
                "Source-contract verifier did not return valid:true; conversion was not performed")
        result["sourceVerification"] = "verified-authoritative-contract"
        result["referenceValidation"] = report
    else:
        require(contract_path is None, "A source contract override requires a reference conversion model")
        result["sourceVerification"] = "not-applicable-new-design"
    return result


SVG_ELEMENTS = {"svg", "g", "defs", "title", "desc", "metadata", "path", "rect", "circle",
                "ellipse", "line", "polyline", "polygon", "linearGradient", "radialGradient",
                "stop", "clipPath", "mask", "use", "style", "symbol"}


def safe_svg(data):
    require(len(data) <= MAX_SVG, "SVG exceeds 4 MiB")
    try:
        source = data.decode("utf-8-sig")
    except UnicodeDecodeError as exc:
        raise PortableError("SVG must be UTF-8") from exc
    require(not re.search(r"<!DOCTYPE|<!ENTITY|<\?(?!xml\s)", source, re.I),
            "Unsafe SVG: declarations/entities/processing instructions")
    try:
        root = ET.fromstring(source)
    except ET.ParseError as exc:
        raise PortableError(f"Invalid SVG XML: {exc}") from exc
    require(root.tag == "{http://www.w3.org/2000/svg}svg", "SVG root/namespace is required")
    ids, refs = set(), []
    count = 0
    for element in root.iter():
        count += 1
        require(count <= 20000, "SVG element limit exceeded")
        local = element.tag.rsplit("}", 1)[-1]
        require(element.tag.startswith("{http://www.w3.org/2000/svg}")
                and local in SVG_ELEMENTS, f"Unsupported/unsafe SVG element: {local}")
        if "id" in element.attrib:
            require(element.attrib["id"] not in ids, "Duplicate SVG id")
            ids.add(element.attrib["id"])
        for key, value in element.attrib.items():
            attr = key.rsplit("}", 1)[-1]
            require(not attr.lower().startswith("on") and attr not in ("base", "src"),
                    "Unsafe SVG event/resource attribute")
            if attr in ("href",):
                require(value.startswith("#") and len(value) > 1, "External SVG reference forbidden")
                refs.append(value[1:])
    # Check decoded XML values too: entities must not hide a CSS resource reference.
    css = " ".join(value for element in root.iter()
                   for value in [*element.attrib.values(), element.text or ""])
    require("\\" not in css and not re.search(r"@|expression\s*\(|javascript:|data:|/\*", css, re.I),
            "Unsupported/unsafe SVG CSS")
    for match in re.finditer(r"url\s*\(([^)]*)\)", css, re.I):
        reference = match.group(1).strip().strip("'\"")
        require(re.fullmatch(r"#[A-Za-z0-9_.:-]+", reference) is not None,
                "External/unsupported SVG CSS resource forbidden")
        refs.append(reference[1:])
    require(not re.search(r"url\s*\(", re.sub(r"url\s*\([^)]*\)", "", css, flags=re.I), re.I),
            "Malformed SVG CSS resource reference")
    require(all(ref in ids for ref in refs), "Unresolved local SVG reference")
    # Recursive <use> graphs and filters are deliberately not supported.
    for element in root.iter():
        if element.tag.endswith("}use"):
            ref = element.get("href", element.get("{http://www.w3.org/1999/xlink}href"))
            require(ref is not None, "SVG use without a reference")
            target = next(e for e in root.iter() if e.get("id") == ref[1:])
            require(not any(e.tag.endswith("}use") for e in target.iter()),
                    "Nested/recursive SVG use is unsupported")
    viewbox = root.get("viewBox")
    if viewbox:
        try:
            numbers = [float(x) for x in re.split(r"[,\s]+", viewbox.strip())]
        except ValueError as exc:
            raise PortableError("Invalid SVG viewBox") from exc
        require(len(numbers) == 4 and all(math.isfinite(n) for n in numbers)
                and numbers[2] > 0 and numbers[3] > 0, "Invalid SVG viewBox")
        aspect = numbers[2] / numbers[3]
    else:
        def dimension(name):
            match = re.fullmatch(r"(\d+(?:\.\d+)?)(?:px)?", root.get(name, ""))
            require(match is not None and float(match.group(1)) > 0, "SVG needs dimensions or viewBox")
            return float(match.group(1))
        aspect = dimension("width") / dimension("height")
    require(0.02 <= aspect <= 50, "SVG aspect ratio exceeds safe limits")
    return source, aspect


def catalog_id(relative, collection="azure"):
    canonical = "\\".join(relative_parts(relative)).lower()
    return collection + "-" + hashlib.sha256(canonical.encode("utf-8")).hexdigest()[:12]


class IconCatalog:
    def __init__(self, directory):
        self.directory = safe_path(directory)
        records = load_json(self.directory / "catalog.json")
        array(records, "icon catalog", 10000)
        self.entries, self.cache = {}, {}
        for entry in records:
            require(isinstance(entry, dict), "Invalid catalog entry")
            key = identifier(entry.get("id"), "catalog id")
            require(key not in self.entries, f"Duplicate catalog id: {key}")
            self.entries[key] = entry

    def resolve(self, icon_ref):
        if icon_ref in self.cache:
            return self.cache[icon_ref]
        require(icon_ref in self.entries, f"Unresolved iconRef: {icon_ref}")
        entry = self.entries[icon_ref]
        require(entry.get("usable") is True, f"Catalog icon is marked unusable: {icon_ref}")
        parts = relative_parts(entry.get("path"))
        path = safe_path(self.directory.joinpath(*parts), must_exist=True)
        require(path.is_relative_to(self.directory), "Catalog path escaped library")
        require(path.suffix.lower() == ".svg", "Only vetted safe SVG catalog resources are supported")
        data = read_bytes(path, MAX_SVG)
        require(isinstance(entry.get("sha256"), str)
                and hashlib.sha256(data).hexdigest() == entry["sha256"].lower(),
                f"Icon SHA256 mismatch: {icon_ref}")
        source, aspect = safe_svg(data)
        self.cache[icon_ref] = {"svg": source, "aspect": aspect, "sha256": entry["sha256"],
                               "path": entry["path"]}
        return self.cache[icon_ref]


def validate_resources(model, catalog):
    refs = {n["iconRef"] for p in model["pages"] for n in p["nodes"] if n.get("iconRef")}
    require(not refs or catalog is not None, "An icon directory/catalog is required")
    for ref in sorted(refs):
        catalog.resolve(ref)
    return len(refs)


def capabilities():
    modules = {"reportlab": "reportlab", "resvg-py": "resvg_py", "Pillow": "PIL"}
    available = {name: importlib.util.find_spec(module) is not None for name, module in modules.items()}
    return {"python": sys.version.split()[0], "platform": sys.platform,
            "dependencies": available, "renderAvailable": all(available.values()),
            "comRequired": False, "networkRequiredWithLocalCatalog": False,
            "sourceVerifierAvailable": importlib.util.find_spec("portable_reference") is not None,
            "note": "Reports this process only; it does not establish Cowork tool permissions."}


def dependencies():
    missing = [name for name, present in capabilities()["dependencies"].items() if not present]
    require(not missing, "Unavailable rendering dependencies: " + ", ".join(missing)
            + ". If execution policy permits, install requirements-portable.txt; no handoff is performed.")
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfgen.canvas import Canvas
    from reportlab.lib.utils import ImageReader
    import resvg_py
    return pdfmetrics, Canvas, ImageReader, resvg_py


@dataclass
class TextBox:
    text: str
    x: float
    y: float
    width: float
    height: float
    size: float
    align: str = "left"
    color: str = "#1e293b"
    lines: list = field(default_factory=list)


def layout_text(value, x, y, width, height, size=10, align="left", context="text"):
    pdfmetrics, _, _, _ = dependencies()
    value = visible_text(value, context)
    lines = []
    available = (width - 0.08) * 72
    require(available > 0, f"{context}: text region too narrow")
    for paragraph in value.split("\n"):
        if not paragraph:
            lines.append("")
            continue
        line = ""
        for word in paragraph.split(" "):
            proposed = (line + " " + word) if line else word
            require(pdfmetrics.stringWidth(word, "Helvetica", size) <= available,
                    f"{context}: unbreakable word exceeds text width; enlarge geometry")
            if pdfmetrics.stringWidth(proposed, "Helvetica", size) <= available:
                line = proposed
            else:
                lines.append(line)
                line = word
        lines.append(line)
    require(len(lines) * size * 1.25 / 72 + 0.08 <= height + 1e-8,
            f"{context}: visible text would clip; increase height or reduce fontSize")
    return TextBox(value, x, y, width, height, size, align, lines=lines)


def node_layout(node, catalog):
    kind, style = node["kind"], card_style(node)
    w, h, x, y = (node[k] for k in ("width", "height", "x", "y"))
    size = node.get("fontSize", default_font_size(node))
    glyph = None
    if kind == "card" and style in ("icon", "label"):
        cap = caption(node)
        cap_h = len(cap.splitlines()) * size * 1.25 / 72 + 0.08
        box_y = y - h / 2 + cap_h / 2 if style == "icon" else y
        box = layout_text(cap, x, box_y, w, cap_h if style == "icon" else h,
                          size, "center", node["id"])
        require(len(box.lines) <= 3, f"{node['id']}: caption wrapping exceeds three lines")
        if style == "label":
            box.height = len(box.lines) * size * 1.25 / 72 + 0.08
        if style == "icon":
            available = min(w - 0.08, h - cap_h - 0.10)
            icon_size = node.get("iconSize", min(0.72, available))
            require(0.35 <= icon_size <= available + 1e-8,
                    f"{node['id']}: insufficient icon/caption space")
            art = catalog.resolve(node["iconRef"])
            gw, gh = (icon_size, icon_size / art["aspect"]) if art["aspect"] >= 1 else (
                icon_size * art["aspect"], icon_size)
            glyph = (x, y + h / 2 - 0.05 - icon_size / 2, gw, gh)
    elif kind == "container":
        box_h = size * 1.25 / 72 + 0.08
        box = layout_text(node["label"], x, y + h / 2 - box_h / 2 - 0.03,
                          w - 0.12, box_h, size, "left", node["id"])
    else:
        box_x, box_w = x, w - 0.16
        if node.get("iconRef"):
            art = catalog.resolve(node["iconRef"])
            extent = min(0.55, h * 0.65)
            gw, gh = (extent, extent / art["aspect"]) if art["aspect"] >= 1 else (
                extent * art["aspect"], extent)
            glyph = (x - w / 2 + 0.4, y + h / 2 - 0.44 if style == "detail" else y, gw, gh)
            box_x, box_w = x + 0.33, w - 0.85
        box = layout_text(node["label"], box_x, y, box_w, h if kind == "note" else h - 0.16, size,
                          "left" if kind == "note" or style == "detail" else "center", node["id"])
    if glyph:
        require(rect(node)[0] <= glyph[0] - glyph[2] / 2 and
                rect(node)[2] >= glyph[0] + glyph[2] / 2 and
                rect(node)[1] <= glyph[1] - glyph[3] / 2 and
                rect(node)[3] >= glyph[1] + glyph[3] / 2, "Glyph lies outside node bounds")
    return {"node": node, "text": box, "glyph": glyph}


def anchors(source, target, edge):
    dx = target["node"]["x"] - source["node"]["x"]
    dy = target["node"]["y"] - source["node"]["y"]
    a, b = (("right", "left") if dx >= 0 else ("left", "right")) if abs(dx) >= abs(dy) else (
        ("top", "bottom") if dy >= 0 else ("bottom", "top"))
    result = []
    for endpoint, layout, default in (("source", source, a), ("target", target, b)):
        side, position = edge.get(endpoint + "Side", default), edge.get(endpoint + "Position", 0.5)
        node = layout["node"]
        x, y, w, h = (node[k] for k in ("x", "y", "width", "height"))
        region = "body"
        if node["kind"] == "card" and card_style(node) in ("icon", "label"):
            if card_style(node) == "icon" and layout["glyph"] and side != "bottom":
                x, y, w, h = layout["glyph"]
                region = "glyph"
            else:
                left, bottom, right, top = text_rect(layout["text"])
                x, y, w, h = (left + right) / 2, (bottom + top) / 2, right - left, top - bottom
                region = "caption"
        if side in ("left", "right"):
            point = (x + (-w / 2 if side == "left" else w / 2), y - h / 2 + h * position)
        else:
            point = (x - w / 2 + w * position, y + (-h / 2 if side == "bottom" else h / 2))
        point = tuple(round(v, 10) for v in point)
        nx, ny = point[0] - node["x"] + node["width"] / 2, point[1] - node["y"] + node["height"] / 2
        result.append({"point": point, "side": side, "position": position,
                       "u": nx / node["width"], "v": ny / node["height"], "region": region})
    return result


def text_rect(box):
    pdfmetrics, _, _, _ = dependencies()
    width = max((pdfmetrics.stringWidth(line, "Helvetica", box.size) for line in box.lines), default=0) / 72 + 0.08
    height = len(box.lines) * box.size * 1.25 / 72 + 0.08
    left = box.x - width / 2 if box.align == "center" else box.x - box.width / 2
    return left, box.y + box.height / 2 - height, left + width, box.y + box.height / 2


def routing_obstacles(layouts):
    obstacles = []
    for node_id, info in layouts.items():
        node = info["node"]
        if node["kind"] == "container":
            obstacles.append((node_id, "caption", text_rect(info["text"])))
        elif node["kind"] == "card" and card_style(node) in ("icon", "label"):
            if info["glyph"]:
                x, y, w, h = info["glyph"]
                obstacles.append((node_id, "glyph", (x - w / 2, y - h / 2, x + w / 2, y + h / 2)))
            obstacles.append((node_id, "caption", text_rect(info["text"])))
        else:
            obstacles.append((node_id, "body", rect(node)))
    return obstacles


def inflate(bounds, amount=ROUTE_CLEARANCE):
    left, bottom, right, top = bounds
    return tuple(round(v, 10) for v in (left - amount, bottom - amount, right + amount, top + amount))


def segment_hits(a, b, bounds):
    left, bottom, right, top = bounds
    if abs(a[0] - b[0]) < EPSILON:
        return (left + EPSILON < a[0] < right - EPSILON
                and max(a[1], b[1]) > bottom + EPSILON and min(a[1], b[1]) < top - EPSILON)
    if abs(a[1] - b[1]) < EPSILON:
        return (bottom + EPSILON < a[1] < top - EPSILON
                and max(a[0], b[0]) > left + EPSILON and min(a[0], b[0]) < right - EPSILON)
    raise PortableError("Connector route has a diagonal segment")


def simplify_route(points):
    result = []
    for point in points:
        if result and point == result[-1]:
            continue
        while len(result) >= 2:
            a, b = result[-2:]
            if not ((a[0] == b[0] == point[0] and (b[1] - a[1]) * (point[1] - b[1]) >= 0)
                    or (a[1] == b[1] == point[1] and (b[0] - a[0]) * (point[0] - b[0]) >= 0)):
                break
            result.pop()
        result.append(point)
    return result


def port_departure(anchor, next_point):
    point = anchor["point"]
    vx, vy = PORT_DIRECTIONS[anchor["side"]]
    dx, dy = next_point[0] - point[0], next_point[1] - point[1]
    return abs(dx * vy - dy * vx) < EPSILON and dx * vx + dy * vy > EPSILON


def endpoint_exemptions(edge, start, end, layouts):
    result = []
    for endpoint, anchor in (("source", start), ("target", end)):
        node_id = edge[endpoint]
        region = "caption" if layouts[node_id]["node"]["kind"] == "container" else anchor["region"]
        result.append((node_id, region))
    return result


def route_is_clear(points, obstacles, exemptions):
    for index, (a, b) in enumerate(zip(points, points[1:])):
        for owner, region, bounds in obstacles:
            if ((index == 0 and (owner, region) == exemptions[0])
                    or (index == len(points) - 2 and (owner, region) == exemptions[1])):
                continue
            if segment_hits(a, b, inflate(bounds)):
                return False
    return True


def orthogonal_route(edge, start, end, obstacles, exemptions, bounds):
    a, b = start["point"], end["point"]
    if (port_departure(start, b) and port_departure(end, a)
            and route_is_clear([a, b], obstacles, exemptions)):
        return [a, b]
    stubs = [tuple(round(v + d * PORT_ESCAPE, 10) for v, d in zip(anchor["point"], PORT_DIRECTIONS[anchor["side"]]))
             for anchor in (start, end)]
    left, bottom, right, top = bounds
    inside = lambda point: left <= point[0] <= right and bottom <= point[1] <= top
    require(all(inside(point) for point in (a, b, *stubs)),
            f"{edge['id']}: no orthogonal route; selected port exits the page")
    for anchor, stub, exempt in zip((start, end), stubs, exemptions):
        require(not any(segment_hits(anchor["point"], stub, inflate(box))
                        for owner, region, box in obstacles if (owner, region) != exempt),
                f"{edge['id']}: no orthogonal route; selected port is blocked by a glyph or caption")
    padded = [inflate(box) for _, _, box in obstacles]
    xs = sorted({left, right, *(s[0] for s in stubs),
                 *(v for box in padded for v in (box[0], box[2]) if left <= v <= right)})
    ys = sorted({bottom, top, *(s[1] for s in stubs),
                 *(v for box in padded for v in (box[1], box[3]) if bottom <= v <= top)})
    require(len(xs) * len(ys) <= 250000,
            f"{edge['id']}: orthogonal routing grid exceeds safe limit; split the view")
    directions = ((-1, 0), (1, 0), (0, -1), (0, 1))
    initial = (xs.index(stubs[0][0]), ys.index(stubs[0][1]),
               directions.index(PORT_DIRECTIONS[start["side"]]))
    goal = (xs.index(stubs[1][0]), ys.index(stubs[1][1]))
    arrival = tuple(-v for v in PORT_DIRECTIONS[end["side"]])
    costs, previous, pending, clear_cache = {initial: 0}, {}, [(0, 0, initial)], {}
    final = None
    while pending:
        _, cost, current = heapq.heappop(pending)
        if cost != costs[current]:
            continue
        ix, iy, incoming = current
        if (ix, iy) == goal and directions[incoming] != tuple(-v for v in arrival):
            final = current
            break
        for direction, (sx, sy) in enumerate(directions):
            if directions[incoming] == (-sx, -sy):
                continue
            nx, ny = ix + sx, iy + sy
            if not (0 <= nx < len(xs) and 0 <= ny < len(ys)):
                continue
            key = tuple(sorted(((ix, iy), (nx, ny))))
            if key not in clear_cache:
                clear_cache[key] = not any(segment_hits((xs[ix], ys[iy]), (xs[nx], ys[ny]), box)
                                           for box in padded)
            if not clear_cache[key]:
                continue
            state = (nx, ny, direction)
            distance = abs(xs[nx] - xs[ix]) + abs(ys[ny] - ys[iy])
            updated = cost + distance + (0.25 if incoming != direction else 0)
            if updated < costs.get(state, math.inf):
                costs[state], previous[state] = updated, current
                estimate = abs(xs[nx] - stubs[1][0]) + abs(ys[ny] - stubs[1][1])
                heapq.heappush(pending, (updated + estimate, updated, state))
    require(final is not None, f"{edge['id']}: no obstacle-free orthogonal route; adjust layout or ports")
    path = []
    while final is not None:
        path.append((xs[final[0]], ys[final[1]]))
        final = previous.get(final)
    return simplify_route([a, *reversed(path), b])


def edge_layout(edge, layouts, bounds=None, extra_obstacles=(), label_obstacles=()):
    start, end = anchors(layouts[edge["source"]], layouts[edge["target"]], edge)
    a, b = start["point"], end["point"]
    require(math.dist(a, b) > 0.001, f"{edge['id']}: coincident anchors")
    obstacles = [*routing_obstacles(layouts), *extra_obstacles]
    exemptions = endpoint_exemptions(edge, start, end, layouts)
    if bounds is None:
        bounds = (0, 0, max(rect(n["node"])[2] for n in layouts.values()) + 1,
                  max(rect(n["node"])[3] for n in layouts.values()) + 1)
    if "points" in edge:
        points = [(round(p["x"], 10), round(p["y"], 10)) for p in edge["points"]]
        require(math.dist(points[0], a) < 1e-6 and math.dist(points[-1], b) < 1e-6,
                "Explicit route endpoints must equal semantic node anchors")
        points[0], points[-1] = a, b
        points = simplify_route(points)
    else:
        points = orthogonal_route(edge, start, end, obstacles, exemptions, bounds)
    require(len(points) >= 2 and all(p[0] == q[0] or p[1] == q[1] for p, q in zip(points, points[1:])),
            f"{edge['id']}: orthogonal route has a diagonal segment")
    require(port_departure(start, points[1]) and port_departure(end, points[-2]),
            f"{edge['id']}: route must leave and enter the selected visible ports from outside")
    require(route_is_clear(points, obstacles, exemptions),
            f"{edge['id']}: orthogonal route intersects a glyph or caption")
    require(all(bounds[0] <= p[0] <= bounds[2] and bounds[1] <= p[1] <= bounds[3] for p in points),
            f"{edge['id']}: route lies outside page")
    label = None
    if edge.get("label"):
        pdfmetrics, _, _, _ = dependencies()
        lw = max(pdfmetrics.stringWidth(line, "Helvetica", 10) for line in edge["label"].splitlines()) / 72 + 0.16
        lh = len(edge["label"].splitlines()) * 10 * 1.25 / 72 + 0.08
        for p, q in sorted(zip(points, points[1:]), key=lambda pair: math.dist(*pair), reverse=True):
            for position in (0.5, 0.25, 0.75):
                x, y = (p[i] + (q[i] - p[i]) * position for i in (0, 1))
                offsets = [(0, 0)]
                for gap in (0.06, ROUTE_CLEARANCE + 0.05, ROUTE_CLEARANCE + 0.15):
                    offsets.extend([(0, lh / 2 + gap), (0, -lh / 2 - gap)] if p[1] == q[1] else [
                        (lw / 2 + gap, 0), (-lw / 2 - gap, 0)])
                for ox, oy in offsets:
                    box = (x + ox - lw / 2, y + oy - lh / 2, x + ox + lw / 2, y + oy + lh / 2)
                    if (box[0] >= bounds[0] and box[1] >= bounds[1] and box[2] <= bounds[2] and box[3] <= bounds[3]
                            and not any(box[0] < r and box[2] > l and box[1] < t and box[3] > b
                                        for _, _, raw in [*obstacles, *label_obstacles]
                                        for l, b, r, t in [inflate(raw, 0.02)])):
                        label = layout_text(edge["label"], x + ox, y + oy, lw, lh, 10, "center", edge["id"])
                        break
                if label:
                    break
            if label:
                break
        require(label is not None, f"{edge['id']}: no clear connector-label position; adjust layout")
    return {"edge": edge, "start": start, "end": end, "points": points, "text": label}


def prepare(model, catalog):
    result, images = [], {}
    _, _, _, resvg = dependencies()
    for page in model["pages"]:
        nodes = {n["id"]: node_layout(n, catalog) for n in page["nodes"]}
        furniture = []
        if page.get("furniture", True):
            for role, value, y, height, size in (
                ("title", page.get("title", page["name"]), page["height"] - 0.4, 0.55,
                 page.get("titleFontSize", 23)),
                ("subtitle", page.get("subtitle", ""), page["height"] - 0.85, 0.3, 11),
                ("footer", page.get("footer", ""), 0.28, 0.3, 10)):
                if value:
                    furniture.append((role, layout_text(value, page["width"] / 2, y,
                                                       page["width"] - 0.8, height, size, context=role)))
        obstacles = [("furniture-" + role, "caption", text_rect(box)) for role, box in furniture]
        ports = []
        for edge in page["edges"]:
            for anchor in anchors(nodes[edge["source"]], nodes[edge["target"]], edge):
                point = anchor["point"]
                stub = tuple(v + d * PORT_ESCAPE for v, d in zip(point, PORT_DIRECTIONS[anchor["side"]]))
                ports.append(("port-" + edge["id"], "port",
                              inflate((min(point[0], stub[0]), min(point[1], stub[1]),
                                       max(point[0], stub[0]), max(point[1], stub[1])), ROUTE_CLEARANCE)))
        edges = []
        for edge in page["edges"]:
            info = edge_layout(edge, nodes, (0.02, 0.02, page["width"] - 0.02, page["height"] - 0.02),
                               obstacles, ports)
            edges.append(info)
            if info["text"]:
                obstacles.append(("edge-" + edge["id"], "caption", text_rect(info["text"])))
        for item in [e["text"] for e in edges if e["text"]] + [f[1] for f in furniture]:
            require(item.x - item.width / 2 >= 0 and item.y - item.height / 2 >= 0
                    and item.x + item.width / 2 <= page["width"]
                    and item.y + item.height / 2 <= page["height"], "Text/furniture lies outside page")
        for layout in nodes.values():
            if layout["glyph"]:
                ref = layout["node"]["iconRef"]
                art = catalog.resolve(ref)
                extent = max(layout["glyph"][2:])
                pixels = max(256, math.ceil(extent * 600))
                require(pixels <= 2400, "Icon exceeds raster rendering size limit")
                if ref not in images or images[ref]["pixels"] < pixels:
                    png = resvg.svg_to_bytes(svg_string=art["svg"], width=pixels if art["aspect"] >= 1 else None,
                                             height=pixels if art["aspect"] < 1 else None,
                                             skip_system_fonts=True)
                    require(png.startswith(b"\x89PNG\r\n\x1a\n"), f"SVG rendering failed: {ref}")
                    from PIL import Image
                    with Image.open(io.BytesIO(png)) as image:
                        require(image.width * image.height <= 2400 * 2400, "Rendered icon too large")
                        require(image.getbbox() is not None, f"SVG rendered empty: {ref}")
                        require(abs(image.width / image.height / art["aspect"] - 1) < 0.01,
                                f"SVG intrinsic viewport differs from viewBox aspect: {ref}")
                    images[ref] = {"data": png, "pixels": pixels}
        result.append({"page": page, "nodes": nodes, "edges": edges, "furniture": furniture})
    return result, images


def sub(parent, tag, **attrs):
    return ET.SubElement(parent, f"{{{V}}}{tag}", {k: str(v) for k, v in attrs.items()})


def cell(parent, name, value, formula=None, unit=None):
    attrs = {"N": name, "V": str(value)}
    if formula is not None:
        attrs["F"] = formula
    if unit:
        attrs["U"] = unit
    return sub(parent, "Cell", **attrs)


def properties(parent, values):
    section = sub(parent, "Section", N="Property")
    for index, (key, value) in enumerate(values.items()):
        if value is None:
            value = ""
        if not isinstance(value, str):
            value = json.dumps(value, ensure_ascii=False, separators=(",", ":"))
        row = sub(section, "Row", N=key, IX=index)
        cell(row, "Label", key, unit="STR")
        cell(row, "Type", 0)
        cell(row, "Value", value, '"' + value.replace('"', '""') + '"', "STR")
        cell(row, "Invisible", 0)


def xml_bytes(root):
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def shape(parent, sid, name, x, y, width, height, shape_type="Shape"):
    result = sub(parent, "Shape", ID=sid, NameU=name, Name=name, Type=shape_type,
                 LineStyle=0, FillStyle=0, TextStyle=0)
    for key, value, formula in (
        ("PinX", x, None), ("PinY", y, None), ("Width", width, None), ("Height", height, None),
        ("LocPinX", width / 2, "Width*0.5"), ("LocPinY", height / 2, "Height*0.5"), ("Angle", 0, None)):
        cell(result, key, value, formula)
    return result


def appearance(element, fill="#ffffff", stroke="#94a3b8", pattern=1, weight=0.75):
    for key, value in (("FillPattern", 1 if fill else 0), ("FillForegnd", fill or "#ffffff"),
                       ("LinePattern", pattern if stroke else 0), ("LineColor", stroke or "#ffffff"),
                       ("LineWeight", weight / 72), ("ShdwPattern", 0)):
        cell(element, key, value)


def geometry(element, width, height):
    section = sub(element, "Section", N="Geometry", IX=0)
    cell(section, "NoFill", 0)
    cell(section, "NoLine", 0)
    for index, (x, y, fx, fy) in enumerate(((0, 0, "0", "0"), (width, 0, "Width", "0"),
                                           (width, height, "Width", "Height"), (0, height, "0", "Height"),
                                           (0, 0, "0", "0")), 1):
        row = sub(section, "Row", T="MoveTo" if index == 1 else "LineTo", IX=index)
        cell(row, "X", x, fx)
        cell(row, "Y", y, fy)


def native_text(element, box, local_x=None, local_y=None, hide=False, *, horizontal_margin=0.04):
    cell(element, "TxtPinX", box.width / 2 if local_x is None else local_x)
    cell(element, "TxtPinY", box.height / 2 if local_y is None else local_y)
    cell(element, "TxtWidth", box.width)
    cell(element, "TxtHeight", box.height)
    cell(element, "TxtLocPinX", box.width / 2, "TxtWidth*0.5")
    cell(element, "TxtLocPinY", box.height / 2, "TxtHeight*0.5")
    cell(element, "TxtAngle", 0)
    for key in ("LeftMargin", "RightMargin", "TopMargin", "BottomMargin"):
        cell(element, key, horizontal_margin if key in ("LeftMargin", "RightMargin") else 0.04)
    cell(element, "VerticalAlign", 0)
    cell(element, "HideText", int(hide))
    section = sub(element, "Section", N="Character")
    row = sub(section, "Row", IX=0)
    cell(row, "Font", "Arial", 'FONT("Arial")')
    cell(row, "Size", box.size / 72, f"{box.size} pt")
    cell(row, "Color", box.color)
    cell(row, "Style", 0)
    section = sub(element, "Section", N="Paragraph")
    row = sub(section, "Row", IX=0)
    cell(row, "HorzAlign", 1 if box.align == "center" else 0)
    cell(row, "SpLine", -1.25)
    cell(row, "SpBefore", 0)
    cell(row, "SpAfter", 0)
    body = sub(element, "Text")
    sub(body, "cp", IX=0)
    pp = sub(body, "pp", IX=0)
    pp.tail = "\n".join(box.lines)


def node_properties(node, glyph=None):
    values = {"AvId": node["id"], "ComponentId": node.get("component", node["id"]),
              "Kind": node["kind"], "FullLabel": node["label"], "DisplayLabel": caption(node),
              "Details": node.get("details", ""), "CardStyle": card_style(node) if node["kind"] == "card" else "",
              "IconRef": node.get("iconRef", ""), "ParentId": node.get("parent") or "",
              "BaseFontSize": node.get("fontSize", default_font_size(node)),
              "ContainerStyle": node.get("containerStyle", ""),
              "BoundaryType": node.get("boundaryType", "")}
    if glyph:
        values.update(IconSize=max(glyph[2:]), IconAspect=glyph[2] / glyph[3])
    for source, target in (("mainNodeIds", "MainNodeIds"), ("mainEdgeIds", "MainEdgeIds")):
        if source in node:
            values[target] = node[source]
    for source, target in (("state", "State"), ("purpose", "Purpose"), ("sourceRef", "SourceRef"),
                           ("sourceId", "SourceIdRef"), ("requirementIds", "RequirementIds"),
                           ("confidence", "Confidence"), ("url", "Source"), ("scope", "Scope")):
        if source in node:
            values[target] = node[source]
    return values


def connector_geometry(points):
    """Keep each shared page-axis coordinate shared after endpoint glue moves."""
    if len(points) == 2:
        a, b = points
        if a[1] == b[1]:
            mid = (a[0] + b[0]) / 2
            points = [a, (mid, a[1]), (mid, b[1]), b]
            axes = (0, 1, 0)
        else:
            mid = (a[1] + b[1]) / 2
            points = [a, (a[0], mid), (b[0], mid), b]
            axes = (1, 0, 1)
    else:
        axes = [0 if a[1] == b[1] else 1 for a, b in zip(points, points[1:])]
    formulas = [[None, None] for _ in points]
    for axis, name in enumerate(("X", "Y")):
        groups = [[0]]
        for index, changing_axis in enumerate(axes, 1):
            if changing_axis == axis:
                groups.append([])
            groups[-1].append(index)
        delta = points[-1][axis] - points[0][axis]
        for group in groups:
            offset = points[group[0]][axis] - points[0][axis]
            if 0 in group:
                formula = "0"
            elif len(points) - 1 in group:
                formula = f"(End{name}-Begin{name})"
            elif abs(delta) > EPSILON:
                formula = f"{offset / delta:.12g}*(End{name}-Begin{name})"
            else:
                formula = f"{offset:.12g}"
            for index in group:
                formulas[index][axis] = formula
    return points, formulas


def native_page(layout, image_ids):
    root = ET.Element(f"{{{V}}}PageContents")
    shapes = sub(root, "Shapes")
    page = layout["page"]
    native_ids = {node["id"]: i + 1 for i, node in enumerate(page["nodes"])}
    next_id = len(native_ids) + len(page["edges"]) + 1
    node_elements, relations, connection_rows = {}, [], {key: [] for key in native_ids}
    # Connectors must be behind entities, while boundaries must be behind connectors.
    order = sorted(page["nodes"], key=lambda node: node["kind"] != "container")
    for node in order:
        info = layout["nodes"][node["id"]]
        x, y, w, h = (node[k] for k in ("x", "y", "width", "height"))
        group = info["glyph"] is not None or node["kind"] == "card" and card_style(node) in ("icon", "label")
        element = shape(shapes, native_ids[node["id"]], "av-" + node["id"], x, y, w, h,
                        "Group" if group else "Shape")
        node_elements[node["id"]] = element
        if group and card_style(node) in ("icon", "label"):
            fill, stroke = None, None
        elif node["kind"] == "container":
            fill, stroke = node.get("fill"), color(node.get("color"), "#0078d4")
            fill = color(fill) if fill else None
        else:
            fill, stroke = color(node.get("fill"), "#f1f5f9" if node["kind"] == "note" else "#ffffff"), (
                None if node["kind"] == "note" else color(node.get("color"), "#94a3b8"))
        appearance(element, fill, stroke, node.get("linePattern", 1))
        if group:
            cell(element, "SelectMode", 1)
            cell(element, "IsTextEditTarget", 0)
        properties(element, node_properties(node, info["glyph"]))
        geometry(element, w, h)
        box = info["text"]
        if group:
            # Full group text is hidden; editable caption children carry visible text.
            hidden = TextBox(node["label"], 0, 0, w, h, box.size, lines=[node["label"]])
            native_text(element, hidden, hide=True)
            children = sub(element, "Shapes")
            if info["glyph"]:
                gx, gy, gw, gh = info["glyph"]
                glyph = shape(children, next_id, "glyph-" + node["id"],
                              gx - x + w / 2, gy - y + h / 2, gw, gh, "Foreign")
                next_id += 1
                appearance(glyph, None, None)
                cell(glyph, "ImgOffsetX", 0)
                cell(glyph, "ImgOffsetY", 0)
                cell(glyph, "ImgWidth", gw, "Width")
                cell(glyph, "ImgHeight", gh, "Height")
                cell(glyph, "LockAspect", 1)
                properties(glyph, {"AvRole": "icon", "IconRef": node["iconRef"]})
                geometry(glyph, gw, gh)
                foreign = sub(glyph, "ForeignData", ForeignType="Bitmap", CompressionType="PNG")
                relation_id = "rId" + str(len(relations) + 1)
                sub(foreign, "Rel", **{f"{{{R}}}id": relation_id})
                relations.append((relation_id, R + "/image", "../media/" + image_ids[node["iconRef"]]))
            left, bottom, right, top = text_rect(box)
            caption_box = replace(box, x=(left + right) / 2, y=(bottom + top) / 2,
                                  width=right - left, height=top - bottom)
            child = shape(children, next_id, "caption-" + node["id"],
                          caption_box.x - x + w / 2, caption_box.y - y + h / 2,
                          caption_box.width, caption_box.height)
            next_id += 1
            appearance(child, None, None)
            properties(child, {"AvRole": "caption"})
            geometry(child, caption_box.width, caption_box.height)
            native_text(child, caption_box, horizontal_margin=0)
        else:
            native_text(element, box, box.x - x + w / 2, box.y - y + h / 2)
    connectors = []
    connects = sub(root, "Connects")
    for i, info in enumerate(layout["edges"]):
        edge = info["edge"]
        eid = len(native_ids) + i + 1
        a, b = info["points"][0], info["points"][-1]
        dx, dy = b[0] - a[0], b[1] - a[1]
        length, angle = math.hypot(dx, dy), math.atan2(dy, dx)
        element = shape(shapes, eid, "av-" + edge["id"], (a[0] + b[0]) / 2,
                        (a[1] + b[1]) / 2, length, 0)
        connectors.append(element)
        for key, value, formula in (("PinX", (a[0] + b[0]) / 2, "(BeginX+EndX)*0.5"),
                                    ("PinY", (a[1] + b[1]) / 2, "(BeginY+EndY)*0.5"),
                                    ("Width", length, "SQRT((EndX-BeginX)^2+(EndY-BeginY)^2)"),
                                    ("Angle", angle, "ATAN2(EndY-BeginY,EndX-BeginX)")):
            entry = element.find(f"{{{V}}}Cell[@N='{key}']")
            entry.set("V", str(value))
            entry.set("F", formula)
        for key, value in (("ObjType", 2), ("GlueType", 0), ("ConFixedCode", 0),
                           ("ShapeRouteStyle", 1), ("ConLineRouteExt", 1),
                           ("ConLineJumpCode", 1), ("Rounding", 0)):
            cell(element, key, value)
        for endpoint, prefix, part in (("source", "Begin", 9), ("target", "End", 12)):
            anchor = info["start"] if endpoint == "source" else info["end"]
            target = edge[endpoint]
            row_index = len(connection_rows[target])
            connection_rows[target].append(anchor)
            ref = f"Sheet.{native_ids[target]}!Connections."
            formula = f"PAR(PNT({ref}X{row_index + 1},{ref}Y{row_index + 1}))"
            for axis, value in zip(("X", "Y"), anchor["point"]):
                cell(element, prefix + axis, value, formula)
            cell(element, "BegTrigger" if endpoint == "source" else "EndTrigger", 2,
                 f"_XFTRIGGER(Sheet.{native_ids[target]}!EventXFMod)")
            sub(connects, "Connect", FromSheet=eid, FromCell=prefix + "X", FromPart=part,
                ToSheet=native_ids[target], ToCell=f"Connections.X{row_index + 1}", ToPart=100 + row_index)
        stroke = color(edge.get("color"), COLORS[edge["kind"]])
        appearance(element, None, stroke, 2 if edge_dashed(edge) else 1, 1.5)
        direction = edge_direction(edge)
        cell(element, "BeginArrow", 13 if direction in ("backward", "both") else 0)
        cell(element, "EndArrow", 13 if direction in ("forward", "both") else 0)
        cell(element, "BeginArrowSize", 1)
        cell(element, "EndArrowSize", 1)
        edge_props = {"AvId": edge["id"], "ComponentId": edge.get("component", edge["id"]),
                             "Kind": "edge", "Relationship": edge["kind"],
                             "FullLabel": edge.get("label", ""), "Details": edge.get("details", ""),
                             "SourceId": edge["source"], "TargetId": edge["target"],
                             "SourceSide": info["start"]["side"], "TargetSide": info["end"]["side"],
                             "SourcePosition": info["start"]["position"],
                             "TargetPosition": info["end"]["position"], "Direction": direction,
                             "Dashed": edge_dashed(edge),
                             "RouteStyle": "orthogonal"}
        for source, target in (("state", "State"), ("purpose", "Purpose"), ("sourceRef", "SourceRef"),
                               ("sourceId", "SourceIdRef"), ("requirementIds", "RequirementIds"),
                               ("confidence", "Confidence")):
            if source in edge:
                edge_props[target] = edge[source]
        properties(element, edge_props)
        section = sub(element, "Section", N="Geometry", IX=0)
        cell(section, "NoFill", 1)
        geometry_points, formulas = connector_geometry(info["points"])
        for index, (point, (gx_f, gy_f)) in enumerate(zip(geometry_points, formulas), 1):
            gx, gy = point[0] - a[0], point[1] - a[1]
            lx, ly = gx * math.cos(angle) + gy * math.sin(angle), gy * math.cos(angle) - gx * math.sin(angle)
            row = sub(section, "Row", T="MoveTo" if index == 1 else "LineTo", IX=index)
            if index == 1:
                fx, fy = "0", "0"
            elif index == len(geometry_points):
                fx, fy = "Width", "0"
            else:
                fx = f"({gx_f})*COS(Angle)+({gy_f})*SIN(Angle)"
                fy = f"({gy_f})*COS(Angle)-({gx_f})*SIN(Angle)"
            cell(row, "X", lx, fx)
            cell(row, "Y", ly, fy)
        if info["text"]:
            box = info["text"]
            tx, ty = box.x - a[0], box.y - a[1]
            native_text(element, box, tx * math.cos(angle) + ty * math.sin(angle),
                        ty * math.cos(angle) - tx * math.sin(angle))
            element.find(f"{{{V}}}Cell[@N='TxtAngle']").set("F", "-Angle")
            element.find(f"{{{V}}}Cell[@N='TxtAngle']").set("V", str(-angle))
            cell(element, "TextBkgnd", "#ffffff", "RGB(255,255,255)+1")
            cell(element, "TextBkgndTrans", 0)
    for node_id, rows in connection_rows.items():
        if rows:
            element = node_elements[node_id]
            section = ET.Element(f"{{{V}}}Section", {"N": "Connection"})
            # Cells/Sections precede Text/Shapes in the published ShapeSheet sequence.
            position = next((i for i, child in enumerate(element)
                             if child.tag in (f"{{{V}}}Text", f"{{{V}}}Shapes")), len(element))
            element.insert(position, section)
            for index, anchor in enumerate(rows):
                row = sub(section, "Row", IX=index)
                node = layout["nodes"][node_id]["node"]
                cell(row, "X", anchor["u"] * node["width"], f"Width*{anchor['u']:.12g}")
                cell(row, "Y", anchor["v"] * node["height"], f"Height*{anchor['v']:.12g}")
                vx, vy = PORT_DIRECTIONS[anchor["side"]]
                # Type 0 is an inward connection: Visio escapes opposite its vector.
                cell(row, "DirX", -vx)
                cell(row, "DirY", -vy)
                cell(row, "Type", 0)
    for element in connectors:
        shapes.remove(element)
    boundary_count = sum(n["kind"] == "container" for n in page["nodes"])
    for i, element in enumerate(connectors):
        shapes.insert(boundary_count + i, element)
    for role, box in layout["furniture"]:
        element = shape(shapes, next_id, "furniture-" + role, box.x, box.y, box.width, box.height)
        next_id += 1
        appearance(element, None, None)
        properties(element, {"AvFurniture": role})
        native_text(element, box)
    # ShapeSheet extends Sheet: Cell*, Section*, Shapes?, Text?, Data*, ForeignData?.
    priority = {"Cell": 0, "Section": 1, "Shapes": 2, "Text": 3, "ForeignData": 4}
    for element in root.iter(f"{{{V}}}Shape"):
        element[:] = sorted(element, key=lambda child: priority.get(child.tag.rsplit("}", 1)[-1], 5))
    return xml_bytes(root), relations


def relationships(items):
    root = ET.Element(f"{{{PKG}}}Relationships")
    for rid, kind, target in items:
        ET.SubElement(root, f"{{{PKG}}}Relationship", Id=rid, Type=kind, Target=target)
    return xml_bytes(root)


def make_vsdx(model, layouts, images):
    contents = {}
    document = ET.Element(f"{{{V}}}VisioDocument")
    settings = sub(document, "DocumentSettings", TopPage=0, DefaultTextStyle=0,
                   DefaultLineStyle=0, DefaultFillStyle=0, DefaultGuideStyle=0)
    sub(settings, "GlueSettings").text = "9"
    faces = sub(document, "FaceNames")
    sub(faces, "FaceName", NameU="Arial", UnicodeRanges="-1 -1 -1 -1",
        CharSets="0", Panos="020b0604020202020204", Flags=325)
    styles = sub(document, "StyleSheets")
    style = sub(styles, "StyleSheet", ID=0, NameU="No Style", Name="No Style")
    appearance(style)
    sheet = sub(document, "DocumentSheet", NameU="TheDoc", LineStyle=0, FillStyle=0, TextStyle=0)
    values = {"OutputContract": model["outputContract"], "Hardening": model["hardening"],
              "PresentationProfile": model["presentationProfile"], "SchemaVersion": 1,
              "ConversionMode": model.get("conversionMode", "new-design"),
              "PortableRenderer": "azure-visio-1.6", "ModelPart": "/customXml/item1.xml"}
    if "scope" in model:
        values["Scope"] = model["scope"]
    if "referenceContract" in model:
        values["ReferenceContract"] = model["referenceContract"]
    topology = json.dumps([{"page": p["name"], "nodes": [n["id"] for n in p["nodes"]],
                           "edges": [{"id": e["id"], "source": e["source"], "target": e["target"]}
                                     for e in p["edges"]]} for p in model["pages"]], separators=(",", ":"))
    chunks = [topology[i:i + 24000] for i in range(0, len(topology), 24000)]
    values["TopologyChunkCount"] = len(chunks)
    for index, chunk in enumerate(chunks):
        values["Topology" if len(chunks) == 1 else f"Topology{index + 1}"] = chunk
    properties(sheet, values)
    contents["visio/document.xml"] = xml_bytes(document)
    model_root = ET.Element("ArchitectureModel", {"schemaVersion": "1"})
    model_root.text = json.dumps(model, ensure_ascii=False, separators=(",", ":"), allow_nan=False)
    contents["customXml/item1.xml"] = xml_bytes(model_root)
    contents["_rels/.rels"] = relationships([("rId1", VR + "document", "visio/document.xml"),
                                            ("rId2", "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
                                             "docProps/core.xml")])
    contents["visio/_rels/document.xml.rels"] = relationships(
        [("rId1", VR + "pages", "pages/pages.xml"), ("rId2", R + "/customXml", "../customXml/item1.xml"),
         ("rId3", VR + "windows", "windows.xml")])
    windows = ET.Element(f"{{{V}}}Windows")
    sub(windows, "Window", ID=0, WindowType="Drawing", Page=0)
    contents["visio/windows.xml"] = xml_bytes(windows)
    core_ns = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
    core = ET.Element(f"{{{core_ns}}}coreProperties")
    ET.SubElement(core, "{http://purl.org/dc/elements/1.1/}title").text = model.get("title", CONTRACT)
    ET.SubElement(core, "{http://purl.org/dc/elements/1.1/}creator").text = "Azure Visio portable renderer"
    contents["docProps/core.xml"] = xml_bytes(core)
    pages = ET.Element(f"{{{V}}}Pages")
    page_rels = []
    image_ids = {ref: f"image{index + 1}.png" for index, ref in enumerate(sorted(images))}
    for ref, image in images.items():
        contents["visio/media/" + image_ids[ref]] = image["data"]
    for index, layout in enumerate(layouts, 1):
        spec = layout["page"]
        page = sub(pages, "Page", ID=index - 1, Name=spec["name"], NameU=spec["name"],
                   IsCustomName=1, IsCustomNameU=1)
        psheet = sub(page, "PageSheet", LineStyle=0, FillStyle=0, TextStyle=0)
        for key, value in (("PageWidth", spec["width"]), ("PageHeight", spec["height"]),
                           ("PageScale", 1), ("DrawingScale", 1), ("DrawingSizeType", 0),
                           ("DrawingScaleType", 0), ("DrawingResizeType", 0), ("InhibitSnap", 0),
                           ("RouteStyle", 1), ("LineRouteExt", 1), ("LineJumpCode", 0)):
            cell(psheet, key, value)
        properties(psheet, {"AvPageView": spec["view"], "AvPageRole": spec["role"],
                            "AvFurniture": str(spec.get("furniture", True)),
                            "AvTitle": spec.get("title", ""), "AvSubtitle": spec.get("subtitle", "")})
        sub(page, "Rel", **{f"{{{R}}}id": f"rId{index}"})
        page_rels.append((f"rId{index}", VR + "page", f"page{index}.xml"))
        data, rels = native_page(layout, image_ids)
        contents[f"visio/pages/page{index}.xml"] = data
        if rels:
            contents[f"visio/pages/_rels/page{index}.xml.rels"] = relationships(rels)
    contents["visio/pages/pages.xml"] = xml_bytes(pages)
    contents["visio/pages/_rels/pages.xml.rels"] = relationships(page_rels)
    types = ET.Element(f"{{{CT}}}Types")
    for extension, mime in (("rels", "application/vnd.openxmlformats-package.relationships+xml"),
                            ("xml", "application/xml"), ("png", "image/png")):
        ET.SubElement(types, f"{{{CT}}}Default", Extension=extension, ContentType=mime)
    overrides = {"visio/document.xml": "application/vnd.ms-visio.drawing.main+xml",
                 "visio/windows.xml": "application/vnd.ms-visio.windows+xml",
                 "visio/pages/pages.xml": "application/vnd.ms-visio.pages+xml",
                 "docProps/core.xml": "application/vnd.openxmlformats-package.core-properties+xml"}
    overrides.update({f"visio/pages/page{i}.xml": "application/vnd.ms-visio.page+xml" for i in range(1, 4)})
    for part, mime in overrides.items():
        ET.SubElement(types, f"{{{CT}}}Override", PartName="/" + part, ContentType=mime)
    contents["[Content_Types].xml"] = xml_bytes(types)
    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in contents.items():
            archive.writestr(name, data)
    return output.getvalue()


def pdf_text(canvas, box):
    canvas.setFillColor(box.color)
    canvas.setFont("Helvetica", box.size)
    baseline = (box.y + box.height / 2 - 0.04) * 72 - box.size
    for line in box.lines:
        if box.align == "center":
            canvas.drawCentredString(box.x * 72, baseline, line)
        else:
            canvas.drawString((box.x - box.width / 2 + 0.04) * 72, baseline, line)
        baseline -= box.size * 1.25


def make_pdf(model, layouts, images):
    _, Canvas, ImageReader, _ = dependencies()
    output = io.BytesIO()
    canvas = Canvas(output, pageCompression=1, invariant=1)
    canvas.setTitle(model.get("title", CONTRACT))
    canvas.setAuthor("Azure Visio portable renderer")

    def draw_node(info):
        node = info["node"]
        x, y, w, h = (node[k] for k in ("x", "y", "width", "height"))
        unboxed = node["kind"] == "card" and card_style(node) in ("icon", "label")
        if not unboxed:
            fill = color(node.get("fill"), "#f1f5f9" if node["kind"] == "note" else "#ffffff")
            canvas.setFillColor(fill)
            canvas.setStrokeColor(color(node.get("color"), "#0078d4" if node["kind"] == "container" else "#94a3b8"))
            canvas.setLineWidth(0.75)
            canvas.setDash([4, 3] if node.get("linePattern") == 2 else [])
            canvas.rect((x - w / 2) * 72, (y - h / 2) * 72, w * 72, h * 72,
                        fill=int(node["kind"] != "container" or bool(node.get("fill"))),
                        stroke=int(node["kind"] != "note" and node.get("linePattern", 1) != 0))
        if info["glyph"]:
            gx, gy, gw, gh = info["glyph"]
            canvas.drawImage(ImageReader(io.BytesIO(images[node["iconRef"]]["data"])),
                             (gx - gw / 2) * 72, (gy - gh / 2) * 72, gw * 72, gh * 72, mask="auto")
        pdf_text(canvas, info["text"])

    def arrow(tip, previous):
        angle = math.atan2(tip[1] - previous[1], tip[0] - previous[0])
        tip = (tip[0] * 72, tip[1] * 72)
        size, half = 7, 2.6
        path = canvas.beginPath()
        path.moveTo(*tip)
        for sign in (-1, 1):
            path.lineTo(tip[0] - size * math.cos(angle) + sign * half * math.sin(angle),
                        tip[1] - size * math.sin(angle) - sign * half * math.cos(angle))
        path.close()
        canvas.drawPath(path, fill=1, stroke=0)

    for layout in layouts:
        page = layout["page"]
        canvas.setPageSize((page["width"] * 72, page["height"] * 72))
        for node in page["nodes"]:
            if node["kind"] == "container":
                draw_node(layout["nodes"][node["id"]])
        for info in layout["edges"]:
            edge, points = info["edge"], info["points"]
            stroke = color(edge.get("color"), COLORS[edge["kind"]])
            canvas.setStrokeColor(stroke)
            canvas.setFillColor(stroke)
            canvas.setLineWidth(1.5)
            canvas.setDash([5, 3] if edge_dashed(edge) else [])
            path = canvas.beginPath()
            path.moveTo(points[0][0] * 72, points[0][1] * 72)
            for point in points[1:]:
                path.lineTo(point[0] * 72, point[1] * 72)
            canvas.drawPath(path)
            direction = edge_direction(edge)
            if direction in ("forward", "both"):
                arrow(points[-1], points[-2])
            if direction in ("backward", "both"):
                arrow(points[0], points[1])
            if info["text"]:
                box = info["text"]
                canvas.setFillColor("#ffffff")
                canvas.rect((box.x - box.width / 2) * 72, (box.y - box.height / 2) * 72,
                            box.width * 72, box.height * 72, fill=1, stroke=0)
                pdf_text(canvas, box)
        for node in page["nodes"]:
            if node["kind"] != "container":
                draw_node(layout["nodes"][node["id"]])
        for _, box in layout["furniture"]:
            pdf_text(canvas, box)
        canvas.showPage()
    canvas.save()
    return output.getvalue()


def inspect_package(data):
    """Static OPC, image, editable-shape, metadata, and glue consistency check."""
    source = io.BytesIO(data) if isinstance(data, bytes) else safe_path(data, must_exist=True)
    with zipfile.ZipFile(source) as archive:
        require(len(archive.infolist()) <= 10000, "Package part limit exceeded")
        require(sum(item.file_size for item in archive.infolist()) <= MAX_EXPANDED,
                "Expanded package exceeds limit")
        names = archive.namelist()
        require(len(set(names)) == len(names), "Duplicate OPC part")
        for name in names:
            relative_parts(name)
        require(archive.testzip() is None, "Corrupt OPC ZIP CRC")
        require("visio/windows.xml" in names, "Missing native Windows part")
        windows = ET.fromstring(archive.read("visio/windows.xml"))
        require(windows.tag == f"{{{V}}}Windows", "Invalid native Windows part")
        document_rels = ET.fromstring(archive.read("visio/_rels/document.xml.rels"))
        require(any(r.get("Type") == VR + "windows" and r.get("Target") == "windows.xml"
                    for r in document_rels), "Missing native Windows relationship")
        content_types = ET.fromstring(archive.read("[Content_Types].xml"))
        require(any(r.get("PartName") == "/visio/windows.xml"
                    and r.get("ContentType") == "application/vnd.ms-visio.windows+xml"
                    for r in content_types), "Missing native Windows content type")
        for name in names:
            if name.endswith(".rels"):
                root = ET.fromstring(archive.read(name))
                base = PurePosixPath(name).parent.parent
                if name == "_rels/.rels":
                    base = PurePosixPath(".")
                for relation in root:
                    require(relation.get("TargetMode", "Internal") == "Internal", "External OPC relation forbidden")
                    parts = list(base.parts)
                    for part in PurePosixPath(relation.attrib["Target"]).parts:
                        if part == "..":
                            require(bool(parts), "OPC relation escaped package")
                            parts.pop()
                        elif part != ".":
                            parts.append(part)
                    require("/".join(parts) in names, "Unresolved OPC relation")
        pages = ET.fromstring(archive.read("visio/pages/pages.xml"))
        require(len(pages) == 3, "Package must contain exactly three pages")
        summary = []
        for index, page in enumerate(pages, 1):
            root = ET.fromstring(archive.read(f"visio/pages/page{index}.xml"))
            shapes = {s.attrib["ID"]: s for s in root.iter(f"{{{V}}}Shape")}
            top = root.find(f"{{{V}}}Shapes")
            require(top is not None and len(shapes) > 0, "No editable shapes")
            connects = list(root.findall(f"{{{V}}}Connects/{{{V}}}Connect"))
            edge_ids = {s.attrib["ID"] for s in top if s.find(f"{{{V}}}Cell[@N='BeginX']") is not None}
            for sid in edge_ids:
                connector = shapes[sid]
                values = {c.get("N"): float(c.get("V")) for c in connector.findall(f"{{{V}}}Cell")
                          if c.get("N") in ("BeginX", "BeginY", "EndX", "EndY", "Angle", "Width", "LocPinX",
                                            "LocPinY", "PinX", "PinY", "ShapeRouteStyle")}
                require(values.get("ShapeRouteStyle") == 1, "Connector must use native right-angle routing")
                points = []
                for row in connector.findall(f"{{{V}}}Section[@N='Geometry']/{{{V}}}Row"):
                    require(row.get("T") in ("MoveTo", "LineTo"), "Unsupported connector geometry")
                    x = float(row.find(f"{{{V}}}Cell[@N='X']").get("V")) - values["LocPinX"]
                    y = float(row.find(f"{{{V}}}Cell[@N='Y']").get("V")) - values["LocPinY"]
                    angle = values["Angle"]
                    points.append((values["PinX"] + x * math.cos(angle) - y * math.sin(angle),
                                   values["PinY"] + x * math.sin(angle) + y * math.cos(angle)))
                require(len(points) >= 2 and all(abs(a[0] - b[0]) < 1e-6 or abs(a[1] - b[1]) < 1e-6
                                                for a, b in zip(points, points[1:])),
                        "Native connector geometry has a diagonal segment")
                require(math.dist(points[0], (values["BeginX"], values["BeginY"])) < 1e-6
                        and math.dist(points[-1], (values["EndX"], values["EndY"])) < 1e-6,
                        "Native connector geometry misses its endpoints")
                records = [c for c in connects if c.get("FromSheet") == sid]
                require(len(records) == 2 and {c.get("FromCell") for c in records} == {"BeginX", "EndX"},
                        "Connector must have native begin/end glue")
                for connection in records:
                    require(connection.get("ToSheet") in shapes, "Glue target is missing")
                    target = shapes[connection.get("ToSheet")]
                    row_index = int(connection.get("ToCell").split("X")[-1]) - 1
                    row = target.find(f"{{{V}}}Section[@N='Connection']/{{{V}}}Row[@IX='{row_index}']")
                    require(row is not None, "Glue target connection row is missing")
                    target_values = {c.get("N"): float(c.get("V")) for c in target.findall(f"{{{V}}}Cell")
                                     if c.get("N") in ("PinX", "PinY", "LocPinX", "LocPinY")}
                    prefix = connection.get("FromCell")[:-1]
                    for axis in ("X", "Y"):
                        port = float(row.find(f"{{{V}}}Cell[@N='{axis}']").get("V"))
                        expected = target_values["Pin" + axis] - target_values["LocPin" + axis] + port
                        require(abs(values[prefix + axis] - expected) < 1e-6,
                                "Connector endpoint misses its native glue port")
                    role = "SourceId" if prefix == "Begin" else "TargetId"
                    semantic = connector.find(f"{{{V}}}Section[@N='Property']/{{{V}}}Row[@N='{role}']/"
                                              f"{{{V}}}Cell[@N='Value']")
                    identity = target.find(f"{{{V}}}Section[@N='Property']/{{{V}}}Row[@N='AvId']/"
                                           f"{{{V}}}Cell[@N='Value']")
                    require(semantic is not None and identity is not None and semantic.get("V") == identity.get("V"),
                            "Connector glue targets the wrong semantic service")
            summary.append({"name": page.get("Name"), "topLevelShapes": len(top),
                            "nativeConnectors": len(edge_ids), "glueRecords": len(connects),
                            "foreignGlyphs": sum(s.get("Type") == "Foreign" for s in shapes.values())})
        stored = ET.fromstring(archive.read("customXml/item1.xml"))
        stored_model = json.loads(stored.text)
        _validate_structure(stored_model)
        return {"valid": True, "pages": summary, "comUsed": False,
                "sourceVerification": "not-revalidated-by-static-package-inspection" if requires_reference(stored_model)
                                      else "not-applicable-new-design",
                "nativeApplicationAcceptance": "not-tested-by-static-inspection"}


def render(model, icon_directory, output_directory, name="architecture", *, model_path=None,
           contract_path=None, bundle=False):
    validation = validate_model(model, model_path, contract_path)
    require(type(bundle) is bool, "bundle must be boolean")
    identifier(name, "output basename")
    require(not name.endswith("."), "Invalid output basename")
    destination = safe_path(output_directory)
    require(not destination.exists(), "Output directory must be NEW; overwrites are forbidden")
    require(destination.parent.is_dir(), "Output parent directory must already exist")
    catalog = IconCatalog(icon_directory) if icon_directory else None
    validate_resources(model, catalog)
    layouts, images = prepare(model, catalog)
    vsdx = make_vsdx(model, layouts, images)
    inspection = inspect_package(vsdx)
    pdf = make_pdf(model, layouts, images)
    require(pdf.startswith(b"%PDF-") and vsdx.startswith(b"PK"), "Output format generation failed")
    files = {name + ".vsdx": vsdx, name + ".pdf": pdf}
    if bundle:
        bundle_stream = io.BytesIO()
        with zipfile.ZipFile(bundle_stream, "w", zipfile.ZIP_DEFLATED) as archive:
            for filename, data in files.items():
                archive.writestr(filename, data)
        files[name + ".zip"] = bundle_stream.getvalue()
    destination.mkdir(exist_ok=False)
    written = []
    try:
        for filename, data in files.items():
            path = destination / filename
            with path.open("xb") as stream:
                written.append(path)
                stream.write(data)
                stream.flush()
                os.fsync(stream.fileno())
    except OSError:
        # Delete only files created by this invocation; never clean an existing folder.
        for path in written:
            path.unlink()
        destination.rmdir()
        raise
    return {"vsdx": str(destination / (name + ".vsdx")),
            "pdf": str(destination / (name + ".pdf")), "inspection": inspection,
            "bundle": str(destination / (name + ".zip")) if bundle else None,
            "pdfCreation": "Direct same-model rendering; not a desktop Visio export",
            "glyphFormat": "600dpi transparent PNG within editable native groups",
            "visibleText": "Arial native / metric-compatible Helvetica PDF; Windows-1252",
            "sourceVerification": validation["sourceVerification"],
            "referenceValidation": validation.get("referenceValidation")}


class _NoRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        raise PortableError("Official icon download redirected; refusing any unvetted URL")


def extract_official_archive(data, output_directory):
    require(len(data) <= MAX_ZIP, "Icon archive exceeds 100 MiB")
    destination = safe_path(output_directory)
    require(not destination.exists() and destination.parent.is_dir(), "Icon directory must be NEW")
    records, entries = [], []
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        require(len(archive.infolist()) <= 10000, "Icon archive entry limit exceeded")
        total, seen = 0, set()
        for info in archive.infolist():
            parts = relative_parts(info.filename)
            key = "/".join(parts).casefold()
            require(key not in seen, "Duplicate archive destination")
            seen.add(key)
            mode = info.external_attr >> 16
            require(not stat.S_ISLNK(mode) and not info.flag_bits & 1,
                    "Symlink/encrypted archive entry forbidden")
            require(info.is_dir() or stat.S_IFMT(mode) in (0, stat.S_IFREG),
                    "Non-regular archive entry forbidden")
            total += info.file_size
            require(total <= MAX_EXPANDED and info.file_size <= MAX_ZIP,
                    "Expanded archive size limit exceeded")
            require(info.file_size <= max(1024 * 1024, info.compress_size * 200),
                    "Excessive ZIP expansion ratio")
            target = safe_path(destination.joinpath("azure", *parts))
            if not info.is_dir():
                payload = archive.read(info)
                if target.suffix.lower() == ".svg":
                    # Unsupported icons remain catalogued but cannot be selected.
                    issue = ""
                    try:
                        safe_svg(payload)
                    except PortableError as exc:
                        issue = str(exc)
                    relative = "\\".join(("azure", *parts))
                    records.append({"id": catalog_id(relative), "path": relative, "collection": "azure",
                                    "name": target.stem, "sha256": hashlib.sha256(payload).hexdigest(),
                                    "sourcePage": ICON_TERMS, "downloadUrl": OFFICIAL_URL,
                                    "usable": not issue, "issue": issue})
                entries.append((target, payload))
        require(records, "Official package contains no SVG icons")
    destination.mkdir()
    for path, payload in entries:
        path.parent.mkdir(parents=True, exist_ok=True)
        with path.open("xb") as stream:
            stream.write(payload)
    with (destination / "catalog.json").open("x", encoding="utf-8") as stream:
        json.dump(sorted(records, key=lambda item: item["path"].lower()), stream, ensure_ascii=False, indent=2)
    return {"directory": str(destination), "icons": len(records),
            "usable": sum(r["usable"] for r in records), "source": OFFICIAL_URL}


def download_icons(output_directory, accept_icon_terms=False):
    require(accept_icon_terms, f"Read {ICON_TERMS} and pass --accept-icon-terms for permitted icon use")
    destination = safe_path(output_directory)
    require(not destination.exists(), "Icon output directory must be NEW")
    opener = urllib.request.build_opener(urllib.request.ProxyHandler({}), _NoRedirect())
    request = urllib.request.Request(OFFICIAL_URL, headers={"User-Agent": "AzureVisio-Portable/1.6"})
    with opener.open(request, timeout=90) as response:
        require(response.geturl() == OFFICIAL_URL, "Unvetted icon download location")
        content_length = response.headers.get("Content-Length")
        require(content_length is None or int(content_length) <= MAX_ZIP, "Icon download exceeds byte limit")
        data = response.read(MAX_ZIP + 1)
    return extract_official_archive(data, destination)


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    commands = parser.add_subparsers(dest="command", required=True)
    commands.add_parser("capabilities", help="Report actual Python/dependency availability")
    for command in ("validate", "render"):
        p = commands.add_parser(command)
        p.add_argument("--model", required=True)
        p.add_argument("--icon-directory", required=True)
        if command == "render":
            p.add_argument("--output-directory", required=True, help="Must not already exist")
            p.add_argument("--name", default="architecture")
            p.add_argument("--bundle", action="store_true",
                           help="Also create a ZIP with the actual VSDX and directly rendered PDF; "
                                "does not replace either file or grant host artifact permissions")
    p = commands.add_parser("inspect")
    p.add_argument("--vsdx", required=True)
    p = commands.add_parser("download-icons")
    p.add_argument("--output-directory", required=True)
    p.add_argument("--accept-icon-terms", action="store_true")
    args = parser.parse_args(argv)
    try:
        if args.command == "capabilities":
            result = capabilities()
        elif args.command == "inspect":
            result = inspect_package(args.vsdx)
        elif args.command == "download-icons":
            result = download_icons(args.output_directory, args.accept_icon_terms)
        else:
            model = load_json(args.model)
            if args.command == "validate":
                result = validate_model(model, args.model)
                result["resolvedIcons"] = validate_resources(model, IconCatalog(args.icon_directory))
            else:
                result = render(model, args.icon_directory, args.output_directory, args.name,
                                model_path=args.model, bundle=args.bundle)
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0
    except (PortableError, OSError, zipfile.BadZipFile, ET.ParseError, urllib.error.URLError) as exc:
        print(json.dumps({"valid": False, "error": str(exc), "comUsed": False}), file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
