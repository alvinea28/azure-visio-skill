"""Filesystem-only source fidelity gate; no extraction, Office, COM, or networking."""

from __future__ import annotations

import hashlib
import json
import math
import os
from pathlib import Path
import re
import stat
from typing import Any


JSON_LIMIT = 10 * 1024 * 1024
SOURCE_LIMIT = 100 * 1024 * 1024
DEPTH_LIMIT = 64
RECORD_LIMIT = 2000
_DIRECTIONS = {"forward", "backward", "both", "none"}
_KINDS = {"card", "container", "note"}
_DEVICE = re.compile(r"^(CON|PRN|AUX|NUL|COM[1-9¹²³]|LPT[1-9¹²³])(?:\.|$)", re.I)


def _fail(code: str, message: str) -> None:
    raise ValueError(f"Reference fidelity failed: {code}: {message}")


def _object(value: Any, required=(), optional=(), *, open_fields=False) -> dict:
    if not isinstance(value, dict) or any(not isinstance(key, str) for key in value):
        _fail("InvalidInput", "Expected a JSON object.")
    missing = set(required) - value.keys()
    if missing:
        _fail("InvalidInput", f"Missing properties: {', '.join(sorted(missing))}.")
    if not open_fields:
        unknown = value.keys() - set(required) - set(optional)
        if unknown:
            _fail("InvalidInput", f"Unknown properties: {', '.join(sorted(unknown))}.")
    return value


def _text(value: Any, context: str, *, empty=False) -> str:
    if not isinstance(value, str) or len(value) > 32768 or (not empty and not value.strip()):
        _fail("InvalidInput", f"{context} must be a bounded string.")
    return value


def _array(value: Any, context: str, maximum=RECORD_LIMIT) -> list:
    if not isinstance(value, list) or len(value) > maximum:
        _fail("InvalidInput", f"{context} must be an array of at most {maximum} items.")
    return value


def _number(value: Any, context: str, *, positive=False) -> float:
    if (type(value) not in (int, float) or abs(value) > 1000000
            or not math.isfinite(value) or (positive and value <= 0)):
        _fail("InvalidInput", f"{context} must be a finite number in range.")
    return value


def _bounded_tree(value: Any, depth=0) -> None:
    if isinstance(value, (dict, list)):
        if depth >= DEPTH_LIMIT:
            _fail("InputLimit", "JSON exceeds 64 nesting levels.")
        if isinstance(value, dict):
            keys = set()
            for key, child in value.items():
                if not isinstance(key, str):
                    _fail("InvalidInput", "JSON property names must be strings.")
                if key.upper() in keys:
                    _fail("DuplicateProperty", f"Duplicate JSON property '{key}'.")
                keys.add(key.upper())
                _bounded_tree(child, depth + 1)
        else:
            for child in value:
                _bounded_tree(child, depth + 1)
    elif value is not None and type(value) not in (str, bool, int, float):
        _fail("InvalidInput", "Only JSON values are supported.")
    elif isinstance(value, float) and not math.isfinite(value):
        _fail("InvalidInput", "Nonfinite JSON numbers are forbidden.")


def _pairs(pairs) -> dict:
    result = {}
    seen = set()
    for key, value in pairs:
        if key.upper() in seen:
            _fail("DuplicateProperty", f"Duplicate JSON property '{key}'.")
        seen.add(key.upper())
        result[key] = value
    return result


def _check_stat(info, path: Path) -> None:
    if stat.S_ISLNK(info.st_mode):
        _fail("UnsafePath", f"Symbolic links are forbidden: {path}")
    attributes = getattr(info, "st_file_attributes", 0)
    if not stat.S_ISDIR(info.st_mode) and attributes & 0x441000:
        _fail("UnsafePath", "Offline/recall files must already be locally hydrated.")
    if attributes & 0x400:
        tag = getattr(info, "st_reparse_tag", 0)
        if tag & 0xFFFF0FFF != 0x9000001A:
            _fail("UnsafePath", "Only non-link cloud reparse tags are allowed.")


def _local_path(value: str | Path, base: Path | None = None) -> Path:
    raw = os.fspath(value) if isinstance(value, (str, Path)) else None
    _text(raw, "path")
    if raw.startswith(("\\", "//")) or re.search(r"[\x00-\x1f<>\"|?*]", raw):
        _fail("UnsafePath", "UNC, device, URL, and invalid paths are forbidden.")
    windows_absolute = re.match(r"^[A-Za-z]:[\\/]", raw) is not None
    if windows_absolute and os.name != "nt":
        _fail("UnsafePath", "A foreign drive path is not local; use a contained relative path.")
    absolute = windows_absolute or (os.name != "nt" and raw.startswith("/"))
    tail = raw[3:] if windows_absolute else raw[1:] if absolute else raw
    if ":" in tail or (os.name == "nt" and raw.startswith("/")):
        _fail("UnsafePath", "Streams, URLs, and drive-relative paths are forbidden.")
    parts = re.split(r"[\\/]", tail)
    for part in parts:
        if (not part or part in (".", "..") or part.endswith((".", " "))
                or _DEVICE.match(part)):
            _fail("UnsafePath", "Traversal, device names, and ambiguous segments are forbidden.")
    if absolute:
        path = Path(raw)
    elif base is not None:
        path = base.joinpath(*parts)
    else:
        _fail("UnsafePath", "An absolute local model path is required.")
    if not path.is_absolute():
        _fail("UnsafePath", "An absolute local path is required.")
    if not absolute and base is not None:
        try:
            path.relative_to(base)
        except ValueError:
            _fail("UnsafePath", "Relative path escaped its containing folder.")
    if os.name == "nt":
        import ctypes

        if ctypes.windll.kernel32.GetDriveTypeW(str(path.anchor)) == 4:
            _fail("UnsafePath", "Network drives are forbidden.")
    for entry in (*reversed(path.parents), path):
        try:
            info = entry.lstat()
        except FileNotFoundError:
            break
        _check_stat(info, entry)
    return path


def _read_json(path: Path) -> Any:
    info = path.lstat()
    if not stat.S_ISREG(info.st_mode):
        _fail("InvalidInput", "JSON input must be a regular file.")
    if info.st_size > JSON_LIMIT:
        _fail("InputLimit", "JSON exceeds 10 MiB.")
    with path.open("rb") as stream:
        raw = stream.read(JSON_LIMIT + 1)
    if len(raw) > JSON_LIMIT:
        _fail("InputLimit", "JSON exceeds 10 MiB.")
    try:
        result = json.loads(
            raw, object_pairs_hook=_pairs,
            parse_constant=lambda _: _fail("InvalidInput", "Nonfinite JSON numbers are forbidden."),
        )
    except (json.JSONDecodeError, UnicodeError, RecursionError) as error:
        _fail("InvalidInput", f"Invalid JSON: {error}")
    _bounded_tree(result)
    return result


def _normalize(text: str) -> str:
    return re.sub(r"\s+", " ", text.strip()).upper()


def _lines(value: Any) -> list:
    for line in _array(value, "requiredText"):
        _text(line, "requiredText entry")
        if "\r" in line or "\n" in line:
            _fail("InvalidInput", "Contract text entries must be single logical lines.")
    return value


def _check_label(actual: Any, title: str, required: list, source_id: str) -> None:
    _text(actual, f"label for '{source_id}'", empty=True)
    expected = {_normalize(line) for line in required}
    if title:
        expected.add(_normalize(title))
    lines = [_normalize(line) for line in re.split(r"\r\n|\n|\r", actual) if line.strip()]
    if title and (not lines or lines[0] != _normalize(title)):
        _fail("TitleMismatch", f"Original title changed for '{source_id}'.")
    seen = set()
    for line in lines:
        if line not in expected:
            _fail("UnexpectedText", f"Uncontracted label line for '{source_id}'.")
        if line in seen:
            _fail("DuplicateText", f"Repeated label line for '{source_id}'.")
        seen.add(line)
    if expected - seen:
        _fail("MissingText", f"Original label detail missing for '{source_id}'.")


def _validate_contract(contract: Any) -> tuple[dict, dict, float]:
    _object(contract, ("schemaVersion", "mode", "source", "referencePage", "components",
                       "relationships", "layout", "unresolved"), ("allowAdditionalPages",))
    if _number(contract["schemaVersion"], "contract schemaVersion") != 1 or contract["mode"] != "source-faithful":
        _fail("InvalidInput", "Require schemaVersion 1 and mode source-faithful.")
    source = _object(contract["source"], ("path", "sha256", "role"))
    _text(source["path"], "source.path")
    if source["role"] != "authoritative-reference":
        _fail("SourceRole", "Only an authoritative-reference can be the source.")
    if not re.fullmatch(r"[0-9a-fA-F]{64}", _text(source["sha256"], "source.sha256")):
        _fail("InvalidInput", "source.sha256 requires exactly 64 hexadecimal characters.")
    _text(contract["referencePage"], "referencePage")
    if type(contract.get("allowAdditionalPages", False)) is not bool:
        _fail("InvalidInput", "allowAdditionalPages must be boolean.")
    for unresolved in _array(contract["unresolved"], "unresolved"):
        _text(unresolved, "unresolved entry")
    if contract["unresolved"]:
        _fail("UnresolvedSource", "Resolve source extraction uncertainties before acceptance.")
    components = {}
    relationships = {}
    all_ids = set()
    for category, mapped in (("components", components), ("relationships", relationships)):
        for item in _array(contract[category], category):
            if category == "components":
                _object(item, ("id", "label", "requiredText", "parent", "kind"))
            else:
                _object(item, ("id", "source", "target", "direction"), ("requiredText",))
            source_id = _text(item["id"], "source id")
            if source_id in all_ids:
                _fail("DuplicateSourceId", f"Duplicate contract source id '{source_id}'.")
            all_ids.add(source_id)
            _lines(item.get("requiredText", []))
            if category == "components":
                label = _text(item["label"], "component label")
                if "\r" in label or "\n" in label or _text(item["kind"], "component kind") not in _KINDS:
                    _fail("InvalidInput", "Component title must be one line; kind must be card/container/note.")
                if item["parent"] is not None:
                    _text(item["parent"], "component parent")
            else:
                _text(item["source"], "relationship source")
                _text(item["target"], "relationship target")
                if _text(item["direction"], "relationship direction") not in _DIRECTIONS:
                    _fail("InvalidInput", "Relationship direction must be explicit.")
            mapped[source_id] = item
    if not components:
        _fail("InvalidInput", "The source contract must contain components.")
    for item in components.values():
        ancestors = set()
        current = item
        while current["parent"] is not None:
            parent = current["parent"]
            if parent not in components or components[parent]["kind"] != "container":
                _fail("InvalidInput", "Contract parent must reference a source container.")
            if parent in ancestors:
                _fail("InvalidInput", "Contract parent cycle.")
            ancestors.add(parent)
            current = components[parent]
    for item in relationships.values():
        if item["source"] not in components or item["target"] not in components:
            _fail("InvalidInput", "Contract relationship has an unknown endpoint.")
    layout = _object(contract["layout"], ("leftToRight", "topToBottom", "aspectRatio"), ("aspectTolerance",))
    _number(layout["aspectRatio"], "aspectRatio", positive=True)
    tolerance = _number(layout.get("aspectTolerance", 0.15), "aspectTolerance")
    if not 0 <= tolerance < 1:
        _fail("InvalidInput", "aspectTolerance must be at least zero and less than one.")
    for axis in ("leftToRight", "topToBottom"):
        for chain in _array(layout[axis], axis):
            _array(chain, "layout chain")
            if len(chain) < 2:
                _fail("InvalidInput", "Layout chains require at least two source IDs.")
            seen = set()
            for source_id in chain:
                _text(source_id, "layout id")
                if source_id not in components or source_id in seen:
                    _fail("InvalidInput", "Layout has an unknown or repeated source ID.")
                seen.add(source_id)
    return components, relationships, tolerance


def _validate_pages(model: dict, contract: dict) -> dict:
    _object(model, ("schemaVersion", "pages"), open_fields=True)
    if _number(model["schemaVersion"], "model schemaVersion") != 1:
        _fail("InvalidInput", "Model schemaVersion must be 1.")
    pages = _array(model["pages"], "pages", 100)
    names = set()
    reference = None
    for page in pages:
        _object(page, ("name", "width", "height", "nodes", "edges"), open_fields=True)
        name = _text(page["name"], "page name")
        if name.upper() in names:
            _fail("InvalidInput", "Duplicate page name.")
        names.add(name.upper())
        for size in ("width", "height"):
            _number(page[size], f"page {size}", positive=True)
        native_ids = set()
        for category in ("nodes", "edges"):
            for item in _array(page[category], category):
                _object(item, ("id",), open_fields=True)
                native_id = _text(item["id"], "model id")
                if native_id.upper() in native_ids:
                    _fail("DuplicateModelId", f"Duplicate model id '{native_id}'.")
                native_ids.add(native_id.upper())
                if "sourceId" in item:
                    _text(item["sourceId"], "model sourceId")
                if category == "nodes":
                    _object(item, ("id", "kind", "label", "x", "y", "width", "height"), open_fields=True)
                    if _text(item["kind"], "node kind") not in _KINDS:
                        _fail("InvalidInput", "Unsupported model node kind.")
                    _text(item["label"], "node label", empty=True)
                    for coordinate in ("x", "y"):
                        _number(item[coordinate], coordinate)
                    for size in ("width", "height"):
                        _number(item[size], size, positive=True)
                    if item.get("parent") is not None:
                        _text(item["parent"], "model parent", empty=True)
                else:
                    _object(item, ("id", "source", "target"), open_fields=True)
                    for endpoint in ("source", "target"):
                        if item[endpoint] is not None:
                            _text(item[endpoint], "model endpoint")
                    _text(item.get("label", ""), "edge label", empty=True)
                    if "direction" in item:
                        _text(item["direction"], "edge direction")
        if name == contract["referencePage"]:
            reference = page
    if len(pages) > 1 and not contract.get("allowAdditionalPages", False):
        _fail("AdditionalPages", "Supplemental pages require explicit contract permission.")
    if reference is None:
        _fail("MissingReferencePage", "The designated reference page is missing.")
    if "outputContract" in model:
        if model["outputContract"] != "architecture-pack-v1.6":
            _fail("InvalidInput", "Unknown outputContract.")
        if len(pages) != 3 or [page.get("view") for page in pages] != ["main", "hardening", "flowchart"]:
            _fail("ReferencePageOrder", "V1.6 requires exactly three ordered main/hardening/flowchart pages.")
        if reference is not pages[0] or not contract.get("allowAdditionalPages", False):
            _fail("ReferencePageOrder", "The faithful reference must be page one, with allowAdditionalPages true.")
    return reference


def _compare(reference: dict, contract: dict, components: dict, relationships: dict, tolerance: float) -> dict:
    native_nodes = {node["id"]: node for node in reference["nodes"]}
    nodes = {}
    edges = {}
    for category, expected, mapped in (("nodes", components, nodes), ("edges", relationships, edges)):
        for item in reference[category]:
            if category == "edges":
                _object(item, ("id", "source", "target", "direction"), open_fields=True)
            source_id = item.get("sourceId")
            if source_id not in expected:
                _fail("UnexpectedItem", f"Uncontracted {category} item '{item['id']}' on reference page.")
            if source_id in mapped:
                _fail("DuplicateSourceId", f"Repeated model sourceId '{source_id}'.")
            mapped[source_id] = item
        if expected.keys() - mapped.keys():
            _fail("MissingItem", f"Missing contracted {category} sourceIds: {sorted(expected.keys() - mapped.keys())}.")
    for source_id, item in nodes.items():
        expected = components[source_id]
        if item["kind"] != expected["kind"]:
            _fail("KindMismatch", f"Component kind changed for '{source_id}'.")
        _check_label(item["label"], expected["label"], expected["requiredText"], source_id)
        parent_source = None
        if item.get("parent"):
            parent = native_nodes.get(item["parent"])
            if parent is None or parent.get("sourceId") is None:
                _fail("ParentMismatch", f"Unknown or uncontracted parent for '{source_id}'.")
            parent_source = parent["sourceId"]
        if parent_source != expected["parent"]:
            _fail("ParentMismatch", f"Source grouping changed for '{source_id}'.")
    for source_id, item in edges.items():
        expected = relationships[source_id]
        for endpoint in ("source", "target"):
            _text(item[endpoint], "model relationship endpoint")
            target = native_nodes.get(item[endpoint])
            if target is None or target.get("sourceId") != expected[endpoint]:
                _fail("EndpointMismatch", f"Relationship {endpoint} changed for '{source_id}'.")
        if item["direction"] != expected["direction"]:
            _fail("DirectionMismatch", f"Relationship direction changed for '{source_id}'.")
        _check_label(item.get("label", ""), "", expected.get("requiredText", []), source_id)
    for axis in ("leftToRight", "topToBottom"):
        for chain in contract["layout"][axis]:
            for before, after in zip(chain, chain[1:]):
                a, b = nodes[before], nodes[after]
                gap = b["x"] - b["width"] / 2 - (a["x"] + a["width"] / 2)
                if axis == "topToBottom":
                    gap = a["y"] - a["height"] / 2 - (b["y"] + b["height"] / 2)
                if gap <= 0:
                    _fail("LayoutMismatch", f"{axis} clear separation lost between '{before}' and '{after}'.")
    error = abs((reference["width"] / reference["height"]) / contract["layout"]["aspectRatio"] - 1)
    if error > tolerance:
        _fail("AspectMismatch", "Reference page aspect ratio is outside its contracted tolerance.")
    return {"expectedComponents": len(components), "matchedComponents": len(nodes),
            "expectedRelationships": len(relationships), "matchedRelationships": len(edges)}


def validate_reference(model: dict, model_path: Path, contract_path: Path | str | None = None) -> dict:
    """Validate the rendered model against a separately persisted source contract.

    Paths are local, source bytes are only hashed, and caller data is never modified.
    The persisted model JSON is also checked for duplicate keys and resource limits;
    fidelity comparisons use the supplied model, including any caller-side edits.
    """
    try:
        _object(model, ("schemaVersion", "pages"), open_fields=True)
        _bounded_tree(model)
        if len(json.dumps(model, ensure_ascii=True, allow_nan=False).encode("utf-8")) > JSON_LIMIT:
            _fail("InputLimit", "Model exceeds 10 MiB.")
        if "conversionMode" in model and model["conversionMode"] not in (
                "new-design", "faithful", "reference-plus-proposal"):
            _fail("InvalidInput", "Unsupported conversionMode.")
        local_model = _local_path(model_path)
        _object(_read_json(local_model), ("schemaVersion", "pages"), open_fields=True)
        selected = contract_path if contract_path is not None else model.get("referenceContract")
        if selected is None:
            _fail("MissingContract", "Reference conversion requires a persisted source contract.")
        reference_path = _local_path(selected, local_model.parent)
        contract = _read_json(reference_path)
        components, relationships, tolerance = _validate_contract(contract)
        source_path = _local_path(contract["source"]["path"], reference_path.parent)
        if not source_path.exists():
            _fail("MissingSource", "Persisted authoritative source file is missing; no fallback is permitted.")
        if source_path.samefile(local_model) or source_path.samefile(reference_path):
            _fail("SourceRole", "Source must be a separate artifact from model and contract.")
        info = source_path.lstat()
        if not stat.S_ISREG(info.st_mode):
            _fail("InvalidInput", "Source must be a regular file.")
        if not 0 < info.st_size <= SOURCE_LIMIT:
            _fail("InputLimit", "Source must contain 1 byte to 100 MiB.")
        digest = hashlib.sha256()
        total = 0
        with source_path.open("rb") as stream:
            while chunk := stream.read(1024 * 1024):
                total += len(chunk)
                if total > SOURCE_LIMIT:
                    _fail("InputLimit", "Source exceeds 100 MiB.")
                digest.update(chunk)
        actual = digest.hexdigest()
        expected = contract["source"]["sha256"].lower()
        if actual != expected:
            _fail("SourceHashMismatch", "Persisted source SHA256 changed; reacquire/review the source.")
        reference = _validate_pages(model, contract)
        counts = _compare(reference, contract, components, relationships, tolerance)
        return {
            "schemaVersion": 1, "mode": "source-faithful", "valid": True,
            "modelPath": str(local_model), "referencePath": str(reference_path),
            "referencePage": contract["referencePage"],
            "source": {"path": str(source_path), "expectedSha256": expected,
                       "actualSha256": actual, "verified": True},
            "counts": counts, "errors": [], "comUsed": False,
            "imageRecognitionPerformed": False, "dataFlowVerified": False,
        }
    except (OSError, UnicodeError, RecursionError, OverflowError) as error:
        _fail("InvalidInput", str(error))
