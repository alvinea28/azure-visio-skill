"""Build the complete v1.7 skill from sibling source and JSON icon data."""

import argparse
import hashlib
import io
import json
from pathlib import Path
import sys
import zipfile

import portable_visio as p


def build(output):
    root = Path(__file__).absolute().parent
    destination = p.safe_path(output)
    p.require(destination.suffix == ".zip" and not destination.exists() and destination.parent.is_dir(),
              "Provide a NEW .zip path in an existing output directory")
    manifest_path = root / "offline-assets.json"
    manifest = p.load_json(manifest_path)
    p.require(isinstance(manifest, dict) and manifest.get("kind") == "azure-visio-offline-png-json",
              "The current distribution must contain regular JSON icon data, not inner ZIPs")
    catalog = p.IconCatalog(offline_assets=manifest_path)
    for key, entry in catalog.entries.items():
        p.require(isinstance(entry.get("sourcePage"), str) and entry["sourcePage"].startswith("https://")
                  and isinstance(entry.get("downloadUrl"), str) and entry["downloadUrl"].startswith("https://"),
                  "Bundled artwork must retain its actual source and usage-terms provenance")
        if entry.get("usable") is True:
            catalog.resolve(key)
    example = p.load_json(root / "example-model.json")
    p.validate_model(example, root / "example-model.json")
    p.validate_resources(example, catalog)
    mapping = {
        "SKILL.md": "Cowork-SKILL.txt",
        "README.txt": "Cowork-README.txt",
        "Architecture-Guide.txt": "Cowork-Architecture-Guide.txt",
        "portable_visio.py": "portable_visio.py",
        "portable_reference.py": "portable_reference.py",
        "example-model.json": "example-model.json",
        "offline-assets.json": "offline-assets.json",
        "NOTICE.txt": "NOTICE.txt",
    }
    for companion in [manifest["catalog"], *manifest["dataFiles"]]:
        name = companion["file"]
        p.require(name not in mapping, "Duplicate generated skill companion")
        mapping[name] = name
    files = {name: p.read_bytes(root / source, 5_000_000) for name, source in mapping.items()}
    skill_text = files["SKILL.md"].decode("utf-8-sig").replace("\r\n", "\n")
    p.require(skill_text.startswith("---\nname: azure-visio\n"), "Invalid skill frontmatter")
    p.require("version: 1.7.0" in skill_text, "Skill version mismatch")
    p.require(len(files) <= 21, "Cowork skill exceeds SKILL.md plus twenty companions")
    p.require(len(files["SKILL.md"]) <= 1_000_000, "SKILL.md exceeds upload limit")
    p.require(sum(map(len, files.values())) <= 10_000_000, "Cowork companions exceed 10 MB")
    stream = io.BytesIO()
    with zipfile.ZipFile(stream, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in files.items():
            archive.writestr(name, data)
    data = stream.getvalue()
    p.require(len(data) <= 10_000_000, "Cowork skill ZIP exceeds 10 MB")
    with destination.open("xb") as target:
        target.write(data)
    return {"version": p.VERSION, "package": str(destination), "bytes": len(data),
            "sha256": hashlib.sha256(data).hexdigest(), "files": list(files),
            "expandedBytes": sum(map(len, files.values())),
            "largestCompanionBytes": max(map(len, files.values())),
            "iconCount": len(catalog.entries),
            "packagesBundled": False, "nestedArchives": False, "runtimeInstallationRequired": False}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", required=True)
    args = parser.parse_args()
    try:
        print(json.dumps(build(args.output), indent=2))
        return 0
    except (p.PortableError, OSError, zipfile.BadZipFile) as exc:
        print(json.dumps({"valid": False, "error": str(exc)}), file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
