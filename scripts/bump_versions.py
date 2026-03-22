#!/usr/bin/env python3
"""Bump PCF manifest and Dataverse solution versions before packaging."""

from __future__ import annotations

import re
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
MANIFEST_PATH = ROOT / "pcf" / "FileExporterComponent" / "FileExporterComponent" / "ControlManifest.Input.xml"
SOLUTION_PATH = ROOT / "solution" / "File_Exporter_Component_Solution" / "src" / "Other" / "Solution.xml"


def bump_semver_patch(version: str) -> str:
    parts = version.split(".")
    if len(parts) != 3 or not all(part.isdigit() for part in parts):
        raise ValueError(f"Unsupported control manifest version format: {version}")

    major, minor, patch = (int(part) for part in parts)
    return f"{major}.{minor}.{patch + 1}"


def bump_four_part_revision(version: str) -> str:
    parts = version.split(".")
    if len(parts) != 4 or not all(part.isdigit() for part in parts):
        raise ValueError(f"Unsupported solution version format: {version}")

    major, minor, build, revision = (int(part) for part in parts)
    return f"{major}.{minor}.{build}.{revision + 1}"


def replace_first(pattern: str, replacement: str, content: str, label: str) -> str:
    updated, count = re.subn(pattern, replacement, content, count=1)
    if count != 1:
        raise ValueError(f"Could not update {label}.")
    return updated


def update_manifest_version() -> tuple[str, str]:
    if not MANIFEST_PATH.exists():
        raise FileNotFoundError(f"Manifest not found: {MANIFEST_PATH}")

    content = MANIFEST_PATH.read_text(encoding="utf-8")
    match = re.search(r'(<control\b[^>]*\bversion=")(\d+\.\d+\.\d+)(")', content)
    if not match:
        raise ValueError("Could not find control manifest version.")

    old_version = match.group(2)
    new_version = bump_semver_patch(old_version)
    updated = replace_first(
        r'(<control\b[^>]*\bversion=")\d+\.\d+\.\d+(")',
        rf"\g<1>{new_version}\g<2>",
        content,
        "control manifest version",
    )
    MANIFEST_PATH.write_text(updated, encoding="utf-8")
    return old_version, new_version


def update_solution_version() -> tuple[str, str]:
    if not SOLUTION_PATH.exists():
        raise FileNotFoundError(f"Solution manifest not found: {SOLUTION_PATH}")

    content = SOLUTION_PATH.read_text(encoding="utf-8")
    match = re.search(r"(<Version>)(\d+\.\d+\.\d+\.\d+)(</Version>)", content)
    if not match:
        raise ValueError("Could not find solution version.")

    old_version = match.group(2)
    new_version = bump_four_part_revision(old_version)
    updated = replace_first(
        r"(<Version>)\d+\.\d+\.\d+\.\d+(</Version>)",
        rf"\g<1>{new_version}\g<2>",
        content,
        "solution version",
    )
    SOLUTION_PATH.write_text(updated, encoding="utf-8")
    return old_version, new_version


def main() -> int:
    try:
        manifest_before, manifest_after = update_manifest_version()
        solution_before, solution_after = update_solution_version()
    except Exception as exc:  # pylint: disable=broad-except
        print(f"Version bump failed: {exc}", file=sys.stderr)
        return 1

    print(f"Control manifest version: {manifest_before} -> {manifest_after}")
    print(f"Solution version: {solution_before} -> {solution_after}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
