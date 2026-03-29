#!/usr/bin/env python3
from __future__ import annotations

import argparse
import pathlib
import re
import sys

ATTR_RE = re.compile(r"^\s*Attribute\s+VB_", re.IGNORECASE)
OPTION_EXPLICIT_RE = re.compile(r"^\s*Option\s+Explicit\s*$", re.IGNORECASE)
OPTION_PRIVATE_RE = re.compile(r"^\s*Option\s+Private\s+Module\s*$", re.IGNORECASE)


def read_text(path: pathlib.Path) -> str:
    data = path.read_bytes()
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", errors="replace")


def normalize_module_text(text: str) -> list[str]:
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    out: list[str] = []

    for line in lines:
        if ATTR_RE.match(line):
            continue
        if OPTION_EXPLICIT_RE.match(line):
            continue
        if OPTION_PRIVATE_RE.match(line):
            continue
        out.append(line)

    while out and out[0].strip() == "":
        out.pop(0)

    while out and out[-1].strip() == "":
        out.pop()

    return out


def build_merged(module_name: str, files: list[pathlib.Path]) -> str:
    parts: list[str] = []
    parts.append(f'Attribute VB_Name = "{module_name}"')
    parts.append("Option Explicit")
    parts.append("")

    for path in files:
        rel = str(path).replace("\\", "/")
        parts.append("'" + "=" * 70)
        parts.append(f"' BEGIN {rel}")
        parts.append("'" + "=" * 70)

        body_lines = normalize_module_text(read_text(path))
        if body_lines:
            parts.extend(body_lines)
        else:
            parts.append("' (empty module after normalization)")

        parts.append("'" + "=" * 70)
        parts.append(f"' END {rel}")
        parts.append("'" + "=" * 70)
        parts.append("")

    return "\n".join(parts).rstrip() + "\n"


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--output", required=True, help="Path to merged .bas output")
    parser.add_argument("--module-name", default="KATS_All", help="VB module name")
    parser.add_argument("inputs", nargs="+", help="Input .bas files in merge order")
    args = parser.parse_args()

    files = [pathlib.Path(p) for p in args.inputs]
    missing = [str(p) for p in files if not p.exists()]
    if missing:
        print("Missing input files:", file=sys.stderr)
        for m in missing:
            print(f"  {m}", file=sys.stderr)
        return 1

    out_path = pathlib.Path(args.output)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    merged = build_merged(args.module_name, files)
    out_path.write_text(merged, encoding="utf-8", newline="\n")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
