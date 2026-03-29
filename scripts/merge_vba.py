#!/usr/bin/env python3
from __future__ import annotations

import argparse
import pathlib
import re
import sys
from dataclasses import dataclass


ATTR_RE = re.compile(r"^\s*Attribute\s+VB_", re.IGNORECASE)
OPTION_EXPLICIT_RE = re.compile(r"^\s*Option\s+Explicit\s*$", re.IGNORECASE)
OPTION_PRIVATE_RE = re.compile(r"^\s*Option\s+Private\s+Module\s*$", re.IGNORECASE)

# Detect first top-level procedure in a standard module.
PROC_START_RE = re.compile(
    r"""
    ^\s*
    (?:
        Public|Private|Friend|Static
    )?
    \s*
    (?:
        Sub
        |Function
        |Property(?:\s+Get|\s+Let|\s+Set)?
    )
    \b
    """,
    re.IGNORECASE | re.VERBOSE,
)


@dataclass
class ParsedModule:
    path: pathlib.Path
    has_option_private: bool
    decl_lines: list[str]
    proc_lines: list[str]


def read_text(path: pathlib.Path) -> str:
    data = path.read_bytes()
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", errors="replace")


def split_lines(text: str) -> list[str]:
    return text.replace("\r\n", "\n").replace("\r", "\n").split("\n")


def strip_leading_and_trailing_blank_lines(lines: list[str]) -> list[str]:
    out = list(lines)
    while out and out[0].strip() == "":
        out.pop(0)
    while out and out[-1].strip() == "":
        out.pop()
    return out


def find_first_proc_index(lines: list[str]) -> int | None:
    for i, line in enumerate(lines):
        if PROC_START_RE.match(line):
            return i
    return None


def parse_module(path: pathlib.Path) -> ParsedModule:
    raw_lines = split_lines(read_text(path))

    cleaned: list[str] = []
    has_option_private = False

    for line in raw_lines:
        if ATTR_RE.match(line):
            continue
        if OPTION_EXPLICIT_RE.match(line):
            continue
        if OPTION_PRIVATE_RE.match(line):
            has_option_private = True
            continue
        cleaned.append(line)

    first_proc = find_first_proc_index(cleaned)

    if first_proc is None:
        decl_lines = cleaned
        proc_lines: list[str] = []
    else:
        decl_lines = cleaned[:first_proc]
        proc_lines = cleaned[first_proc:]

    decl_lines = strip_leading_and_trailing_blank_lines(decl_lines)
    proc_lines = strip_leading_and_trailing_blank_lines(proc_lines)

    return ParsedModule(
        path=path,
        has_option_private=has_option_private,
        decl_lines=decl_lines,
        proc_lines=proc_lines,
    )


def add_section(parts: list[str], title: str, lines: list[str]) -> None:
    parts.append("'" + "=" * 70)
    parts.append(f"' {title}")
    parts.append("'" + "=" * 70)
    if lines:
        parts.extend(lines)
    else:
        parts.append("' (empty)")
    parts.append("")


def build_merged(module_name: str, files: list[pathlib.Path]) -> str:
    parsed = [parse_module(path) for path in files]

    parts: list[str] = []
    parts.append(f'Attribute VB_Name = "{module_name}"')
    parts.append("Option Explicit")

    if any(m.has_option_private for m in parsed):
        parts.append("Option Private Module")

    parts.append("")

    # Hoisted declarations first
    for module in parsed:
        rel = str(module.path).replace("\\", "/")
        add_section(parts, f"BEGIN DECLARATIONS {rel}", module.decl_lines)

    # Then all procedures
    for module in parsed:
        rel = str(module.path).replace("\\", "/")
        add_section(parts, f"BEGIN PROCEDURES {rel}", module.proc_lines)

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
