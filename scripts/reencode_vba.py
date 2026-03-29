#!/usr/bin/env python3
from __future__ import annotations

import argparse
import pathlib

from merge_vba import default_output_encoding, read_text, write_text


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("input", help="Input .bas file")
    parser.add_argument("--output", required=True, help="Output path")
    parser.add_argument(
        "--output-encoding",
        default=default_output_encoding(),
        help="Encoding for Word/VBA import files (default: OS-specific)",
    )
    args = parser.parse_args()

    in_path = pathlib.Path(args.input)
    out_path = pathlib.Path(args.output)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    write_text(out_path, read_text(in_path), args.output_encoding)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
