# -*- coding: utf-8 -*-
"""Viola wrapper: rebuild from D:\\CORDAS_3\\VIOLA 4."""
from __future__ import annotations

from build_ste_preparatory import DEPOSIT_DEFAULTS, build, last_path

OUT = DEPOSIT_DEFAULTS["viola"]["root"] / DEPOSIT_DEFAULTS["viola"]["prep_name"]
HARM_LO = 72


def main() -> list[str]:
    warnings = build(DEPOSIT_DEFAULTS["viola"]["root"], "viola")
    return warnings


if __name__ == "__main__":
    main()
    if last_path:
        print("viola pack", last_path)
