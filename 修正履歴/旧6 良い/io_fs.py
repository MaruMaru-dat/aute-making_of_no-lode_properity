# io_fs.py

from __future__ import annotations

import re
from pathlib import Path
from typing import List

_natural_key_re = re.compile(r"(\d+)")

def natural_key(s: str):
    parts = _natural_key_re.split(s)
    key = []
    for p in parts:
        key.append(int(p) if p.isdigit() else p.lower())
    return key

def list_subfolders_sorted(parent: Path) -> List[Path]:
    if not parent.exists() or not parent.is_dir():
        return []
    subs = [p for p in parent.iterdir() if p.is_dir()]
    subs.sort(key=lambda p: natural_key(p.name))
    return subs

def circuit_csv_path(n_folder: Path, m_folder: Path, jw_folder_name: str = "jw", csv_name: str = "circuit.csv") -> Path:
    return m_folder / jw_folder_name / csv_name

def circuit_csv_path_root(m_folder: Path, csv_name: str = "circuit.csv") -> Path:
    """
    mフォルダ直下: A/nXX/mXX/circuit.csv
    """
    return m_folder / csv_name
