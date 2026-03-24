"""
Filesystem helpers for the CSV aggregation script.

This module provides utilities to discover and sort subfolders, construct
CSV file paths, and perform natural sorting on folder names containing numbers.
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import List

_natural_key_re = re.compile(r"(\d+)")


def natural_key(s: str) -> List[object]:
    """Generate a key for natural sorting of strings containing numbers.

    Splits the input string into text and numeric parts so that 'n10' comes
    before 'n2'. Non‑numeric parts are compared case‑insensitively.

    Parameters
    ----------
    s : str
        The string to split.

    Returns
    -------
    List[object]
        A list of strings and integers suitable for use as a sort key.
    """
    parts = _natural_key_re.split(s)
    key: List[object] = []
    for p in parts:
        key.append(int(p) if p.isdigit() else p.lower())
    return key


def list_subfolders_sorted(parent: Path) -> List[Path]:
    """List immediate subdirectories of ``parent`` sorted naturally by name."""
    if not parent.exists() or not parent.is_dir():
        return []
    subs = [p for p in parent.iterdir() if p.is_dir()]
    subs.sort(key=lambda p: natural_key(p.name))
    return subs


def circuit_csv_path(n_folder: Path, m_folder: Path, jw_folder_name: str = "jw", csv_name: str = "circuit.csv") -> Path:
    """Construct the path to a jw CSV file.

    Parameters
    ----------
    n_folder : Path
        Path to the n‑folder.
    m_folder : Path
        Path to the m‑folder.
    jw_folder_name : str, default "jw"
        Name of the subfolder containing the jw CSV.
    csv_name : str, default "circuit.csv"
        Name of the CSV file.
    """
    return m_folder / jw_folder_name / csv_name


def circuit_csv_path_root(m_folder: Path, csv_name: str = "circuit.csv") -> Path:
    """Construct the path to the root circuit CSV.

    The root CSV is expected to be located directly under the m‑folder (i.e.
    ``A/nXX/mXX/circuit.csv``).

    Parameters
    ----------
    m_folder : Path
        Path to the m‑folder.
    csv_name : str, default "circuit.csv"
        Name of the CSV file.
    """
    return m_folder / csv_name