from __future__ import annotations

"""
Entry point for aggregating CSV results into an Excel workbook.

This module reads a directory structure of the form::

    A/
      n10/
        m01/
          jw/circuit.csv
          circuit.csv
        m02/
          jw/circuit.csv
          circuit.csv
      n20/
        m01/
          jw/circuit.csv
          circuit.csv
        ...

Each ``n*`` folder contains several ``m*`` folders. For each ``m*`` folder, the script
creates two worksheets in the output Excel file: one for single‑point values
(labelled by the m‑folder's name) and one for transient data (named
``{m_name}_過渡(V3+V6)``). The single‑point sheet records values from
``jw/circuit.csv`` for each ``n*`` folder. The transient sheet pastes whole
columns from ``circuit.csv`` for each ``n*`` folder, offsetting the data
horizontally so that each ``n*`` occupies its own group of three columns.

The default cell extraction and column mappings are defined via the
``read_to_write_map`` argument and may be customised by callers. See the
documentation for ``run`` for details.
"""

import sys
import traceback
from pathlib import Path
from datetime import datetime
from typing import Optional, List, Dict, Any

from io_fs import list_subfolders_sorted, circuit_csv_path, circuit_csv_path_root, natural_key
from io_table import (
    load_workbook,
    ensure_sheet_by_index_and_rename,
    read_csv_cell_1based,
    read_csv_column_1based,
    col_letter,
)

try:
    import tkinter as tk  # type: ignore
    from tkinter import filedialog, messagebox  # type: ignore
except Exception:
    # In headless environments tkinter may not be available
    tk = None  # type: ignore


def run(
    A: Path,
    template_xlsx: Path,
    output_xlsx: Optional[Path] = None,
    header_row: bool = True,
    read_to_write_map: Optional[List[Dict[str, Any]]] = None,
) -> Path:
    """
    Aggregate measurement results from a hierarchical folder structure into an Excel workbook.

    Parameters
    ----------
    A : Path
        Root folder that contains n‑folders (e.g. ``n10``, ``n20``). Each n‑folder contains
        m‑folders containing measurement data.
    template_xlsx : Path
        Path to the Excel template to populate. The template should provide enough
        worksheets to accommodate the number of m‑folders; additional worksheets will be
        created automatically if necessary.
    output_xlsx : Optional[Path], default ``None``
        Destination path for the generated workbook. If omitted, a timestamped filename
        derived from the template's name will be used.
    header_row : bool, default ``True``
        Indicates whether the first row of the jw sheets contains headers. If ``True``
        then data begins on the second row; otherwise it starts on the first row.
    read_to_write_map : Optional[List[Dict[str, Any]]]
        Mapping describing which cell to extract from each ``jw/circuit.csv`` and which
        column to write it into. Each mapping should contain the keys:
        ``src_row`` (1‑based), ``src_col`` (1‑based), ``dst_col`` (1‑based Excel column), and
        an optional ``label`` for logging.

    Returns
    -------
    Path
        The path to the generated Excel workbook.

    Notes
    -----
    The original implementation iterated over n‑folders first and m‑folders second. That
    meant the transient (過渡) sheets were recreated and overwritten on each iteration of
    the outer loop, leading to only the last set of results being retained. The logic has
    been refactored so that the m‑folders are processed outermost and results for each
    n‑folder are accumulated horizontally on the corresponding transient sheet. This
    matches the expected output where each transient sheet contains multiple three‑column
    groups, one per n‑folder.
    """
    A = A.resolve()
    template_xlsx = template_xlsx.resolve()

    # Provide a default mapping if none is supplied
    if read_to_write_map is None:
        read_to_write_map = [
            {"src_row": 8, "src_col": 4, "dst_col": 4, "label": "D8"},
            {"src_row": 8, "src_col": 8, "dst_col": 2, "label": "H8"},
        ]

    # Generate output filename if not explicitly provided
    if output_xlsx is None:
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_xlsx = template_xlsx.with_name(f"{template_xlsx.stem}_out_{stamp}.xlsx")
    else:
        output_xlsx = output_xlsx.resolve()

    # Basic sanity checks
    if not A.is_dir():
        raise FileNotFoundError(f"A folder not found: {A}")
    if not template_xlsx.is_file():
        raise FileNotFoundError(f"Template xlsx not found: {template_xlsx}")

    # Load the workbook from the template
    wb = load_workbook(template_xlsx)

    # Prepare a log buffer
    log_path = output_xlsx.with_suffix(".log.txt")
    log: List[str] = []
    log.append(f"Run at: {datetime.now().isoformat(timespec='seconds')}")
    log.append(f"A: {A}")
    log.append(f"Template: {template_xlsx}")
    log.append(f"Output: {output_xlsx}")
    log.append("")

    # Discover all n‑folders
    n_folders = list_subfolders_sorted(A)
    log.append(f"Found n folders: {len(n_folders)}")
    log.append("")

    # Exit early if no n‑folders present
    if not n_folders:
        wb.save(output_xlsx)
        log_path.write_text("\n".join(log), encoding="utf-8")
        return output_xlsx

    # Build a unique ordered list of m‑folder names across all n‑folders
    m_folder_names: List[str] = []
    m_folder_map: Dict[str, Path] = {}
    for n_folder in n_folders:
        for m_path in list_subfolders_sorted(n_folder):
            name = m_path.name
            if name not in m_folder_map:
                m_folder_map[name] = m_path
                m_folder_names.append(name)

    # Sort m‑folders naturally to maintain stable sheet ordering
    m_folder_names.sort(key=lambda s: natural_key(s))

    # Compute row offsets
    row_offset = 1 if header_row else 0
    # The row on which to start pasting transient data columns
    start_row_for_columns = 2 if header_row else 1

    # Number of m‑folders dictates how sheets are divided between jw and transient
    mmax = len(m_folder_names)

    # Iterate over each m‑folder. Each m gets its own jw sheet and transient sheet.
    for m_idx, m_name in enumerate(m_folder_names):
        # Sheet positions are 1‑based: jw sheets first, transient sheets after
        jw_sheet_index = m_idx + 1
        tr_sheet_index = mmax + m_idx + 1

        # Acquire or create the jw sheet and rename it once
        ws = ensure_sheet_by_index_and_rename(wb, jw_sheet_index, desired_title=m_name)

        # Acquire or create the transient sheet and rename it appropriately
        trans_title = f"{m_name}_過渡(V3+V6)"
        ws_tr = ensure_sheet_by_index_and_rename(wb, tr_sheet_index, desired_title=trans_title)

        # Process each n‑folder for this m
        for n_idx, n_folder in enumerate(n_folders):
            # Row on jw sheet where data for this n will be written
            target_row = row_offset + n_idx + 1

            # Determine the actual m_folder path for this n (combine n and m names)
            m_folder = n_folder / m_name
            log.append(f"[m={m_idx+1}/{mmax}] {m_name} @ n={n_folder.name}")

            # Populate column A in the jw sheet with the n folder name (n value).
            # This provides a label for each n on the left side of the sheet.  We place
            # this before reading jw or root data so it executes regardless of CSV existence.
            ws.cell(row=target_row, column=1, value=n_folder.name)

            # (A) Extract specified cell values from jw/circuit.csv
            jw_csv = circuit_csv_path(n_folder, m_folder, jw_folder_name="jw", csv_name="circuit.csv")
            try:
                if not jw_csv.exists():
                    log.append(f"  MISSING {jw_csv}")
                else:
                    for item in read_to_write_map:
                        src_r = int(item["src_row"])
                        src_c = int(item["src_col"])
                        dst_c = int(item["dst_col"])
                        label = str(item.get("label", f"R{src_r}C{src_c}"))

                        val, raw = read_csv_cell_1based(jw_csv, row_1based=src_r, col_1based=src_c)
                        dstL = col_letter(dst_c)

                        ws.cell(row=target_row, column=dst_c, value=val if val is not None else None)

                        if val is None:
                            log.append(
                                f"    {label}(raw='{raw}') -> sheet#{jw_sheet_index}('{ws.title}') {dstL}{target_row}=BLANK"
                            )
                        else:
                            log.append(
                                f"    {label}={val} -> sheet#{jw_sheet_index}('{ws.title}') {dstL}{target_row}"
                            )
            except Exception as e:
                log.append(f"    ERROR reading jw {jw_csv}: {e}")
                log.append("      " + traceback.format_exc().replace("\n", "\n      "))

            # (B) Paste entire columns from root circuit.csv onto the transient sheet
            root_csv = circuit_csv_path_root(m_folder, csv_name="circuit.csv")
            try:
                if not root_csv.exists():
                    log.append(f"  MISSING {root_csv}")
                else:
                    # Columns H, Q and T in the CSV correspond to 1‑based indices 8, 17 and 20
                    src_cols = [8, 17, 20]
                    # Determine the starting destination column for this n. n_idx is 0‑based.
                    base_dst = 10 + 3 * (n_idx + 1)
                    dst_cols = [base_dst, base_dst + 1, base_dst + 2]

                    for src_c, dst_c in zip(src_cols, dst_cols):
                        values = read_csv_column_1based(root_csv, col_1based=src_c)
                        dstL = col_letter(dst_c)

                        for i, v in enumerate(values):
                            ws_tr.cell(row=start_row_for_columns + i, column=dst_c, value=v)

                        log.append(
                            f"    root {col_letter(src_c)}(:) -> sheet#{tr_sheet_index}('{ws_tr.title}') "
                            f"{dstL}{start_row_for_columns}:{dstL}{start_row_for_columns + len(values) - 1}"
                        )

                    # -- Additional header and formula handling for transient analysis --
                    # ② Place the n folder name in the header row (row 1) of the transient sheet.
                    #     Each n occupies three columns starting at base_dst (I, V3, V6). Only set the first column.
                    ws_tr.cell(row=1, column=base_dst, value=n_folder.name)

                    # ③ Write the series labels "I", "V3", "V6" into the second row of the transient sheet
                    ws_tr.cell(row=2, column=base_dst, value="I")
                    ws_tr.cell(row=2, column=base_dst + 1, value="V3")
                    ws_tr.cell(row=2, column=base_dst + 2, value="V6")

                    # ④ Set RMS formulas for I and V on the jw sheet.
                    #     B列 (column 2) : RMS of the current (I) over rows 1283–1538 of the I column for this n
                    #     D列 (column 4) : RMS of V3 plus RMS of V6 over the same row range. V3 and V6 are the
                    #                     second and third columns of each 3‑column group.
                    col_I = col_letter(base_dst)
                    col_V3 = col_letter(base_dst + 1)
                    col_V6 = col_letter(base_dst + 2)
                    # B列 formula
                    formula_I = f"=SQRT(SUMSQ(${col_I}$1283:${col_I}$1538)/256)"
                    # D列 formula: sum of RMS of V3 and V6
                    formula_V = (
                        f"=SQRT(SUMSQ(${col_V3}$1283:${col_V3}$1538)/256)"
                        f"+SQRT(SUMSQ(${col_V6}$1283:${col_V6}$1538)/256)"
                    )
                    ws.cell(row=target_row, column=2, value=formula_I)
                    ws.cell(row=target_row, column=4, value=formula_V)
            except Exception as e:
                log.append(f"    ERROR reading root {root_csv}: {e}")
                log.append("      " + traceback.format_exc().replace("\n", "\n      "))

        log.append("")

    # Persist the workbook and log
    wb.save(output_xlsx)
    log_path.write_text("\n".join(log), encoding="utf-8")
    return output_xlsx


def gui_main() -> None:
    """Entry point when running the script with a GUI.

    Prompts the user to select the root folder and template workbook using file dialogs.
    Collects whether the template contains header rows and then runs the aggregation.
    Displays the output and log locations on completion. Only available when tkinter is
    installed.
    """
    if tk is None:
        raise RuntimeError("tkinter is not available.")

    root = tk.Tk()
    root.withdraw()

    messagebox.showinfo("No-load aggregator", "A（親フォルダ）を選択してください。")
    A_dir = filedialog.askdirectory(title="Select A folder")
    if not A_dir:
        return

    messagebox.showinfo("No-load aggregator", "テンプレートExcel（無負荷特性...xlsx）を選択してください。")
    xlsx_file = filedialog.askopenfilename(
        title="Select template xlsx",
        filetypes=[("Excel files", "*.xlsx")]
    )
    if not xlsx_file:
        return

    header = messagebox.askyesno(
        "No-load aggregator",
        "Excelの1行目はヘッダですか？（はい→2行目から書き込み）"
    )

    out_path = run(
        Path(A_dir),
        Path(xlsx_file),
        output_xlsx=None,
        header_row=header,
        read_to_write_map=[
            {"src_row": 8, "src_col": 4, "dst_col": 4, "label": "D8"},
            {"src_row": 8, "src_col": 8, "dst_col": 2, "label": "H8"},
        ],
    )

    messagebox.showinfo(
        "No-load aggregator",
        f"完了\n出力: {out_path}\nログ: {out_path.with_suffix('.log.txt')}"
    )


if __name__ == "__main__":
    if len(sys.argv) >= 3:
        A_path = Path(sys.argv[1])
        xlsx_path = Path(sys.argv[2])
        out = Path(sys.argv[3]) if len(sys.argv) >= 4 else None
        header_arg = bool(int(sys.argv[4])) if len(sys.argv) >= 5 else True

        out_path = run(
            A_path,
            xlsx_path,
            output_xlsx=out,
            header_row=header_arg,
        )
        print(f"Done: {out_path}")
        print(f"Log : {out_path.with_suffix('.log.txt')}")
    else:
        gui_main()