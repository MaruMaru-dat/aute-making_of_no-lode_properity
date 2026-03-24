# main.py

from __future__ import annotations

import sys
import traceback
from pathlib import Path
from datetime import datetime
from typing import Optional, List, Dict, Any

from io_fs import list_subfolders_sorted, circuit_csv_path, circuit_csv_path_root
from io_table import (
    load_workbook,
    ensure_sheet_by_index_and_rename,
    read_csv_cell_1based,
    read_csv_column_1based,
    col_letter,
)

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox
except Exception:
    tk = None


def run(
    A: Path,
    template_xlsx: Path,
    output_xlsx: Optional[Path] = None,
    header_row: bool = True,
    read_to_write_map: Optional[List[Dict[str, Any]]] = None,
) -> Path:
    """
    read_to_write_map: jw/circuit.csv から「単点」抽出して貼る用（今まで通り）
    例:
      [
        {"src_row": 8, "src_col": 4, "dst_col": 4, "label": "D8"},
        {"src_row": 8, "src_col": 8, "dst_col": 2, "label": "H8"},
      ]
    """
    A = A.resolve()
    template_xlsx = template_xlsx.resolve()

    if read_to_write_map is None:
        read_to_write_map = [
            {"src_row": 8, "src_col": 4, "dst_col": 4, "label": "D8"},  # jw: D8 -> D
            {"src_row": 8, "src_col": 8, "dst_col": 2, "label": "H8"},  # jw: H8 -> B
        ]

    if output_xlsx is None:
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_xlsx = template_xlsx.with_name(f"{template_xlsx.stem}_out_{stamp}.xlsx")
    else:
        output_xlsx = output_xlsx.resolve()

    if not A.is_dir():
        raise FileNotFoundError(f"A folder not found: {A}")
    if not template_xlsx.is_file():
        raise FileNotFoundError(f"Template xlsx not found: {template_xlsx}")

    wb = load_workbook(template_xlsx)

    log_path = output_xlsx.with_suffix(".log.txt")
    log = []
    log.append(f"Run at: {datetime.now().isoformat(timespec='seconds')}")
    log.append(f"A: {A}")
    log.append(f"Template: {template_xlsx}")
    log.append(f"Output: {output_xlsx}")
    log.append("")

    n_folders = list_subfolders_sorted(A)
    log.append(f"Found n folders: {len(n_folders)}")
    log.append("")

    row_offset = 1 if header_row else 0   # n=1 -> row2（単点の表）
    start_row_for_columns = 2 if header_row else 1  # 列貼りの開始行（あなたの運用に合わせる）

    for n_idx, n_folder in enumerate(n_folders, start=1):
        m_folders = list_subfolders_sorted(n_folder)
        log.append(f"[n={n_idx}/{len(n_folders)}] {n_folder.name} -> m folders: {len(m_folders)}")

        # 単点貼り（jw）の行：nごとに下へ
        target_row = (n_idx + 1) + row_offset

        for m_idx, m_folder in enumerate(m_folders, start=1):
            sheet_index = m_idx
            ws = ensure_sheet_by_index_and_rename(wb, sheet_index, desired_title=m_folder.name)

            # -------------------------
            # (A) jw/circuit.csv：単点貼り（既存）
            # -------------------------
            jw_csv = circuit_csv_path(n_folder, m_folder, jw_folder_name="jw", csv_name="circuit.csv")
            try:
                if not jw_csv.exists():
                    log.append(f"  [m={m_idx}] {m_folder.name}: MISSING {jw_csv}")
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
                                f"  [m={m_idx}] {m_folder.name}: {label}(raw='{raw}') -> "
                                f"sheet#{sheet_index}('{ws.title}') {dstL}{target_row}=BLANK"
                            )
                        else:
                            log.append(
                                f"  [m={m_idx}] {m_folder.name}: {label}={val} -> "
                                f"sheet#{sheet_index}('{ws.title}') {dstL}{target_row}"
                            )

            except Exception as e:
                log.append(f"  [m={m_idx}] {m_folder.name}: ERROR (jw) {e}")
                log.append("    " + traceback.format_exc().replace("\n", "\n    "))

            root_csv = circuit_csv_path_root(m_folder, csv_name="circuit.csv")
            try:
                if not root_csv.exists():
                    log.append(f"  [m={m_idx}] {m_folder.name}: MISSING {root_csv}")
                else:
                    # mmax は「その n フォルダ配下にある m フォルダ数」
                    mmax = len(m_folders)

                    # ★ 新しい貼り先シート：mmax+m 番目
                    transient_sheet_index = mmax + m_idx
                    transient_title = f"{m_folder.name}_過渡(V3+V6)"
                    ws_tr = ensure_sheet_by_index_and_rename(
                        wb, transient_sheet_index, desired_title=transient_title
                    )

                    # 入力列：H,Q,T（Excel表記）= 8,17,20（1始まり）
                    src_cols = [8, 17, 20]

                    # ★ 出力列：10+3n, 10+3n+1, 10+3n+2（nは1始まり）
                    base_dst = 10 + 3 * n_idx
                    dst_cols = [base_dst, base_dst + 1, base_dst + 2]

                    for src_c, dst_c in zip(src_cols, dst_cols):
                        values = read_csv_column_1based(root_csv, col_1based=src_c)
                        dstL = col_letter(dst_c)

                        for i, v in enumerate(values):
                            ws_tr.cell(row=start_row_for_columns + i, column=dst_c, value=v)

                        log.append(
                            f"  [m={m_idx}] {m_folder.name}: root {col_letter(src_c)}(:) -> "
                            f"sheet#{transient_sheet_index}('{ws_tr.title}') "
                            f"{dstL}{start_row_for_columns}:{dstL}{start_row_for_columns + len(values) - 1}"
                        )

            except Exception as e:
                log.append(f"  [m={m_idx}] {m_folder.name}: ERROR (root->transient) {e}")
                log.append("    " + traceback.format_exc().replace("\n", "\n    "))

        log.append("")

    wb.save(output_xlsx)
    log_path.write_text("\n".join(log), encoding="utf-8")
    return output_xlsx


def gui_main():
    if tk is None:
        raise RuntimeError("tkinter is not available.")

    root = tk.Tk()
    root.withdraw()

    messagebox.showinfo("No-load aggregator", "A（親フォルダ）を選択してください。")
    A = filedialog.askdirectory(title="Select A folder")
    if not A:
        return

    messagebox.showinfo("No-load aggregator", "テンプレートExcel（無負荷特性...xlsx）を選択してください。")
    xlsx = filedialog.askopenfilename(
        title="Select template xlsx",
        filetypes=[("Excel files", "*.xlsx")]
    )
    if not xlsx:
        return

    header = messagebox.askyesno(
        "No-load aggregator",
        "Excelの1行目はヘッダですか？（はい→2行目から書き込み）"
    )

    out_path = run(
        Path(A),
        Path(xlsx),
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
        A = Path(sys.argv[1])
        xlsx = Path(sys.argv[2])
        out = Path(sys.argv[3]) if len(sys.argv) >= 4 else None
        header = bool(int(sys.argv[4])) if len(sys.argv) >= 5 else True

        out_path = run(
            A,
            xlsx,
            output_xlsx=out,
            header_row=header,
        )
        print(f"Done: {out_path}")
        print(f"Log : {out_path.with_suffix('.log.txt')}")
    else:
        gui_main()
