"""작업 폴더·파일 역할·열 매핑 후 실제 파일로 BOM↔원본 검토 (Tkinter)."""

from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from bom_review._version import __version__
from bom_review.bom_parse import split_designators
from bom_review.bom_qty import bom_quantity_mismatch_findings
from bom_review.matching import (
    bom_vs_source_findings,
    duplicate_reference_findings,
    iter_error_findings,
    iter_info_findings,
)
from bom_review.table_io import list_files_in_folder, load_header_and_rows, values_for_column

ROLE_NONE = ""
ROLE_BOM = "BOM"
ROLE_SOURCE = "원본"
ROLE_METAL_TOP = "메탈TOP"
ROLE_METAL_BOT = "메탈BOT"

ROLES = [ROLE_NONE, ROLE_BOM, ROLE_SOURCE, ROLE_METAL_TOP, ROLE_METAL_BOT]


def run_gui() -> None:
    app = ReviewApp()
    app.mainloop()


class ReviewApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(f"BOM 검토 — {__version__}")
        self.geometry("920x640")
        self.minsize(800, 520)

        self._folder: Path | None = None
        self._paths: list[Path] = []
        self._role_by_key: dict[str, str] = {}

        self._bom_headers: list[str] = []
        self._src_headers: list[str] = []

        top = ttk.Frame(self, padding=8)
        top.pack(fill=tk.X)

        ttk.Label(top, text="작업 폴더:").pack(side=tk.LEFT)
        self.var_folder = tk.StringVar(value="(선택 안 함)")
        ttk.Entry(top, textvariable=self.var_folder, width=70, state="readonly").pack(
            side=tk.LEFT, padx=6, fill=tk.X, expand=True
        )
        ttk.Button(top, text="폴더 선택…", command=self._pick_folder).pack(side=tk.LEFT)

        mid = ttk.Frame(self, padding=(8, 0))
        mid.pack(fill=tk.BOTH, expand=True)

        left = ttk.LabelFrame(mid, text="파일 목록 (더블클릭: 역할 지정)", padding=6)
        left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.list_files = tk.Listbox(left, height=18, exportselection=False)
        self.list_files.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb = ttk.Scrollbar(left, orient=tk.VERTICAL, command=self.list_files.yview)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.list_files.config(yscrollcommand=sb.set)
        self.list_files.bind("<Double-Button-1>", self._on_file_double_click)

        right = ttk.Frame(mid, padding=(8, 0))
        right.pack(side=tk.LEFT, fill=tk.Y)

        self.lbl_bom = ttk.Label(right, text="BOM: —")
        self.lbl_bom.pack(anchor=tk.W, pady=2)
        self.lbl_src = ttk.Label(right, text="원본좌표: —")
        self.lbl_src.pack(anchor=tk.W, pady=2)

        ttk.Separator(right, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=8)

        ttk.Label(right, text="BOM 좌표명 구분자").pack(anchor=tk.W)
        self.var_delim = tk.StringVar(value=", ")
        ttk.Entry(right, textvariable=self.var_delim, width=16).pack(anchor=tk.W, pady=2)

        ttk.Label(right, text="BOM 좌표명 열 (헤더)").pack(anchor=tk.W, pady=(8, 0))
        self.combo_bom_ref = ttk.Combobox(right, width=28, state="readonly")
        self.combo_bom_ref.pack(anchor=tk.W)

        ttk.Label(right, text="BOM 수량 열 (선택, 비우면 검사 생략)").pack(anchor=tk.W, pady=(8, 0))
        self.combo_bom_qty = ttk.Combobox(right, width=28, state="readonly")
        self.combo_bom_qty.pack(anchor=tk.W)

        ttk.Label(right, text="원본 좌표명 열 (헤더)").pack(anchor=tk.W, pady=(8, 0))
        self.combo_src_ref = ttk.Combobox(right, width=28, state="readonly")
        self.combo_src_ref.pack(anchor=tk.W)

        ttk.Button(right, text="헤더 다시 읽기", command=self._refresh_headers).pack(
            anchor=tk.W, pady=10
        )
        ttk.Button(right, text="검토 실행", command=self._run_review).pack(anchor=tk.W)

        bot = ttk.LabelFrame(self, text="결과", padding=6)
        bot.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))

        self.txt = tk.Text(bot, height=14, wrap=tk.WORD, state=tk.DISABLED)
        ys = ttk.Scrollbar(bot, orient=tk.VERTICAL, command=self.txt.yview)
        self.txt.config(yscrollcommand=ys.set)
        self.txt.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ys.pack(side=tk.RIGHT, fill=tk.Y)

        self._log(
            "사용 순서:\n"
            "1) 작업 폴더 선택\n"
            "2) 파일 더블클릭 → BOM / 원본좌표 역할 지정\n"
            "3) 열 이름 선택 후 「검토 실행」\n"
            "(Excel은 첫 번째 시트만 사용합니다.)\n"
        )

    def _path_key(self, p: Path) -> str:
        return str(p.resolve())

    def _pick_folder(self) -> None:
        d = filedialog.askdirectory(title="작업 폴더 선택")
        if not d:
            return
        self._folder = Path(d)
        self.var_folder.set(str(self._folder))
        self._paths = list_files_in_folder(self._folder)
        self._role_by_key.clear()
        self.list_files.delete(0, tk.END)
        for p in self._paths:
            self.list_files.insert(tk.END, p.name)
        self._update_role_labels()
        self._clear_combos()
        self._log(f"폴더: {self._folder}\n파일 {len(self._paths)}개\n")

    def _selected_path(self) -> Path | None:
        sel = self.list_files.curselection()
        if not sel:
            return None
        i = int(sel[0])
        if 0 <= i < len(self._paths):
            return self._paths[i]
        return None

    def _on_file_double_click(self, _evt: tk.Event) -> None:  # noqa: ANN401
        p = self._selected_path()
        if p is None:
            return
        key = self._path_key(p)
        current = self._role_by_key.get(key, ROLE_NONE)

        dlg = tk.Toplevel(self)
        dlg.title("역할 선택")
        dlg.transient(self)
        dlg.grab_set()
        ttk.Label(dlg, text=p.name, wraplength=400).pack(padx=12, pady=8)

        var = tk.StringVar(value=current if current in ROLES else ROLE_NONE)
        for r in ROLES:
            if r == ROLE_NONE:
                ttk.Radiobutton(dlg, text="(지정 안 함)", variable=var, value=r).pack(anchor=tk.W, padx=12)
            else:
                ttk.Radiobutton(dlg, text=r, variable=var, value=r).pack(anchor=tk.W, padx=12)

        def ok() -> None:
            v = var.get()
            if v == ROLE_NONE:
                self._role_by_key.pop(key, None)
            else:
                for k, rv in list(self._role_by_key.items()):
                    if rv == v and k != key:
                        self._role_by_key.pop(k, None)
                self._role_by_key[key] = v
            dlg.destroy()
            self._update_role_labels()
            self._refresh_headers()

        bf = ttk.Frame(dlg, padding=8)
        bf.pack(fill=tk.X)
        ttk.Button(bf, text="확인", command=ok).pack(side=tk.RIGHT)

    def _path_for_role(self, role: str) -> Path | None:
        for p in self._paths:
            if self._role_by_key.get(self._path_key(p)) == role:
                return p
        return None

    def _update_role_labels(self) -> None:
        b = self._path_for_role(ROLE_BOM)
        s = self._path_for_role(ROLE_SOURCE)
        self.lbl_bom.config(text=f"BOM: {b.name if b else '—'}")
        self.lbl_src.config(text=f"원본좌표: {s.name if s else '—'}")

    def _clear_combos(self) -> None:
        self._bom_headers = []
        self._src_headers = []
        for c in (self.combo_bom_ref, self.combo_bom_qty, self.combo_src_ref):
            c.set("")
            c.configure(values=[])

    def _refresh_headers(self) -> None:
        bom = self._path_for_role(ROLE_BOM)
        src = self._path_for_role(ROLE_SOURCE)
        if bom is None or src is None:
            self._clear_combos()
            return
        try:
            self._bom_headers, _ = load_header_and_rows(bom, sheet_index=0, max_data_rows=0)
            self._src_headers, _ = load_header_and_rows(src, sheet_index=0, max_data_rows=0)
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("헤더 읽기 실패", str(e))
            return
        self.combo_bom_ref.configure(values=self._bom_headers)
        self.combo_bom_qty.configure(values=["(없음)"] + self._bom_headers)
        self.combo_bom_qty.set("(없음)")
        self.combo_src_ref.configure(values=self._src_headers)
        if self._bom_headers:
            self.combo_bom_ref.set(self._bom_headers[0])
        if self._src_headers:
            self.combo_src_ref.set(self._src_headers[0])
        self._log("헤더를 불러왔습니다. 열 이름을 확인하세요.\n")

    def _append_text(self, s: str) -> None:
        self.txt.config(state=tk.NORMAL)
        self.txt.insert(tk.END, s)
        self.txt.see(tk.END)
        self.txt.config(state=tk.DISABLED)

    def _log(self, s: str) -> None:
        self.txt.config(state=tk.NORMAL)
        self.txt.delete("1.0", tk.END)
        self.txt.insert("1.0", s)
        self.txt.config(state=tk.DISABLED)

    def _run_review(self) -> None:
        bom_p = self._path_for_role(ROLE_BOM)
        src_p = self._path_for_role(ROLE_SOURCE)
        if bom_p is None or src_p is None:
            messagebox.showwarning("확인", "BOM 파일과 원본좌표 파일 역할을 모두 지정하세요.")
            return
        bom_col = self.combo_bom_ref.get().strip()
        src_col = self.combo_src_ref.get().strip()
        if not bom_col or not src_col:
            messagebox.showwarning("확인", "BOM·원본의 좌표명 열을 선택하세요. 「헤더 다시 읽기」를 눌러 주세요.")
            return
        delim = self.var_delim.get()
        if delim == "":
            delim = ", "

        qf: list = []
        try:
            bh, br = load_header_and_rows(bom_p, sheet_index=0, max_data_rows=None)
            sh, sr = load_header_and_rows(src_p, sheet_index=0, max_data_rows=None)
            bom_cells = values_for_column(bh, br, bom_col)
            src_cells = values_for_column(sh, sr, src_col)
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("파일 읽기 실패", str(e))
            return

        bom_refs_flat: list[str] = []
        for v in bom_cells:
            bom_refs_flat.extend(split_designators(v, delimiter=delim))

        src_refs: list[str] = []
        for v in src_cells:
            if v is None:
                continue
            t = str(v).strip()
            if t:
                src_refs.append(t)

        lines: list[str] = []
        lines.append(f"=== 검토 결과 — {bom_p.name} ↔ {src_p.name} ===\n")
        lines.append(f"BOM에서 펼친 Reference 수: {len(bom_refs_flat)} (행 수 {len(bom_cells)})\n")
        lines.append(f"원본 Reference 행 수: {len(src_refs)}\n\n")

        qty_choice = self.combo_bom_qty.get().strip()
        if qty_choice and qty_choice != "(없음)":
            try:
                qty_vals = values_for_column(bh, br, qty_choice)
                qf = bom_quantity_mismatch_findings(
                    bom_cells, qty_vals, delimiter=delim
                )
            except KeyError as e:
                messagebox.showerror("열 오류", str(e))
                return
            lines.append("--- BOM 수량 vs 좌표명 개수 ---\n")
            if qf:
                for f in qf:
                    lines.append(f"[오류] {f.code} {f.detail}\n")
            else:
                lines.append("(불일치 없음)\n")
            lines.append("\n")

        has_qty_err = bool(qf)

        dup_bom = duplicate_reference_findings(bom_refs_flat, scope_label="BOM(펼친 목록)")
        dup_src = duplicate_reference_findings(src_refs, scope_label="원본좌표")
        lines.append("--- 좌표명 중복 ---\n")
        if dup_bom:
            for f in dup_bom:
                lines.append(f"[오류] BOM {f.reference}: {f.detail}\n")
        if dup_src:
            for f in dup_src:
                lines.append(f"[오류] 원본 {f.reference}: {f.detail}\n")
        if not dup_bom and not dup_src:
            lines.append("(중복 없음)\n")
        lines.append("\n")

        report = bom_vs_source_findings(bom_refs_flat, src_refs)
        lines.append("--- BOM ↔ 원본 매칭 ---\n")
        lines.append("[오류] BOM에만 있고 원본에 없음\n")
        errs = list(iter_error_findings(report))
        if errs:
            for f in errs:
                lines.append(f"  - {f.reference}\n")
        else:
            lines.append("  (없음)\n")
        lines.append("\n[참고] 원본에만 있음 (오류 아님)\n")
        infos = list(iter_info_findings(report))
        if infos:
            for f in infos:
                lines.append(f"  - {f.reference}\n")
        else:
            lines.append("  (없음)\n")

        lines.append("\n")
        if errs or dup_bom or dup_src:
            lines.append("요약: 오류 항목이 있습니다. 위 내용을 확인하세요.\n")
        elif has_qty_err:
            lines.append("요약: 수량 불일치가 있습니다.\n")
        else:
            lines.append("요약: 치명 오류는 없습니다 (참고 항목은 정책상 허용).\n")

        self._log("".join(lines))
