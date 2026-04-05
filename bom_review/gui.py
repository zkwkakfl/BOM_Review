"""작업 폴더·파일 역할·열 매핑 후 실제 파일로 BOM↔원본 검토 (Tkinter)."""

from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Any

from bom_review._version import __version__
from bom_review.bom_parse import split_designators
from bom_review.excel_com import is_excel_path
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
        self.title(f"BOM 검토 정식  v{__version__}")
        self.geometry("920x660")
        self.minsize(800, 540)

        self._folder: Path | None = None
        self._paths: list[Path] = []
        self._role_by_key: dict[str, str] = {}

        self._bom_headers: list[str] = []
        self._src_headers: list[str] = []
        # Excel에서 잡은 범위(첫 행=헤더). None이면 파일 전체(첫 시트) 읽기
        self._bom_table_override: tuple[list[str], list[list[Any]]] | None = None
        self._bom_override_key: str | None = None
        self._src_table_override: tuple[list[str], list[list[Any]]] | None = None
        self._src_override_key: str | None = None

        self._build_menubar()

        head = ttk.Frame(self, padding=(8, 6, 8, 0))
        head.pack(fill=tk.X)
        ttk.Label(
            head,
            text="BOM · 원본좌표 매칭 검토 (정식)",
            font=("Segoe UI", 12, "bold"),
        ).pack(anchor=tk.W)
        ttk.Label(
            head,
            text="작업 폴더 안의 CSV / Excel(xlsx·xlsm·xlsb). BOM은 필수, 원본·메탈 역할은 선택.",
            wraplength=880,
        ).pack(anchor=tk.W, pady=(2, 0))

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

        left = ttk.LabelFrame(
            mid,
            text="파일 목록 (더블클릭: 역할 → Excel이면 범위 지정 여부)",
            padding=6,
        )
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

        ttk.Label(
            right,
            text="원본 좌표명 열 (선택 — 원본 파일 지정 시에만)",
        ).pack(anchor=tk.W, pady=(8, 0))
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

        self._status = ttk.Label(
            self,
            text=f"정식 v{__version__}  |  Excel: 역할 지정 시 COM으로 범위 선택 가능  |  CSV는 파일·열 선택",
            anchor=tk.W,
            relief=tk.SUNKEN,
            padding=(6, 2),
        )
        self._status.pack(side=tk.BOTTOM, fill=tk.X)

        self._log(self._welcome_text())

    def _build_menubar(self) -> None:
        menubar = tk.Menu(self)

        m_file = tk.Menu(menubar, tearoff=0)
        m_file.add_command(label="작업 폴더 선택…", command=self._pick_folder)
        m_file.add_separator()
        m_file.add_command(label="종료", command=self.destroy)
        menubar.add_cascade(label="파일", menu=m_file)

        m_help = tk.Menu(menubar, tearoff=0)
        m_help.add_command(label="사용 안내", command=self._show_usage)
        m_help.add_command(label="버전 정보", command=self._show_about)
        menubar.add_cascade(label="도움말", menu=m_help)

        self.config(menu=menubar)

    def _welcome_text(self) -> str:
        return (
            "【정식 사용 순서】\n\n"
            "1. 「폴더 선택」 또는 메뉴 파일 → 작업 폴더 선택\n"
            "2. BOM 파일은 역할 필수. 원본·메탈은 선택.\n"
            "3. Excel(xlsx 등)에서 BOM/원본 역할을 고르면 Excel이 열려 드래그·Shift+방향키로 범위를 잡을 수 있습니다. "
            "「아니오」면 첫 시트 전체 + 콤보.\n"
            "4. 「헤더 다시 읽기」는 파일 기준으로 다시 읽으며, Excel로 잡은 범위는 해제됩니다.\n"
            "5. 「검토 실행」 — 원본이 없으면 BOM 수량·중복만 검토합니다.\n\n"
            "※ 원본이 있을 때: 원본이 기준이며, BOM에만 있는 Reference는 오류, 원본에만 있는 항목은 참고(오류 아님).\n"
        )

    def _show_usage(self) -> None:
        messagebox.showinfo("사용 안내", self._welcome_text())

    def _show_about(self) -> None:
        messagebox.showinfo(
            "버전 정보",
            f"BOM 검토 정식\n\n버전: {__version__}\n\n"
            "사용 중 불편·추가 기능은 이슈나 내부 요청으로 알려 주세요.",
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
        self._clear_overrides_and_combos()
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
            if v in (ROLE_BOM, ROLE_SOURCE) and is_excel_path(p):
                ask = messagebox.askyesno(
                    "Excel 범위 지정",
                    f"{p.name}\n\n"
                    "Excel을 열고 마우스 드래그·Shift+방향키로 범위를 선택한 뒤 적용할까요?\n\n"
                    "「아니오」는 첫 시트 전체를 읽어 콤보박스로 열 고르기(기존 방식)입니다.",
                    parent=self,
                )
                if ask:
                    self._start_excel_range_pick(p, v)
                    return
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

    def _clear_overrides_and_combos(self) -> None:
        self._bom_table_override = None
        self._bom_override_key = None
        self._src_table_override = None
        self._src_override_key = None
        self._clear_combos()

    def _clear_combos(self) -> None:
        self._bom_headers = []
        self._src_headers = []
        for c in (self.combo_bom_ref, self.combo_bom_qty):
            c.set("")
            c.configure(values=[])
        self.combo_src_ref.set("")
        self.combo_src_ref.configure(values=[])
        self.combo_src_ref.configure(state="disabled")

    def _start_excel_range_pick(self, path: Path, role: str) -> None:
        from bom_review.excel_range_dialog import ExcelRangeDialog

        def on_ok(headers: list[str], data: list[list[Any]]) -> None:
            self._apply_excel_table(role, path, headers, data)
            self._fill_other_combos_after_excel(role)

        def on_cancel() -> None:
            self._refresh_headers()

        ExcelRangeDialog(self, path, on_ok=on_ok, on_cancel=on_cancel)

    def _apply_excel_table(
        self,
        role: str,
        path: Path,
        headers: list[str],
        data_rows: list[list[Any]],
    ) -> None:
        key = self._path_key(path)
        if role == ROLE_BOM:
            self._bom_table_override = (headers, data_rows)
            self._bom_override_key = key
            self._bom_headers = headers
            self.combo_bom_ref.configure(values=headers, state="readonly")
            self.combo_bom_qty.configure(values=["(없음)"] + headers)
            self.combo_bom_qty.set("(없음)")
            if headers:
                self.combo_bom_ref.set(headers[0])
        elif role == ROLE_SOURCE:
            self._src_table_override = (headers, data_rows)
            self._src_override_key = key
            self._src_headers = headers
            self.combo_src_ref.configure(values=headers, state="readonly")
            if headers:
                self.combo_src_ref.set(headers[0])
        self._log(
            f"Excel 범위 적용: {path.name} ({role})\n"
            "첫 행을 헤더로 썼습니다. 좌표명·수량 열을 확인하세요.\n"
        )

    def _fill_other_combos_after_excel(self, edited_role: str) -> None:
        if edited_role == ROLE_BOM:
            self._load_src_combos_from_file_if_needed()
        elif edited_role == ROLE_SOURCE:
            self._load_bom_combos_from_file_if_needed()

    def _load_src_combos_from_file_if_needed(self) -> None:
        src = self._path_for_role(ROLE_SOURCE)
        if src is None:
            self.combo_src_ref.set("")
            self.combo_src_ref.configure(values=[], state="disabled")
            self._src_headers = []
            return
        sk = self._path_key(src)
        if self._src_override_key == sk and self._src_table_override is not None:
            return
        try:
            h, _ = load_header_and_rows(src, sheet_index=0, max_data_rows=0)
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("원본 헤더", str(e))
            return
        self._src_headers = h
        self.combo_src_ref.configure(values=h, state="readonly")
        if h:
            self.combo_src_ref.set(h[0])

    def _load_bom_combos_from_file_if_needed(self) -> None:
        bom = self._path_for_role(ROLE_BOM)
        if bom is None:
            return
        bk = self._path_key(bom)
        if self._bom_override_key == bk and self._bom_table_override is not None:
            return
        try:
            h, _ = load_header_and_rows(bom, sheet_index=0, max_data_rows=0)
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("BOM 헤더", str(e))
            return
        self._bom_headers = h
        self.combo_bom_ref.configure(values=h, state="readonly")
        self.combo_bom_qty.configure(values=["(없음)"] + h)
        self.combo_bom_qty.set("(없음)")
        if h:
            self.combo_bom_ref.set(h[0])

    def _refresh_headers(self) -> None:
        self._bom_table_override = None
        self._bom_override_key = None
        self._src_table_override = None
        self._src_override_key = None
        bom = self._path_for_role(ROLE_BOM)
        src = self._path_for_role(ROLE_SOURCE)
        if bom is None:
            self._clear_combos()
            return
        try:
            self._bom_headers, _ = load_header_and_rows(bom, sheet_index=0, max_data_rows=0)
            if src is not None:
                self._src_headers, _ = load_header_and_rows(src, sheet_index=0, max_data_rows=0)
            else:
                self._src_headers = []
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("헤더 읽기 실패", str(e))
            return
        self.combo_bom_ref.configure(values=self._bom_headers, state="readonly")
        self.combo_bom_qty.configure(values=["(없음)"] + self._bom_headers)
        self.combo_bom_qty.set("(없음)")
        if self._bom_headers:
            self.combo_bom_ref.set(self._bom_headers[0])
        if src is not None and self._src_headers:
            self.combo_src_ref.configure(values=self._src_headers, state="readonly")
            self.combo_src_ref.set(self._src_headers[0])
        else:
            self.combo_src_ref.set("")
            self.combo_src_ref.configure(values=[], state="disabled")
        msg = "헤더를 불러왔습니다. 열 이름을 확인하세요.\n"
        if src is None:
            msg += "(원본 파일이 없어 BOM↔원본 매칭은 실행하지 않습니다.)\n"
        self._log(msg)

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
        if bom_p is None:
            messagebox.showwarning("확인", "BOM 파일 역할을 지정하세요.")
            return
        bom_col = self.combo_bom_ref.get().strip()
        if not bom_col:
            messagebox.showwarning(
                "확인",
                "BOM 좌표명 열을 선택하세요. 「헤더 다시 읽기」를 눌러 주세요.",
            )
            return
        src_col = self.combo_src_ref.get().strip()
        if src_p is not None and not src_col:
            messagebox.showwarning(
                "확인",
                "원본 파일을 지정했습니다. 원본 좌표명 열을 선택하세요. 「헤더 다시 읽기」를 눌러 주세요.",
            )
            return
        delim = self.var_delim.get()
        if delim == "":
            delim = ", "

        qf: list = []
        try:
            if (
                self._bom_table_override is not None
                and self._bom_override_key == self._path_key(bom_p)
            ):
                bh, br = self._bom_table_override
            else:
                bh, br = load_header_and_rows(bom_p, sheet_index=0, max_data_rows=None)
            bom_cells = values_for_column(bh, br, bom_col)
            sh: list[str] = []
            sr: list[list[Any]] = []
            src_cells: list[Any] = []
            if src_p is not None:
                if (
                    self._src_table_override is not None
                    and self._src_override_key == self._path_key(src_p)
                ):
                    sh, sr = self._src_table_override
                else:
                    sh, sr = load_header_and_rows(src_p, sheet_index=0, max_data_rows=None)
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
        if src_p is not None:
            lines.append(f"=== 검토 결과 — {bom_p.name} ↔ {src_p.name} ===\n")
        else:
            lines.append(f"=== 검토 결과 — {bom_p.name} (원본 미제공, BOM만) ===\n")
        lines.append(f"BOM에서 펼친 Reference 수: {len(bom_refs_flat)} (행 수 {len(bom_cells)})\n")
        if src_p is not None:
            lines.append(f"원본 Reference 행 수: {len(src_refs)}\n\n")
        else:
            lines.append("원본: (파일 없음 — 매칭·원본 중복 검사 생략)\n\n")

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
        dup_src: list = []
        lines.append("--- 좌표명 중복 ---\n")
        if dup_bom:
            for f in dup_bom:
                lines.append(f"[오류] BOM {f.reference}: {f.detail}\n")
        if src_p is not None:
            dup_src = duplicate_reference_findings(src_refs, scope_label="원본좌표")
            if dup_src:
                for f in dup_src:
                    lines.append(f"[오류] 원본 {f.reference}: {f.detail}\n")
        if not dup_bom and not dup_src:
            lines.append("(중복 없음)\n")
        lines.append("\n")

        errs: list = []
        if src_p is not None:
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
        else:
            lines.append("--- BOM ↔ 원본 매칭 ---\n")
            lines.append("(원본 파일이 없어 생략)\n\n")

        lines.append("\n")
        if src_p is None:
            if dup_bom or has_qty_err:
                lines.append(
                    "요약: BOM 검토 중 오류가 있습니다. 원본은 없어 매칭은 하지 않았습니다.\n"
                )
            else:
                lines.append(
                    "요약: BOM만 검토했으며 (수량·중복 기준) 치명 오류는 없습니다. 원본 매칭은 생략.\n"
                )
        elif errs or dup_bom or dup_src:
            lines.append("요약: 오류 항목이 있습니다. 위 내용을 확인하세요.\n")
        elif has_qty_err:
            lines.append("요약: 수량 불일치가 있습니다.\n")
        else:
            lines.append("요약: 치명 오류는 없습니다 (참고 항목은 정책상 허용).\n")

        self._log("".join(lines))
