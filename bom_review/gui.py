"""작업 폴더·파일 역할·열 매핑 후 실제 파일로 BOM↔원본 검토 (Tkinter)."""

from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Any

from bom_review._version import __version__
from bom_review.bom_parse import split_designators
from bom_review.excel_com import SelectionSourceMeta, is_excel_path
from bom_review.excel_snapshot import new_snapshot_workbook_path, write_role_range_to_snapshot
from bom_review.bom_qty import bom_quantity_mismatch_findings
from bom_review.matching import (
    bom_vs_source_findings,
    duplicate_reference_findings,
    iter_error_findings,
    iter_info_findings,
)
from bom_review.table_io import (
    list_files_in_folder,
    load_header_and_rows,
    load_header_and_rows_by_sheet_name,
    resolve_column_index,
    values_for_column,
)

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
        self.geometry("980x720")
        self.minsize(860, 600)

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
        # Excel 범위 적용 시 생성되는 타임스탬프 결과 통합문서(복사본 + Range_Set)
        self._snapshot_workbook: Path | None = None
        self._bom_snapshot_sheet: str | None = None
        self._src_snapshot_sheet: str | None = None

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
        self.lbl_bom.pack(anchor=tk.W, pady=(0, 2))
        self.lbl_src = ttk.Label(right, text="원본좌표: —")
        self.lbl_src.pack(anchor=tk.W, pady=(0, 4))

        map_row = ttk.Frame(right)
        map_row.pack(fill=tk.BOTH, expand=True)

        bom_lf = ttk.LabelFrame(map_row, text="BOM 테이블 (자재명·좌표명·수량 필수)", padding=6)
        bom_lf.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 4))
        src_lf = ttk.LabelFrame(map_row, text="원본 테이블 — 열 매핑", padding=6)
        src_lf.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(4, 0))

        self.var_delim = tk.StringVar(value=", ")
        self.combo_bom_material = ttk.Combobox(bom_lf, width=22, state="readonly")
        self.combo_bom_ref = ttk.Combobox(bom_lf, width=22, state="readonly")
        self.combo_bom_qty = ttk.Combobox(bom_lf, width=22, state="readonly")
        self.combo_bom_mount = ttk.Combobox(bom_lf, width=22, state="readonly")
        self.combo_src_ref = ttk.Combobox(src_lf, width=22, state="readonly")
        self.combo_src_x = ttk.Combobox(src_lf, width=22, state="readonly")
        self.combo_src_y = ttk.Combobox(src_lf, width=22, state="readonly")
        self.combo_src_layer = ttk.Combobox(src_lf, width=22, state="readonly")

        bom_rows: list[tuple[str, tk.Widget]] = [
            (
                "자재명 열 (필수)\n품번·MPN·부품명 등",
                self.combo_bom_material,
            ),
            (
                "좌표명 열 (필수)\n매칭·중복·펼침 기준 (Reference)",
                self.combo_bom_ref,
            ),
            (
                "좌표명 구분자 (필수)\n좌표명 셀 안에서 토큰 나눔",
                ttk.Entry(bom_lf, textvariable=self.var_delim, width=18),
            ),
            (
                "수량 열 (필수)\n행별 수량 ↔ 같은 행 좌표명 펼친 개수 비교",
                self.combo_bom_qty,
            ),
            (
                "마운트 타입 (선택)\nSMD/THT 등",
                self.combo_bom_mount,
            ),
        ]
        for r, (text, w) in enumerate(bom_rows):
            ttk.Label(bom_lf, text=text, justify=tk.LEFT).grid(
                row=r, column=0, sticky=tk.NW, padx=(0, 8), pady=3
            )
            w.grid(row=r, column=1, sticky=tk.EW, pady=3)
        bom_lf.columnconfigure(1, weight=1)

        src_rows: list[tuple[str, tk.Widget]] = [
            (
                "좌표명 열 (원본, 필수)\nBOM 좌표명과 1:1 매칭",
                self.combo_src_ref,
            ),
            ("X좌표 열 (선택)", self.combo_src_x),
            ("Y좌표 열 (선택)", self.combo_src_y),
            ("Layer 열 (선택)\nTOP/BOT 등", self.combo_src_layer),
        ]
        for r, (text, w) in enumerate(src_rows):
            ttk.Label(src_lf, text=text, justify=tk.LEFT).grid(
                row=r, column=0, sticky=tk.NW, padx=(0, 8), pady=3
            )
            w.grid(row=r, column=1, sticky=tk.EW, pady=3)
        src_lf.columnconfigure(1, weight=1)

        btn_row = ttk.Frame(right)
        btn_row.pack(fill=tk.X, pady=(8, 0))
        ttk.Button(btn_row, text="헤더 다시 읽기", command=self._refresh_headers).pack(
            side=tk.LEFT, padx=(0, 6)
        )
        ttk.Button(btn_row, text="검토 실행", command=self._run_review).pack(side=tk.LEFT)

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
            "3. Excel(xlsx 등)에서 BOM/원본 역할을 고르면 Excel이 열려 범위를 잡을 수 있습니다. "
            "「예」면 작업 폴더에 타임스탬프 이름의 검토용 통합문서가 생기고, 선택 범위가 시트로 복사되며 "
            "Range_Set 시트에 원본 파일·시트·주소가 기록됩니다. 검토 실행 시 그 복사본을 읽습니다. "
            "「아니오」는 첫 시트 전체 + 콤보(기존 방식).\n"
            "4. 「헤더 다시 읽기」는 파일 기준으로 다시 읽으며, Excel 복사본·검토용 통합문서 연동도 해제됩니다.\n"
            "5. 「검토 실행」 — 원본이 없으면 BOM 수량·중복만 검토합니다.\n\n"
            "※ 원본이 있을 때: 원본이 기준이며, BOM에만 있는 Reference는 오류, 원본에만 있는 항목은 참고(오류 아님).\n"
            "※ BOM은 자재명·좌표명·수량 열이 필수입니다. 마운트 타입·원본 X/Y/Layer는 선택입니다.\n"
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
        self._snapshot_workbook = None
        self._bom_snapshot_sheet = None
        self._src_snapshot_sheet = None
        self._clear_combos()

    def _clear_combos(self) -> None:
        self._bom_headers = []
        self._src_headers = []
        self._configure_bom_combos([])
        self._configure_source_combos([])

    def _configure_bom_combos(self, headers: list[str]) -> None:
        """BOM 쪽 열 콤보를 채운다. 자재명·좌표명·수량은 필수(콤보에 '(없음)' 없음). 빈 목록이면 비활성."""
        if not headers:
            for c in (self.combo_bom_material, self.combo_bom_ref, self.combo_bom_qty):
                c.set("")
                c.configure(values=[], state="disabled")
            self.combo_bom_mount.set("(없음)")
            self.combo_bom_mount.configure(values=[], state="disabled")
            return
        opt_mount = ["(없음)"] + headers
        for c in (self.combo_bom_material, self.combo_bom_ref, self.combo_bom_qty):
            c.configure(values=list(headers), state="readonly")
        n = len(headers)
        if n >= 3:
            self.combo_bom_material.set(headers[0])
            self.combo_bom_ref.set(headers[1])
            self.combo_bom_qty.set(headers[2])
        elif n == 2:
            self.combo_bom_material.set(headers[0])
            self.combo_bom_ref.set(headers[1])
            self.combo_bom_qty.set(headers[1])
        else:
            h0 = headers[0]
            self.combo_bom_material.set(h0)
            self.combo_bom_ref.set(h0)
            self.combo_bom_qty.set(h0)
        self.combo_bom_mount.configure(values=opt_mount, state="readonly")
        self.combo_bom_mount.set("(없음)")

    @staticmethod
    def _sync_bom_combos_to_table_headers(self, headers: list[str]) -> None:
        """
        검토 시 실제 로드된 헤더(bh)에 맞게 콤보 표시값을 맞춘다.
        Tk readonly 콤보·미세 공백·이전 세션 문자열로 인한 '열 없음'·깨진 표시를 줄인다.
        """
        if not headers:
            return
        for combo in (self.combo_bom_material, self.combo_bom_ref, self.combo_bom_qty):
            v = combo.get().strip()
            if not v:
                continue
            try:
                idx = resolve_column_index(headers, v)
                combo.set(headers[idx])
            except KeyError:
                pass
        vm = self.combo_bom_mount.get().strip()
        if vm and vm != "(없음)":
            try:
                idx = resolve_column_index(headers, vm)
                self.combo_bom_mount.set(headers[idx])
            except KeyError:
                pass

    def _sync_src_combos_to_table_headers(self, headers: list[str]) -> None:
        if not headers:
            return
        vr = self.combo_src_ref.get().strip()
        if vr:
            try:
                idx = resolve_column_index(headers, vr)
                self.combo_src_ref.set(headers[idx])
            except KeyError:
                pass
        for combo in (self.combo_src_x, self.combo_src_y, self.combo_src_layer):
            v = combo.get().strip()
            if v and v != "(없음)":
                try:
                    idx = resolve_column_index(headers, v)
                    combo.set(headers[idx])
                except KeyError:
                    pass

    @staticmethod
    def _count_nonempty_ref_but_empty_aux(ref_cells: list[Any], aux: list[Any]) -> int:
        """기준 열(셀)은 비어 있지 않은데 보조 열이 비어 있는 행 개수. 같은 행 순서 전제."""
        n = min(len(ref_cells), len(aux))
        c = 0
        for i in range(n):
            rv = ref_cells[i]
            if rv is None or str(rv).strip() == "":
                continue
            av = aux[i]
            if av is None or (isinstance(av, str) and str(av).strip() == ""):
                c += 1
        return c

    def _configure_source_combos(self, headers: list[str]) -> None:
        """원본 쪽 4개 콤보를 헤더 목록에 맞게 채운다. 빈 목록이면 비활성."""
        if not headers:
            self.combo_src_ref.set("")
            self.combo_src_ref.configure(values=[], state="disabled")
            for c in (self.combo_src_x, self.combo_src_y, self.combo_src_layer):
                c.set("(없음)")
                c.configure(values=[], state="disabled")
            return
        opt = ["(없음)"] + headers
        self.combo_src_ref.configure(values=headers, state="readonly")
        self.combo_src_ref.set(headers[0])
        for c in (self.combo_src_x, self.combo_src_y, self.combo_src_layer):
            c.configure(values=opt, state="readonly")
            c.set("(없음)")

    def _start_excel_range_pick(self, path: Path, role: str) -> None:
        from bom_review.excel_range_dialog import ExcelRangeDialog

        def on_ok(
            headers: list[str],
            data: list[list[Any]],
            meta: SelectionSourceMeta,
        ) -> None:
            try:
                dest_sheet = self._persist_excel_snapshot(role, path, headers, data, meta)
            except Exception as e:  # noqa: BLE001
                messagebox.showerror("검토용 통합문서", str(e), parent=self)
                on_cancel()
                return
            try:
                snap_path = self._snapshot_workbook
                if snap_path is None:
                    raise RuntimeError("검토용 통합문서 경로가 없습니다.")
                h2, r2 = load_header_and_rows_by_sheet_name(
                    snap_path,
                    sheet_name=dest_sheet,
                    max_data_rows=None,
                )
            except Exception as e:  # noqa: BLE001
                messagebox.showerror("복사본 읽기", str(e), parent=self)
                on_cancel()
                return
            self._apply_excel_table(role, path, h2, r2)
            self._fill_other_combos_after_excel(role)
            self._log(
                "Excel 범위를 검토용 통합문서에 반영했습니다.\n"
                f"· 원본: {path.name} ({role})\n"
                f"· 결과 파일: {self._snapshot_workbook}\n"
                f"· 데이터 시트: {dest_sheet}\n"
                f"· 원본 위치: {meta.source_sheet} / {meta.source_address}\n"
                "첫 행을 헤더로 썼습니다. BOM·원본 열 매핑을 확인하세요.\n"
            )

        def on_cancel() -> None:
            self._refresh_headers()

        ExcelRangeDialog(self, path, on_ok=on_ok, on_cancel=on_cancel)

    def _persist_excel_snapshot(
        self,
        role: str,
        path: Path,
        headers: list[str],
        data_rows: list[list[Any]],
        meta: SelectionSourceMeta,
    ) -> str:
        """타임스탬프 결과 통합문서에 복사하고 Range_Set을 갱신한다. 복사 시트 이름을 반환."""
        folder = self._folder.resolve() if self._folder is not None else path.parent.resolve()
        if self._snapshot_workbook is None:
            self._snapshot_workbook = new_snapshot_workbook_path(folder)
        bom_coord: str | None = None
        if role == ROLE_BOM and headers:
            picked = self.combo_bom_ref.get().strip()
            if picked and picked in headers:
                bom_coord = picked
            else:
                bom_coord = headers[0]
        dest = write_role_range_to_snapshot(
            self._snapshot_workbook,
            role=role,
            headers=headers,
            data_rows=data_rows,
            meta=meta,
            create_new_workbook=True,
            bom_coord_column=bom_coord,
        )
        if role == ROLE_BOM:
            self._bom_snapshot_sheet = dest
        elif role == ROLE_SOURCE:
            self._src_snapshot_sheet = dest
        return dest

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
            self._configure_bom_combos(headers)
        elif role == ROLE_SOURCE:
            self._src_table_override = (headers, data_rows)
            self._src_override_key = key
            self._src_headers = headers
            self._configure_source_combos(headers)

    def _fill_other_combos_after_excel(self, edited_role: str) -> None:
        if edited_role == ROLE_BOM:
            self._load_src_combos_from_file_if_needed()
        elif edited_role == ROLE_SOURCE:
            self._load_bom_combos_from_file_if_needed()

    def _load_src_combos_from_file_if_needed(self) -> None:
        src = self._path_for_role(ROLE_SOURCE)
        if src is None:
            self._src_headers = []
            self._configure_source_combos([])
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
        self._configure_source_combos(h)

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
        self._configure_bom_combos(h)

    def _refresh_headers(self) -> None:
        self._bom_table_override = None
        self._bom_override_key = None
        self._src_table_override = None
        self._src_override_key = None
        self._snapshot_workbook = None
        self._bom_snapshot_sheet = None
        self._src_snapshot_sheet = None
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
        self._configure_bom_combos(self._bom_headers)
        if src is not None and self._src_headers:
            self._configure_source_combos(self._src_headers)
        else:
            self._configure_source_combos([])
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
        bom_mat_choice = self.combo_bom_material.get().strip()
        if not bom_mat_choice:
            messagebox.showwarning(
                "확인",
                "BOM 자재명 열을 선택하세요. 「헤더 다시 읽기」를 눌러 주세요.",
            )
            return
        qty_choice = self.combo_bom_qty.get().strip()
        if not qty_choice:
            messagebox.showwarning(
                "확인",
                "BOM 수량 열을 선택하세요. 「헤더 다시 읽기」를 눌러 주세요.",
            )
            return
        delim = self.var_delim.get()
        if delim == "":
            delim = ", "

        qf: list = []
        bom_mat_vals: list[Any] = []
        bom_mount_vals: list[Any] | None = None
        bom_mount_choice = ""
        src_x_vals: list[Any] | None = None
        src_y_vals: list[Any] | None = None
        src_layer_vals: list[Any] | None = None
        src_x_choice = ""
        src_y_choice = ""
        src_layer_choice = ""
        try:
            if (
                self._bom_table_override is not None
                and self._bom_override_key == self._path_key(bom_p)
            ):
                if (
                    self._snapshot_workbook is not None
                    and self._bom_snapshot_sheet is not None
                ):
                    bh, br = load_header_and_rows_by_sheet_name(
                        self._snapshot_workbook,
                        sheet_name=self._bom_snapshot_sheet,
                        max_data_rows=None,
                    )
                else:
                    bh, br = self._bom_table_override
            else:
                bh, br = load_header_and_rows(bom_p, sheet_index=0, max_data_rows=None)
            self._sync_bom_combos_to_table_headers(bh)
            bom_col = self.combo_bom_ref.get().strip()
            bom_mat_choice = self.combo_bom_material.get().strip()
            qty_choice = self.combo_bom_qty.get().strip()
            if not bom_col or not bom_mat_choice or not qty_choice:
                messagebox.showerror(
                    "BOM 열 매핑",
                    "데이터를 읽은 뒤 열 이름이 비었습니다. 「헤더 다시 읽기」 후 다시 지정하세요.",
                )
                return
            try:
                resolve_column_index(bh, bom_col)
                resolve_column_index(bh, bom_mat_choice)
                resolve_column_index(bh, qty_choice)
            except KeyError as e:
                messagebox.showerror(
                    "BOM 열 오류",
                    f"{e}\n\n콤보에 보이는 이름과 실제 시트 헤더가 어긋났을 수 있습니다. "
                    "「헤더 다시 읽기」 후 자재명·좌표명·수량을 다시 선택해 보세요.",
                )
                return
            bom_cells = values_for_column(bh, br, bom_col)
            bom_mount_choice = self.combo_bom_mount.get().strip()
            try:
                bom_mat_vals = values_for_column(bh, br, bom_mat_choice)
                if bom_mount_choice and bom_mount_choice != "(없음)":
                    resolve_column_index(bh, bom_mount_choice)
                    bom_mount_vals = values_for_column(bh, br, bom_mount_choice)
            except KeyError as e:
                messagebox.showerror("BOM 열 오류", str(e))
                return
            sh: list[str] = []
            sr: list[list[Any]] = []
            src_cells: list[Any] = []
            if src_p is not None:
                if (
                    self._src_table_override is not None
                    and self._src_override_key == self._path_key(src_p)
                ):
                    if (
                        self._snapshot_workbook is not None
                        and self._src_snapshot_sheet is not None
                    ):
                        sh, sr = load_header_and_rows_by_sheet_name(
                            self._snapshot_workbook,
                            sheet_name=self._src_snapshot_sheet,
                            max_data_rows=None,
                        )
                    else:
                        sh, sr = self._src_table_override
                else:
                    sh, sr = load_header_and_rows(src_p, sheet_index=0, max_data_rows=None)
                self._sync_src_combos_to_table_headers(sh)
                src_col = self.combo_src_ref.get().strip()
                if not src_col:
                    messagebox.showerror(
                        "원본 열 매핑",
                        "원본 데이터를 읽은 뒤 좌표명 열이 비었습니다. 열을 다시 선택하세요.",
                    )
                    return
                try:
                    resolve_column_index(sh, src_col)
                except KeyError as e:
                    messagebox.showerror(
                        "원본 열 오류",
                        f"{e}\n\n「헤더 다시 읽기」 후 원본 좌표명 열을 다시 선택해 보세요.",
                    )
                    return
                src_cells = values_for_column(sh, sr, src_col)
                src_x_choice = self.combo_src_x.get().strip()
                src_y_choice = self.combo_src_y.get().strip()
                src_layer_choice = self.combo_src_layer.get().strip()
                try:
                    if src_x_choice and src_x_choice != "(없음)":
                        src_x_vals = values_for_column(sh, sr, src_x_choice)
                    if src_y_choice and src_y_choice != "(없음)":
                        src_y_vals = values_for_column(sh, sr, src_y_choice)
                    if src_layer_choice and src_layer_choice != "(없음)":
                        src_layer_vals = values_for_column(sh, sr, src_layer_choice)
                except KeyError as e:
                    messagebox.showerror("원본 열 오류", str(e))
                    return
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

        try:
            qty_vals = values_for_column(bh, br, qty_choice)
            qf = bom_quantity_mismatch_findings(
                bom_cells, qty_vals, delimiter=delim
            )
        except KeyError as e:
            messagebox.showerror("BOM 수량 열 오류", str(e))
            return
        lines.append("--- BOM 수량 vs 좌표명(펼친) 개수 ---\n")
        if qf:
            for f in qf:
                lines.append(f"[오류] {f.code} {f.detail}\n")
        else:
            lines.append("(불일치 없음)\n")
        lines.append("\n")

        lines.append("--- BOM 자재명·마운트 (행 정합성) ---\n")
        bad_m = self._count_nonempty_ref_but_empty_aux(bom_cells, bom_mat_vals)
        lines.append(
            f"  · 자재명 열 «{bom_mat_choice}»: "
            f"좌표명은 있는데 자재명이 비어 있는 행 수 = {bad_m}\n"
        )
        if bom_mount_vals is not None:
            bad_t = self._count_nonempty_ref_but_empty_aux(bom_cells, bom_mount_vals)
            lines.append(
                f"  · 마운트 타입 열 «{bom_mount_choice}»: "
                f"좌표명은 있는데 값이 비어 있는 행 수 = {bad_t}\n"
            )
        lines.append("\n")

        has_qty_err = bool(qf)

        dup_bom = duplicate_reference_findings(bom_refs_flat, scope_label="BOM(펼친 목록)")
        dup_src: list = []
        lines.append("--- 좌표명(Reference) 중복 ---\n")
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

        if src_p is not None and (
            src_x_vals is not None or src_y_vals is not None or src_layer_vals is not None
        ):
            lines.append("--- 원본 부가 열 (X / Y / Layer) ---\n")
            pairs: list[tuple[str, str, list[Any]]] = []
            if src_x_vals is not None:
                pairs.append(("X", src_x_choice, src_x_vals))
            if src_y_vals is not None:
                pairs.append(("Y", src_y_choice, src_y_vals))
            if src_layer_vals is not None:
                pairs.append(("Layer", src_layer_choice, src_layer_vals))
            for label, hdr, vals in pairs:
                bad = self._count_nonempty_ref_but_empty_aux(src_cells, vals)
                lines.append(
                    f"  · {label} 열 «{hdr}»: "
                    f"Reference는 있는데 본 열이 비어 있는 행 수 = {bad}\n"
                )
            lines.append("\n")

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
