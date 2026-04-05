"""Excel 2단계: (1) 원본에서 시트만 선택·결과 통합문서로 복사 (2) 결과 파일에서 검토 범위·헤더 선택."""

from __future__ import annotations

import tkinter as tk
from collections.abc import Callable
from pathlib import Path
from tkinter import messagebox, ttk
from typing import Any

from bom_review.excel_com import (
    SelectionSourceMeta,
    close_excel_quietly,
    open_workbook_in_new_excel,
    read_active_sheet_full_used_as_selection,
    read_full_sheet_and_review_selection,
)
from bom_review.excel_snapshot import ExcelParsedSelection

ExcelPersistCom = Callable[
    [Any, Any, Path, ExcelParsedSelection],
    tuple[SelectionSourceMeta, str],
]

ExcelSheetCopyOnDone = Callable[[SelectionSourceMeta, str], None]

ExcelRangeOnOk = Callable[
    [
        list[str],
        list[list[Any]],
        list[str],
        list[list[Any]],
        SelectionSourceMeta,
        int,
        int,
        int,
        int,
        int,
        int,
        str,
    ],
    None,
]


class ExcelSheetCopyDialog(tk.Toplevel):
    """
    1단계: 원본 통합문서를 연 뒤 복사할 시트만 활성화하고 「시트 복사」로 결과 통합문서에 붙인다.
    """

    def __init__(
        self,
        parent: tk.Misc,
        path: Path,
        *,
        persist_com: ExcelPersistCom,
        on_sheet_copied: ExcelSheetCopyOnDone,
        on_cancel: Callable[[], None],
    ) -> None:
        super().__init__(parent)
        self.title("Excel — 시트 복사 (1단계)")
        self.transient(parent)
        self.grab_set()
        self._persist_com = persist_com
        self._on_sheet_copied = on_sheet_copied
        self._on_cancel = on_cancel
        self._path = path
        self._xl: Any = None
        self._wb: Any = None
        self._closed = False

        ttk.Label(
            self,
            text=path.name,
            font=("Segoe UI", 10, "bold"),
        ).pack(padx=12, pady=(10, 4))
        ttk.Label(
            self,
            text=(
                "1) 복사할 시트 탭을 클릭해 활성화하세요.\n"
                "2) 「시트를 결과 파일로 복사」를 누르면 이 시트 전체가 검토용 통합문서에 붙고,\n"
                "   원본 통합문서는 닫힙니다.\n"
                "3) 이어서 열리는 창에서 검토 범위를 지정하면 그때 Range_Set 주소가 기록됩니다."
            ),
            justify=tk.LEFT,
        ).pack(padx=12, pady=4)

        bf = ttk.Frame(self, padding=12)
        bf.pack(fill=tk.X)
        ttk.Button(
            bf,
            text="시트를 결과 파일로 복사",
            command=self._apply,
        ).pack(side=tk.RIGHT, padx=4)
        ttk.Button(bf, text="취소", command=self._cancel).pack(side=tk.RIGHT)

        self.protocol("WM_DELETE_WINDOW", self._cancel)

        try:
            self._xl, self._wb = open_workbook_in_new_excel(path)
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("Excel", str(e), parent=self)
            self._invoke_cancel_once()
            self.destroy()
            return

        self.geometry("+%d+%d" % (parent.winfo_rootx() + 40, parent.winfo_rooty() + 40))

    def _invoke_cancel_once(self) -> None:
        if self._closed:
            return
        self._closed = True
        self._on_cancel()

    def _apply(self) -> None:
        if self._closed:
            return
        if self._xl is None:
            self._invoke_cancel_once()
            self.destroy()
            return
        try:
            parsed = read_active_sheet_full_used_as_selection(self._xl)
            if not parsed:
                messagebox.showwarning(
                    "확인",
                    "활성 시트를 읽을 수 없습니다.\n"
                    "복사할 시트를 선택해 UsedRange가 있게 한 뒤 다시 시도하세요.",
                    parent=self,
                )
                return
            try:
                meta, dest_sheet_name = self._persist_com(
                    self._xl,
                    self._wb,
                    self._path,
                    parsed,
                )
            except Exception as e:  # noqa: BLE001
                messagebox.showerror("검토용 통합문서", str(e), parent=self)
                close_excel_quietly(self._xl, self._wb)
                self._xl = None
                self._wb = None
                return
            close_excel_quietly(self._xl, self._wb)
            self._xl = None
            self._wb = None
            self._closed = True
            self._on_sheet_copied(meta, dest_sheet_name)
            self.destroy()
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("오류", str(e), parent=self)

    def _cancel(self) -> None:
        if self._closed:
            return
        close_excel_quietly(self._xl, self._wb)
        self._xl = None
        self._wb = None
        self._invoke_cancel_once()
        self.destroy()


class ExcelReviewRangeDialog(tk.Toplevel):
    """
    2단계: 결과 통합문서를 연 뒤 복사된 시트에서 검토 표 범위를 선택한다.
    """

    def __init__(
        self,
        parent: tk.Misc,
        snapshot_path: Path,
        dest_sheet_name: str,
        *,
        persist_com: ExcelPersistCom,
        on_ok: ExcelRangeOnOk,
        on_cancel: Callable[[], None],
    ) -> None:
        super().__init__(parent)
        self.title("Excel — 검토 범위 (2단계)")
        self.transient(parent)
        self.grab_set()
        self._persist_com = persist_com
        self._on_ok = on_ok
        self._on_cancel = on_cancel
        self._snapshot_path = snapshot_path
        self._dest_sheet = dest_sheet_name
        self._xl: Any = None
        self._wb: Any = None
        self._closed = False

        ttk.Label(
            self,
            text=snapshot_path.name,
            font=("Segoe UI", 10, "bold"),
        ).pack(padx=12, pady=(10, 4))
        ttk.Label(
            self,
            text=(
                f"복사 시트: «{dest_sheet_name}»\n\n"
                "1) 위 시트 탭을 클릭해 활성화되어 있는지 확인하세요.\n"
                "2) 검토·열 콤보에 쓸 표 영역을 드래그(또는 Shift+방향키)로 선택하세요.\n"
                "   맨 위 행은 열 이름(헤더)입니다.\n"
                "3) 「검토 범위 적용」을 누르면 Range_Set에 UsedRange·검토 범위 주소가 기록되고,\n"
                "   (BOM이면) 좌표명 열 정규화가 반영됩니다."
            ),
            justify=tk.LEFT,
        ).pack(padx=12, pady=4)

        bf = ttk.Frame(self, padding=12)
        bf.pack(fill=tk.X)
        ttk.Button(bf, text="검토 범위 적용", command=self._apply).pack(side=tk.RIGHT, padx=4)
        ttk.Button(bf, text="취소", command=self._cancel).pack(side=tk.RIGHT)

        self.protocol("WM_DELETE_WINDOW", self._cancel)

        try:
            self._xl, self._wb = open_workbook_in_new_excel(snapshot_path)
            try:
                self._wb.Worksheets(dest_sheet_name).Activate()
            except Exception:  # noqa: BLE001
                messagebox.showwarning(
                    "시트 선택",
                    f"«{dest_sheet_name}» 시트를 탭에서 직접 클릭해 활성화해 주세요.",
                    parent=self,
                )
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("Excel", str(e), parent=self)
            self._invoke_cancel_once()
            self.destroy()
            return

        self.geometry("+%d+%d" % (parent.winfo_rootx() + 48, parent.winfo_rooty() + 48))

    def _invoke_cancel_once(self) -> None:
        if self._closed:
            return
        self._closed = True
        self._on_cancel()

    def _apply(self) -> None:
        if self._closed:
            return
        if self._xl is None:
            self._invoke_cancel_once()
            self.destroy()
            return
        try:
            try:
                cur = str(self._xl.Selection.Worksheet.Name)
                if cur != self._dest_sheet:
                    messagebox.showwarning(
                        "확인",
                        f"검토 범위는 «{self._dest_sheet}» 시트 안에서만 지정할 수 있습니다.\n"
                        f"(현재 활성 시트: {cur})",
                        parent=self,
                    )
                    return
            except Exception:  # noqa: BLE001
                pass
            parsed = read_full_sheet_and_review_selection(self._xl)
            if not parsed:
                messagebox.showwarning(
                    "확인",
                    "범위를 읽을 수 없거나, 선택이 시트 사용 영역(UsedRange) 밖입니다.\n"
                    "복사 시트에서 UsedRange 안의 영역을 다시 선택하세요.",
                    parent=self,
                )
                return
            (
                full_h,
                full_d,
                rev_h,
                rev_d,
                _meta_discard,
                ur_r,
                ur_c,
                r1,
                c1,
                r2,
                c2,
            ) = parsed
            # 같은 통합문서를 Excel이 열고 있으면 openpyxl 저장 시 Windows에서 잠금(Permission denied)이 난다.
            close_excel_quietly(self._xl, self._wb)
            self._xl = None
            self._wb = None
            try:
                meta, _dest = self._persist_com(
                    None,
                    None,
                    self._snapshot_path,
                    parsed,
                )
            except Exception as e:  # noqa: BLE001
                messagebox.showerror("검토용 통합문서", str(e), parent=self)
                return
            self._closed = True
            self._on_ok(
                full_h,
                full_d,
                rev_h,
                rev_d,
                meta,
                ur_r,
                ur_c,
                r1,
                c1,
                r2,
                c2,
                self._dest_sheet,
            )
            self.destroy()
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("오류", str(e), parent=self)

    def _cancel(self) -> None:
        if self._closed:
            return
        close_excel_quietly(self._xl, self._wb)
        self._xl = None
        self._wb = None
        self._invoke_cancel_once()
        self.destroy()
