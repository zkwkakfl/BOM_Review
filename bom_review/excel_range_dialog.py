"""Excel에서 드래그·Shift+방향키로 범위 선택 후 적용하는 모달 창."""

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
    read_selection_as_header_and_rows,
    read_selection_source_meta,
)


class ExcelRangeDialog(tk.Toplevel):
    """Excel을 연 뒤 사용자가 범위를 고르고 [이 범위 적용]으로 확정한다."""

    def __init__(
        self,
        parent: tk.Misc,
        path: Path,
        *,
        on_ok: Callable[[list[str], list[list[Any]], SelectionSourceMeta], None],
        on_cancel: Callable[[], None],
    ) -> None:
        super().__init__(parent)
        self.title("Excel 범위 선택")
        self.transient(parent)
        self.grab_set()
        self._on_ok = on_ok
        self._on_cancel = on_cancel
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
                "Excel 창이 열리면, 마우스 드래그나 Shift+방향키로\n"
                "데이터 범위를 선택하세요. 첫 행은 열 이름(헤더)로 사용됩니다."
            ),
            justify=tk.LEFT,
        ).pack(padx=12, pady=4)

        bf = ttk.Frame(self, padding=12)
        bf.pack(fill=tk.X)
        ttk.Button(bf, text="이 범위 적용", command=self._apply).pack(side=tk.RIGHT, padx=4)
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
            meta = read_selection_source_meta(self._xl)
            parsed = read_selection_as_header_and_rows(self._xl)
            if not parsed:
                messagebox.showwarning(
                    "확인",
                    "범위가 비어 있거나 읽을 수 없습니다. 셀 범위를 다시 선택하세요.",
                    parent=self,
                )
                return
            headers, data = parsed
            close_excel_quietly(self._xl, self._wb)
            self._xl = None
            self._wb = None
            self._closed = True
            self._on_ok(headers, data, meta)
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
