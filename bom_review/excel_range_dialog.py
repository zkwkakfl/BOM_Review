"""Excel에서 시트·범위 선택 후 적용하는 모달 창."""

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
    read_full_sheet_and_review_selection,
)
from bom_review.excel_snapshot import ExcelParsedSelection

ExcelPersistCom = Callable[
    [Any, Any, Path, ExcelParsedSelection],
    tuple[SelectionSourceMeta, str],
]

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


class ExcelRangeDialog(tk.Toplevel):
    """
    Excel을 연 뒤 시트를 활성화하고, 검토할 셀 범위를 선택한 다음 「이 범위 적용」으로 확정한다.
    시트 전체는 COM 복사로 결과 통합문서에 반영(서식 유지)되고, 콤보·검토는 선택 범위 기준이다.
    """

    def __init__(
        self,
        parent: tk.Misc,
        path: Path,
        *,
        persist_com: ExcelPersistCom,
        on_ok: ExcelRangeOnOk,
        on_cancel: Callable[[], None],
    ) -> None:
        super().__init__(parent)
        self.title("Excel 시트·범위 선택")
        self.transient(parent)
        self.grab_set()
        self._persist_com = persist_com
        self._on_ok = on_ok
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
                "1) 복사할 시트를 클릭해 활성화하세요.\n"
                "2) 검토에 쓸 표 영역을 드래그(또는 Shift+방향키)로 선택하세요. 첫 행은 열 이름(헤더)입니다.\n"
                "3) 「이 범위 적용」을 누르면 이 시트 전체가 Excel 복사로 결과 통합문서에 붙여 넣어지고(서식 유지),\n"
                "   원본 통합문서는 닫힙니다. Range_Set에는 결과 파일 안의 UsedRange·검토 범위 주소가 기록됩니다."
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
            parsed = read_full_sheet_and_review_selection(self._xl)
            if not parsed:
                messagebox.showwarning(
                    "확인",
                    "시트를 읽을 수 없거나, 선택 범위가 시트 사용 영역(UsedRange) 밖입니다.\n"
                    "시트를 선택한 뒤 검토할 범위를 UsedRange 안에서 다시 지정하세요.",
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
            self._xl = None
            self._wb = None
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
                dest_sheet_name,
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
