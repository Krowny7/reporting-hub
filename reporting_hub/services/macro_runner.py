from __future__ import annotations

from dataclasses import dataclass
from typing import Callable, List, Optional

from ..excel.controller import ExcelController


Logger = Callable[[str], None]


@dataclass
class RunRequest:
    workbook_path: str
    macro_name: str
    args: List[str]
    excel_mode: str = "minimized"


class MacroRunner:
    """Orchestrates the full 'open workbook -> run macro' flow."""

    def __init__(self, logger: Logger):
        self.log = logger
        self.controller = ExcelController(logger)

    def run(self, req: RunRequest, quit_excel_when_done: bool = False) -> None:
        if self.controller.excel is None:
            self.controller.launch_new_instance()

        self.controller.set_excel_mode(req.excel_mode)
        wb_name = self.controller.open_or_activate_by_path(req.workbook_path)
        self.controller.run_macro(wb_name, req.macro_name, *req.args)

        if quit_excel_when_done:
            self.controller.quit_excel()
