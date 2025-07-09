# MIT License

# Copyright (c) 2022-2025 Danyal Zia Khan

# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:

# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.

# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.

from __future__ import annotations

import platform
from contextlib import contextmanager
from typing import TYPE_CHECKING, cast

import pandas as pd

from excelsheet.utils import get_cell_range

if platform.system() == "Windows":
    import pythoncom
    import win32com
    import win32com.client
    from win32com.client.dynamic import CDispatch


if TYPE_CHECKING:
    from collections.abc import Sequence
    from typing import Any

    from openpyxl.worksheet.worksheet import Worksheet


@contextmanager
def open_ms_excel(filename: str, /, sheet: str = "Sheet1"):
    """
    Open and read the worksheet in Microsoft Excel format.
    """

    # ? Without CoInitialize, ThreadPoolExecutor is not working :/
    pythoncom.CoInitialize()  # type: ignore

    # ? Force dynamic client
    # ? https://mail.python.org/pipermail/python-win32/2012-July/012447.html
    excel: CDispatch = win32com.client.dynamic.Dispatch("Excel.Application")

    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        # Open up the file
        wb: CDispatch = cast(CDispatch, excel.Workbooks.Open(filename))  # type: ignore
    except pythoncom.com_error as err:  # type: ignore
        raise FileNotFoundError(f"{filename} not found") from err

    try:
        assert wb
    except AssertionError as err:
        from colorama import Fore

        raise AssertionError(
            "".join(
                [
                    Fore.RED,
                    "The Excel template cannot be opened. Make sure it is present in the current directory",
                    Fore.RESET,
                ]
            )
        ) from err
    else:
        ws = cast(CDispatch, wb.Worksheets(sheet))  # type: ignore

    try:
        assert ws
    except AssertionError as err:
        from colorama import Fore

        raise AssertionError(
            "".join(
                [
                    Fore.RED,
                    f"The worksheet {sheet} not found inside the Excel. Are you sure it is the correct name?",
                    Fore.RESET,
                ]
            )
        ) from err
    else:
        yield ws
    finally:
        wb.Close()  # type: ignore
        excel.Quit()  # type: ignore


def write_to_cell_win32(worksheet: Any, values: Sequence[str], range_fmt: str):
    worksheet.Range(range_fmt).Value = [[i] for i in values]


def write_to_excel_template_cell_win32(
    worksheet: Any, df: pd.DataFrame, name: str, alphabet: str
):
    values: list[Any] = df[name].to_list()
    write_to_cell_win32(
        worksheet=worksheet,
        values=values,
        range_fmt=get_cell_range(
            alphabet=alphabet, start=2, end=len(values), end_offset=1
        ),
    )


def write_to_cell_openpyxl(worksheet: Worksheet, values: Sequence[str], range_fmt: str):
    for cell, value in zip(worksheet[range_fmt], values):  # type: ignore
        cell[0].value = value


def write_to_excel_template_cell_openpyxl(
    worksheet: Any, df: pd.DataFrame, name: str, alphabet: str
):
    values: list[Any] = df[name].to_list()
    write_to_cell_openpyxl(
        worksheet=worksheet,
        values=values,
        range_fmt=get_cell_range(
            alphabet=alphabet, start=2, end=len(values), end_offset=1
        ),
    )
