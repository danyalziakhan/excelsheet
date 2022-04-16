# MIT License

# Copyright (c) 2022 Danyal Zia Khan

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

from functools import cache
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from typing import Any

# ? https://stackoverflow.com/questions/39952931/win32-python-to-get-all-column-names-from-excel
LETTERS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"


@cache
def col_row_to_excel(row: int, col: int):
    """Convert given row and column number to an Excel-style cell name."""
    result: list[Any] = []
    while col:
        col, rem = divmod(col - 1, 26)
        result[:0] = LETTERS[rem]
    return "".join(result) + str(row)


@cache
def col_to_excel(col: int):  # col is 1 based
    """Convert given column number to an Excel-style cell name."""
    excel_col = ""
    div = col
    while div:
        (div, mod) = divmod(div - 1, 26)  # will return (x, 0 .. 25)
        excel_col = chr(mod + 65) + excel_col

    return excel_col


def get_cell_range(alphabet: str, start: int, end: int, end_offset: int = 0):
    return f"{alphabet}{start}:{alphabet}{end + end_offset}"
