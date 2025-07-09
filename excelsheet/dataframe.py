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

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from typing import Any

    import pandas as pd


def remove_existing_rows(
    df: pd.DataFrame | pd.Series[Any],
) -> pd.DataFrame | pd.Series[Any]:
    """
    Delete all the rows so that only columns are present
    """
    return df.drop(df.index.to_list(), axis="index").reset_index(drop=True)  # type: ignore


def put_last_column_to_first(df: pd.DataFrame) -> pd.DataFrame:
    """
    Reorder the columns of a DataFrame
    """
    return df[[df.columns[-1]] + [col for col in df.columns if col != df.columns[-1]]]  # type: ignore
