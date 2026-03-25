from __future__ import annotations

from typing import Literal

from pydantic import Field

from .common import ClientRequestMixin, StrictModel

A1_RANGE_PATTERN = r"^\$?[A-Za-z]{1,3}\$?\d{1,7}(?::\$?[A-Za-z]{1,3}\$?\d{1,7})?$"
HEX_COLOR_PATTERN = r"^#[0-9A-Fa-f]{6}$"


class FontStyle(StrictModel):
    bold: bool | None = None
    italic: bool | None = None
    size: float | None = Field(default=None, ge=1, le=409)
    name: str | None = None
    color: str | None = Field(default=None, pattern=HEX_COLOR_PATTERN)


class FillStyle(StrictModel):
    color: str | None = Field(default=None, pattern=HEX_COLOR_PATTERN)


class AlignmentStyle(StrictModel):
    horizontal: Literal["left", "center", "right", "general"] | None = None
    vertical: Literal["top", "center", "bottom"] | None = None
    wrap_text: bool | None = None


class StyleModel(StrictModel):
    font_name: str | None = None
    font_size: float | None = Field(default=None, ge=1, le=409)
    font_bold: bool | None = None
    font_italic: bool | None = None
    font_color: str | None = Field(default=None, pattern=HEX_COLOR_PATTERN)
    fill_color: str | None = Field(default=None, pattern=HEX_COLOR_PATTERN)
    number_format: str | None = None
    horizontal_alignment: Literal["left", "center", "right", "general"] | None = None
    vertical_alignment: Literal["top", "center", "bottom"] | None = None
    wrap_text: bool | None = None
    border_style: Literal["thin", "medium", "thick", "none", "all_thin"] | None = None


class FormatSetRangeStyleRequest(ClientRequestMixin):
    workbook_id: str
    sheet_name: str = Field(max_length=31)
    range: str = Field(pattern=A1_RANGE_PATTERN)
    style: StyleModel | None = None


class FormatSetRangeStyleData(StrictModel):
    sheet_name: str
    affected_range: str
    cells_affected: int
    properties_set: list[str]


class FormatAutoFitColumnsRequest(ClientRequestMixin):
    workbook_id: str
    sheet_name: str = Field(max_length=31)
    columns: str | None = None


class FormatAutoFitColumnsData(StrictModel):
    sheet_name: str
    columns_adjusted: list[str]
    column_count: int


class FormatManageRequest(ClientRequestMixin):
    action: str = Field(pattern="^(set_style|auto_fit_columns)$")
    workbook_id: str
    sheet_name: str = Field(max_length=31)
    range: str = ""
    style: StyleModel | None = None
    columns: str | None = None
