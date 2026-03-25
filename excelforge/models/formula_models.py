from __future__ import annotations

from pydantic import Field

from .common import ClientRequestMixin, StrictModel
from .range_models import A1_CELL_PATTERN, A1_RANGE_PATTERN, ScalarValue


class FormulaValidateExpressionRequest(ClientRequestMixin):
    workbook_id: str
    sheet_name: str = Field(max_length=31)
    anchor_cell: str = Field(pattern=A1_CELL_PATTERN)
    formula: str = Field(max_length=8192)


class FormulaValidateExpressionData(StrictModel):
    syntax_valid: bool
    starts_with_equals: bool
    length_valid: bool
    english_formula_style_expected: bool
    reference_candidates: list[str]
    compatibility_warnings: list[str]
    guarantee_level: str
    notes: list[str]


class FormulaFillRangeRequest(ClientRequestMixin):
    workbook_id: str
    sheet_name: str = Field(max_length=31)
    range: str = Field(pattern=A1_RANGE_PATTERN)
    formula: str = Field(max_length=8192)
    formula_type: str = Field(default="standard")
    preview_rows: int = Field(default=5, ge=1, le=5)


class FormulaFillPreviewItem(StrictModel):
    cell: str
    formula: str
    value: ScalarValue


class FormulaFillRangeData(StrictModel):
    sheet_name: str
    affected_range: str
    cells_written: int
    formula_type: str
    anchor_formula: str
    preview: list[FormulaFillPreviewItem]
    snapshot_id: str
    spill_range: str | None = None
    calculation_completed: bool = True
    has_spill_error: bool = False


class FormulaSetSingleRequest(ClientRequestMixin):
    workbook_id: str
    sheet_name: str = Field(max_length=31)
    cell: str = Field(pattern=A1_CELL_PATTERN)
    formula: str = Field(max_length=8192)
    formula_type: str = Field(default="standard")


class FormulaSetSingleData(StrictModel):
    sheet_name: str
    cell: str
    formula: str
    formula_type: str
    calculated_value: ScalarValue | None
    has_error: bool
    error_type: str | None
    spill_range: str | None = None
    spill_preview: list[dict] = []
    calculation_completed: bool = True
    snapshot_id: str


class FormulaGetDependenciesRequest(ClientRequestMixin):
    workbook_id: str
    sheet_name: str = Field(max_length=31)
    cell: str = Field(pattern=A1_CELL_PATTERN)


class FormulaDependency(StrictModel):
    sheet: str
    range: str


class FormulaGetDependenciesData(StrictModel):
    sheet_name: str
    cell: str
    has_formula: bool
    formula: str | None
    calculated_value: ScalarValue | None
    precedents: list[FormulaDependency] = []
    dependents: list[FormulaDependency] = []


class FormulaRepairReferencesRequest(ClientRequestMixin):
    workbook_id: str
    sheet_name: str = Field(max_length=31)
    range: str = Field(pattern=A1_RANGE_PATTERN)
    action: str = Field(pattern="^(scan|repair)$")
    replacements: list[dict] | None = None


class ScanFinding(StrictModel):
    cell: str
    formula: str
    error_type: str | None
    has_external_ref: bool
    referenced_sheets: list[str]


class ScanSummary(StrictModel):
    ref_errors: int
    name_errors: int
    other_errors: int
    external_refs: int
    unique_referenced_sheets: list[str]


class FormulaRepairScanData(StrictModel):
    action: str = "scan"
    sheet_name: str
    scanned_range: str
    total_cells: int
    formula_cells: int
    error_cells: int
    findings: list[ScanFinding]
    summary: ScanSummary


class RepairModification(StrictModel):
    cell: str
    old_formula: str
    new_formula: str
    new_value: ScalarValue | None
    still_has_error: bool


class FormulaRepairData(StrictModel):
    action: str = "repair"
    sheet_name: str
    repaired_range: str
    replacements_applied: int
    cells_modified: int
    cells_unchanged: int
    modifications: list[RepairModification]
    snapshot_id: str
