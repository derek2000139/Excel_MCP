from __future__ import annotations

from typing import Literal

from .common import ClientRequestMixin, StrictModel

ProcedureKind = Literal["Sub", "Function", "Property Get", "Property Let", "Property Set"]
ModuleType = Literal["standard_module", "class_module", "userform", "document"]


class VbaInspectProjectRequest(ClientRequestMixin):
    workbook_id: str


class VbaProcedure(StrictModel):
    name: str
    kind: ProcedureKind
    start_line: int


class VbaProcedureWithEnd(StrictModel):
    name: str
    kind: ProcedureKind
    start_line: int
    end_line: int


class VbaModuleSummary(StrictModel):
    name: str
    type: ModuleType
    line_count: int
    procedure_count: int
    procedures: list[VbaProcedure]


class VbaReference(StrictModel):
    name: str
    description: str
    major: int
    minor: int
    is_broken: bool


class VbaInspectProjectData(StrictModel):
    workbook_id: str
    project_name: str
    protection: Literal["none", "locked"]
    modules: list[VbaModuleSummary]
    references: list[VbaReference]
    total_modules: int
    total_code_lines: int


class VbaGetModuleCodeRequest(ClientRequestMixin):
    workbook_id: str
    module_name: str


class VbaGetModuleCodeData(StrictModel):
    workbook_id: str
    module_name: str
    module_type: ModuleType
    code: str
    line_count: int
    truncated: bool
    procedures: list[VbaProcedureWithEnd]


class VbaScanCodeRequest(ClientRequestMixin):
    code: str
    module_name: str | None = None
    module_type: ModuleType = "standard_module"


class VbaScanFinding(StrictModel):
    rule_id: str
    severity: str
    category: str
    line_number: int
    message: str
    code_excerpt: str


class VbaScanCodeData(StrictModel):
    passed: bool
    blocked: bool
    risk_level: str
    scan_profile: str
    line_count: int
    procedure_names: list[str]
    findings: list[VbaScanFinding]
    notes: list[str]


class VbaSyncModuleRequest(ClientRequestMixin):
    workbook_id: str
    module_name: str
    module_type: ModuleType = "standard_module"
    code: str
    overwrite: bool = False


class VbaScanResultSummary(StrictModel):
    passed: bool
    blocked: bool
    risk_level: str
    findings_count: int


class VbaSyncModuleData(StrictModel):
    workbook_id: str
    module_name: str
    module_type: ModuleType
    action: str
    line_count: int
    procedure_names: list[str]
    scan_result: VbaScanResultSummary
    backup_id: str


class VbaRemoveModuleRequest(ClientRequestMixin):
    workbook_id: str
    module_name: str


class VbaRemoveModuleData(StrictModel):
    workbook_id: str
    removed_module: str
    module_type: ModuleType
    backup_id: str


class VbaExecuteMacroRequest(ClientRequestMixin):
    workbook_id: str
    procedure_name: str
    arguments: list = []
    timeout_seconds: int = 30


class VbaExecuteMacroData(StrictModel):
    workbook_id: str
    procedure_name: str
    executed: bool
    return_value: str | int | float | bool | None
    execution_time_ms: int
    scan_result: VbaScanResultSummary
    backup_id: str


class VbaExecuteInlineRequest(ClientRequestMixin):
    workbook_id: str
    code: str
    procedure_name: str = "Main"
    timeout_seconds: int = 30


class VbaExecuteInlineData(StrictModel):
    workbook_id: str
    procedure_name: str
    executed: bool
    return_value: str | int | float | bool | None
    execution_time_ms: int
    temp_module_cleaned: bool
    scan_result: VbaScanResultSummary
    backup_id: str


class VbaExportModuleRequest(ClientRequestMixin):
    workbook_id: str
    module_name: str
    file_path: str
    overwrite: bool = False


class VbaExportModuleData(StrictModel):
    workbook_id: str
    module_name: str
    module_type: str
    file_path: str
    file_size_bytes: int
    line_count: int


class VbaImportModuleRequest(ClientRequestMixin):
    workbook_id: str
    file_path: str
    module_name: str | None = None
    overwrite: bool = False


class VbaImportModuleData(StrictModel):
    workbook_id: str
    module_name: str
    module_type: str
    action: str
    line_count: int
    procedure_names: list[str]
    scan_result: VbaScanResultSummary
    backup_id: str


class VbaCompileRequest(ClientRequestMixin):
    workbook_id: str


class CompileError(StrictModel):
    module_name: str
    line_number: int | None
    error_message: str
    code_excerpt: str | None


class VbaCompileData(StrictModel):
    workbook_id: str
    project_name: str
    compile_success: bool
    method: str
    errors: list[CompileError]
    warnings: list[str]
    modules_checked: int
    total_lines: int
