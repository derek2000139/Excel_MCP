from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Literal

Severity = Literal["low", "medium", "high", "critical"]

RULE_TYPE = Literal["CRITICAL", "HIGH", "MEDIUM", "LOW"]


@dataclass
class ScanFinding:
    rule_id: str
    severity: str
    category: str
    line_number: int
    message: str
    code_excerpt: str


@dataclass
class ScanResult:
    passed: bool
    blocked: bool
    risk_level: str
    scan_profile: str
    line_count: int
    procedure_names: list[str]
    findings: list[ScanFinding]
    notes: list[str]

    def to_dict(self) -> dict:
        return {
            "passed": self.passed,
            "blocked": self.blocked,
            "risk_level": self.risk_level,
            "scan_profile": self.scan_profile,
            "line_count": self.line_count,
            "procedure_names": self.procedure_names,
            "findings": [
                {
                    "rule_id": f.rule_id,
                    "severity": f.severity,
                    "category": f.category,
                    "line_number": f.line_number,
                    "message": f.message,
                    "code_excerpt": f.code_excerpt,
                }
                for f in self.findings
            ],
            "notes": self.notes,
        }


class VbaSecurityScanner:
    def __init__(
        self,
        block_levels: list[str] | None = None,
        warn_levels: list[str] | None = None,
        max_code_size: int = 1048576,
    ) -> None:
        if block_levels is None:
            block_levels = ["critical", "high"]
        if warn_levels is None:
            warn_levels = ["medium"]

        self._block_levels = set(block_levels)
        self._warn_levels = set(warn_levels)
        self._max_code_size = max_code_size

        self._rules: list[tuple[str, str, str, str]] = [
            ("VBA-C001", "CRITICAL", "System Command", r'\bShell\s*\('),
            ("VBA-C002", "CRITICAL", "System Object", r'CreateObject\s*\(\s*["\']WScript\.Shell["\']'),
            ("VBA-C003", "CRITICAL", "Key Simulation", r'\bSendKeys\b'),
            ("VBA-C004", "CRITICAL", "WinAPI Declare", r'\bDeclare\s+(?:Function|Sub)\s+\w+\s+Lib\b'),
            ("VBA-C005", "CRITICAL", "Excel4Macro", r'\bExecuteExcel4Macro\b'),
            ("VBA-H001", "HIGH", "File Write", r'\bOpen\s+\S+\s+For\s+(?:Output|Append|Binary)\b'),
            ("VBA-H002", "HIGH", "File Delete/Copy/Rename", r'\b(?:Kill|FileCopy|Name\s+\S+\s+As)\b'),
            ("VBA-H003", "HIGH", "Directory", r'\b(?:MkDir|RmDir)\b'),
            ("VBA-H004", "HIGH", "Network Request", r'\b(?:XMLHTTP|WinHttp|ServerXMLHTTP)\b'),
            ("VBA-H005", "HIGH", "Database", r'\bADODB\.Connection\b'),
            ("VBA-H006", "HIGH", "FileSystemObject", r'CreateObject\s*\(\s*["\']Scripting\.FileSystemObject["\']'),
            ("VBA-M001", "MEDIUM", "Delete Operation", r'\.Delete\b'),
            ("VBA-M002", "MEDIUM", "Suppress Alerts", r'Application\.DisplayAlerts\s*=\s*False'),
            ("VBA-M003", "MEDIUM", "Disable Events", r'Application\.EnableEvents\s*=\s*False'),
            ("VBA-M004", "MEDIUM", "Ignore Errors", r'\bOn\s+Error\s+Resume\s+Next\b'),
            ("VBA-M005", "MEDIUM", "Manual Calculation", r'Calculation\s*=\s*xlCalculationManual'),
        ]

        self._procedure_pattern = re.compile(
            r'^(?:Public\s+|Private\s+)?(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)',
            re.MULTILINE | re.IGNORECASE,
        )

    def scan(self, code: str, module_type: str = "standard_module") -> ScanResult:
        findings: list[ScanFinding] = []
        notes: list[str] = []

        if len(code) > self._max_code_size:
            return ScanResult(
                passed=False,
                blocked=True,
                risk_level="critical",
                scan_profile="default",
                line_count=len(code.splitlines()),
                procedure_names=[],
                findings=[],
                notes=[f"Code size {len(code)} exceeds maximum {self._max_code_size}"],
            )

        code_segments, procedure_names = self._lexical_split(code)

        for segment in code_segments:
            if segment.kind == "code":
                for rule_id, severity, category, pattern in self._rules:
                    matches = list(re.finditer(pattern, segment.content, re.IGNORECASE))
                    for match in matches:
                        m_start = match.start()
                        m_end = match.end()
                        line_num = self._calc_line_for_pos(code, segment.start + m_start)
                        excerpt = self._extract_excerpt(segment.content, m_start, m_end)
                        findings.append(
                            ScanFinding(
                                rule_id=rule_id,
                                severity=severity.lower(),
                                category=category,
                                line_number=line_num,
                                message=f"Potential {category.lower()} detected: pattern '{match.group()}'",
                                code_excerpt=excerpt,
                            )
                        )

        max_severity = self._compute_max_severity(findings)
        blocked = max_severity in self._block_levels

        if blocked:
            notes.append(f"Code blocked due to {max_severity} severity findings")
        elif max_severity == "medium":
            notes.append("Code allowed with warnings")

        return ScanResult(
            passed=not blocked,
            blocked=blocked,
            risk_level=max_severity,
            scan_profile="default",
            line_count=len(code.splitlines()),
            procedure_names=procedure_names,
            findings=findings,
            notes=notes,
        )

    def _lexical_split(self, code: str) -> tuple[list[CodeSegment], list[str]]:
        segments: list[CodeSegment] = []

        proc_matches = self._procedure_pattern.findall(code)
        procedure_names = list(dict.fromkeys(proc_matches))

        lines = code.splitlines(True)
        start_offset = 0

        for line in lines:
            i = 0
            n = len(line)
            in_string = False

            lstripped = line.lstrip()
            if lstripped.startswith("Rem ") or lstripped.startswith("Rem\t") or lstripped == "Rem\n" or lstripped == "Rem":
                start_offset += n
                continue

            code_start = 0
            current_code = ""

            while i < n:
                c = line[i]

                if not in_string:
                    if c == "'":
                        if current_code:
                            segments.append(CodeSegment("code", current_code, start_offset + code_start))
                        current_code = ""
                        break

                    elif c == '"':
                        if current_code:
                            segments.append(CodeSegment("code", current_code, start_offset + code_start))
                        current_code = ""
                        in_string = True
                        i += 1
                        continue

                    else:
                        if not current_code:
                            code_start = i
                        current_code += c
                        i += 1
                else:
                    if c == '"':
                        if i + 1 < n and line[i + 1] == '"':
                            i += 2
                            continue
                        else:
                            in_string = False
                            i += 1
                            code_start = i
                            continue
                    else:
                        i += 1

            if current_code:
                segments.append(CodeSegment("code", current_code, start_offset + code_start))

            start_offset += n

        return segments, procedure_names

    def _calc_line_for_pos(self, code: str, pos: int) -> int:
        return code[:pos].count("\n") + 1

    def _extract_excerpt(self, content: str, start: int, end: int) -> str:
        line_start = content.rfind("\n", 0, start) + 1
        line_end = content.find("\n", end)
        if line_end == -1:
            line_end = len(content)
        return content[line_start:line_end].strip()

    def _compute_max_severity(self, findings: list[ScanFinding]) -> str:
        if not findings:
            return "low"
        severity_order = ["low", "medium", "high", "critical"]
        max_idx = -1
        for f in findings:
            if f.severity in severity_order:
                idx = severity_order.index(f.severity)
                if idx > max_idx:
                    max_idx = idx
        return severity_order[max_idx] if max_idx >= 0 else "low"


@dataclass
class CodeSegment:
    kind: str
    content: str
    start: int
