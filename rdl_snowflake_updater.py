from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import sys
import tempfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


# -----------------------------------------------------------------------------
# Patterns
# -----------------------------------------------------------------------------
DATASET_BLOCK_RE = re.compile(
    r"<(?:(?:[A-Za-z_][\w.\-]*):)?DataSet\b(?P<attrs>[^>]*)>"
    r"(?P<body>.*?)"
    r"</(?:(?:[A-Za-z_][\w.\-]*):)?DataSet\s*>",
    re.IGNORECASE | re.DOTALL,
)

DATASET_NAME_RE = re.compile(
    r"\bName\s*=\s*(?P<quote>['\"])(?P<name>.*?)(?P=quote)",
    re.IGNORECASE | re.DOTALL,
)

COMMAND_TEXT_RE = re.compile(
    r"(?P<open><(?:(?:[A-Za-z_][\w.\-]*):)?CommandText\b[^>]*>)"
    r"(?P<cmd>.*?)"
    r"(?P<close></(?:(?:[A-Za-z_][\w.\-]*):)?CommandText\s*>)",
    re.IGNORECASE | re.DOTALL,
)

# Power BI RDL CommandText commonly contains a JSON payload. Therefore M-code
# string quotes inside MashupDocument appear as \"value\". The same matcher also
# supports plain quotes and XML-escaped quotes for other RDL variants.
SNOWFLAKE_SOURCE_RE = re.compile(
    r"(?P<head>\bSource\s*=\s*Snowflake\.Databases\s*\(\s*)"
    r"(?P<q1>\\\"|&quot;|\")(?P<server>.*?)(?P=q1)"
    r"(?P<separator>\s*,\s*)"
    r"(?P<q2>\\\"|&quot;|\")(?P<warehouse>.*?)(?P=q2)",
    re.IGNORECASE | re.DOTALL,
)

# Matches the ConnectionOverrides JSON metadata shown by Power BI, e.g.:
# "Path":"myprod.snowflakecomputing.com;MY_WAREHOUSE"
# It supports both literal JSON quotes and XML-escaped &quot; forms.
CONNECTION_OVERRIDE_PATH_RE = re.compile(
    r"(?P<head>(?:\"|&quot;)Path(?:\"|&quot;)\s*:\s*(?P<q>\"|&quot;))"
    r"(?P<server>.*?);(?P<warehouse>.*?)"
    r"(?P=q)",
    re.IGNORECASE | re.DOTALL,
)


AUDIT_HEADERS = [
    "Timestamp",
    "File Name",
    "Dataset Name",
    "Modification Type",
    "Status",
    "Before Server",
    "After Server",
    "Before Warehouse",
    "After Warehouse",
    "Before Value",
    "After Value",
    "Before Code Snapshot",
    "After Code Snapshot",
    "Input Path",
    "Output Path",
    "Original SHA256",
    "Modified SHA256",
]


@dataclass
class AuditRow:
    timestamp: str
    file_name: str
    dataset_name: str
    modification_type: str
    status: str
    before_server: str = ""
    after_server: str = ""
    before_warehouse: str = ""
    after_warehouse: str = ""
    before_value: str = ""
    after_value: str = ""
    before_code_snapshot: str = ""
    after_code_snapshot: str = ""
    input_path: str = ""
    output_path: str = ""
    original_sha256: str = ""
    modified_sha256: str = ""

    def as_list(self) -> list[str]:
        return [
            self.timestamp,
            self.file_name,
            self.dataset_name,
            self.modification_type,
            self.status,
            self.before_server,
            self.after_server,
            self.before_warehouse,
            self.after_warehouse,
            self.before_value,
            self.after_value,
            self.before_code_snapshot,
            self.after_code_snapshot,
            self.input_path,
            self.output_path,
            self.original_sha256,
            self.modified_sha256,
        ]


@dataclass
class SourceChange:
    before_server: str
    after_server: str
    before_warehouse: str
    after_warehouse: str
    before_code: str
    after_code: str

    @property
    def changed(self) -> bool:
        return (
            self.before_server != self.after_server
            or self.before_warehouse != self.after_warehouse
        )


@dataclass
class PathChange:
    before_server: str
    after_server: str
    before_warehouse: str
    after_warehouse: str
    before_code: str
    after_code: str

    @property
    def changed(self) -> bool:
        return (
            self.before_server != self.after_server
            or self.before_warehouse != self.after_warehouse
        )


@dataclass
class ModifyResult:
    text: str
    audit_rows: list[AuditRow]
    source_match_count: int
    source_change_count: int
    path_match_count: int
    path_change_count: int


@dataclass
class FileProcessResult:
    input_file: Path
    output_file: Path | None
    dataset_names: list[str]
    source_match_count: int
    source_change_count: int
    path_change_count: int
    deleted_original: bool


# -----------------------------------------------------------------------------
# Encoding / XML helpers
# -----------------------------------------------------------------------------
def decode_xml_bytes(data: bytes) -> tuple[str, str, bytes]:
    """Decode XML while preserving the original BOM/encoding for write-back."""
    if data.startswith(b"\xef\xbb\xbf"):
        return data[3:].decode("utf-8"), "utf-8", b"\xef\xbb\xbf"
    if data.startswith(b"\xff\xfe"):
        return data[2:].decode("utf-16-le"), "utf-16-le", b"\xff\xfe"
    if data.startswith(b"\xfe\xff"):
        return data[2:].decode("utf-16-be"), "utf-16-be", b"\xfe\xff"

    declaration = data[:512]
    match = re.search(br"encoding\s*=\s*['\"]([^'\"]+)['\"]", declaration, re.IGNORECASE)
    encoding = match.group(1).decode("ascii") if match else "utf-8"
    return data.decode(encoding), encoding, b""


def encode_xml_text(text: str, encoding: str, bom: bytes) -> bytes:
    return bom + text.encode(encoding)


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag.rsplit(":", 1)[-1]


def validate_and_list_datasets(xml_text: str) -> list[str]:
    """Validate XML and return DataSet/@Name values located under DataSets."""
    root = ET.fromstring(xml_text)
    names: list[str] = []

    for data_sets in root.iter():
        if local_name(data_sets.tag).lower() != "datasets":
            continue
        for child in list(data_sets):
            if local_name(child.tag).lower() != "dataset":
                continue
            name = child.attrib.get("Name") or child.attrib.get("name")
            if name:
                names.append(name)

    return names


def xml_escape_text_value(value: str) -> str:
    """Escape characters that are unsafe inside XML element text."""
    return value.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def validate_target_value(value: str, label: str) -> None:
    """Reject characters that would break the serialized JSON/M string structure."""
    if any(ch in value for ch in ('"', "\r", "\n")):
        raise ValueError(f"{label} cannot contain a double quote or newline.")


# -----------------------------------------------------------------------------
# RDL modification
# -----------------------------------------------------------------------------
def replace_snowflake_sources(
    command_text: str,
    new_server: str,
    new_warehouse: str,
) -> tuple[str, list[SourceChange]]:
    """Replace the first two positional parameters of Snowflake.Databases(...)."""
    changes: list[SourceChange] = []

    def replacement(match: re.Match[str]) -> str:
        before_server = match.group("server")
        before_warehouse = match.group("warehouse")

        # CommandText is still raw XML text, so target values must remain XML-safe.
        server_for_xml = xml_escape_text_value(new_server)
        warehouse_for_xml = xml_escape_text_value(new_warehouse)

        after = (
            f'{match.group("head")}'
            f'{match.group("q1")}{server_for_xml}{match.group("q1")}'
            f'{match.group("separator")}'
            f'{match.group("q2")}{warehouse_for_xml}{match.group("q2")}'
        )

        changes.append(
            SourceChange(
                before_server=before_server,
                after_server=new_server,
                before_warehouse=before_warehouse,
                after_warehouse=new_warehouse,
                before_code=match.group(0),
                after_code=after,
            )
        )
        return after

    return SNOWFLAKE_SOURCE_RE.sub(replacement, command_text), changes


def replace_connection_override_paths(
    command_text: str,
    source_changes: list[SourceChange],
    new_server: str,
    new_warehouse: str,
) -> tuple[str, list[PathChange]]:
    """
    Keep ConnectionOverrides[].Path synchronized with the Snowflake source.

    Only paths whose server/warehouse pair matches a Snowflake.Databases pair
    found in the same CommandText are updated. This avoids changing unrelated
    connection metadata.
    """
    source_pairs = {
        (change.before_server, change.before_warehouse)
        for change in source_changes
    }
    changes: list[PathChange] = []

    if not source_pairs:
        return command_text, changes

    def replacement(match: re.Match[str]) -> str:
        before_server = match.group("server")
        before_warehouse = match.group("warehouse")

        if (before_server, before_warehouse) not in source_pairs:
            return match.group(0)

        server_for_xml = xml_escape_text_value(new_server)
        warehouse_for_xml = xml_escape_text_value(new_warehouse)
        after = (
            f'{match.group("head")}'
            f'{server_for_xml};{warehouse_for_xml}'
            f'{match.group("q")}'
        )

        changes.append(
            PathChange(
                before_server=before_server,
                after_server=new_server,
                before_warehouse=before_warehouse,
                after_warehouse=new_warehouse,
                before_code=match.group(0),
                after_code=after,
            )
        )
        return after

    return CONNECTION_OVERRIDE_PATH_RE.sub(replacement, command_text), changes


def modify_rdl_text(
    xml_text: str,
    new_server: str,
    new_warehouse: str,
    input_file: Path,
    output_file: Path,
) -> ModifyResult:
    """Modify Snowflake connection values dataset-by-dataset without XML reserialization."""
    audit_rows: list[AuditRow] = []
    source_match_count = 0
    source_change_count = 0
    path_match_count = 0
    path_change_count = 0
    now = datetime.now().astimezone().isoformat(timespec="seconds")

    parts: list[str] = []
    cursor = 0

    for dataset_match in DATASET_BLOCK_RE.finditer(xml_text):
        parts.append(xml_text[cursor:dataset_match.start()])
        dataset_block = dataset_match.group(0)
        attrs = dataset_match.group("attrs")
        name_match = DATASET_NAME_RE.search(attrs)
        dataset_name = name_match.group("name") if name_match else "<unnamed>"

        dataset_source_changes: list[SourceChange] = []
        dataset_path_changes: list[PathChange] = []

        def command_replacement(command_match: re.Match[str]) -> str:
            command_text = command_match.group("cmd")

            modified_cmd, source_changes = replace_snowflake_sources(
                command_text, new_server, new_warehouse
            )
            final_cmd, path_changes = replace_connection_override_paths(
                modified_cmd,
                source_changes,
                new_server,
                new_warehouse,
            )

            dataset_source_changes.extend(source_changes)
            dataset_path_changes.extend(path_changes)
            return f'{command_match.group("open")}{final_cmd}{command_match.group("close")}'

        modified_block = COMMAND_TEXT_RE.sub(command_replacement, dataset_block)
        parts.append(modified_block)
        cursor = dataset_match.end()

        if dataset_source_changes:
            source_match_count += len(dataset_source_changes)
            source_change_count += sum(change.changed for change in dataset_source_changes)
            path_match_count += len(dataset_path_changes)
            path_change_count += sum(change.changed for change in dataset_path_changes)

            for index, change in enumerate(dataset_source_changes, start=1):
                status = "MODIFIED" if change.changed else "NO_CHANGE_REQUIRED"
                audit_rows.append(
                    AuditRow(
                        timestamp=now,
                        file_name=input_file.name,
                        dataset_name=dataset_name,
                        modification_type=(
                            "Snowflake.Databases Source"
                            if len(dataset_source_changes) == 1
                            else f"Snowflake.Databases Source #{index}"
                        ),
                        status=status,
                        before_server=change.before_server,
                        after_server=change.after_server,
                        before_warehouse=change.before_warehouse,
                        after_warehouse=change.after_warehouse,
                        before_value=(
                            f"Server={change.before_server}; "
                            f"Warehouse={change.before_warehouse}"
                        ),
                        after_value=(
                            f"Server={change.after_server}; "
                            f"Warehouse={change.after_warehouse}"
                        ),
                        before_code_snapshot=change.before_code,
                        after_code_snapshot=change.after_code,
                        input_path=str(input_file.resolve()),
                        output_path=str(output_file.resolve()),
                    )
                )

            for index, change in enumerate(dataset_path_changes, start=1):
                status = "MODIFIED" if change.changed else "NO_CHANGE_REQUIRED"
                audit_rows.append(
                    AuditRow(
                        timestamp=now,
                        file_name=input_file.name,
                        dataset_name=dataset_name,
                        modification_type=(
                            "ConnectionOverrides.Path"
                            if len(dataset_path_changes) == 1
                            else f"ConnectionOverrides.Path #{index}"
                        ),
                        status=status,
                        before_server=change.before_server,
                        after_server=change.after_server,
                        before_warehouse=change.before_warehouse,
                        after_warehouse=change.after_warehouse,
                        before_value=f"{change.before_server};{change.before_warehouse}",
                        after_value=f"{change.after_server};{change.after_warehouse}",
                        before_code_snapshot=change.before_code,
                        after_code_snapshot=change.after_code,
                        input_path=str(input_file.resolve()),
                        output_path=str(output_file.resolve()),
                    )
                )
        else:
            audit_rows.append(
                AuditRow(
                    timestamp=now,
                    file_name=input_file.name,
                    dataset_name=dataset_name,
                    modification_type="Snowflake.Databases Source",
                    status="NO_SNOWFLAKE_SOURCE_MATCH",
                    input_path=str(input_file.resolve()),
                    output_path=str(output_file.resolve()),
                )
            )

    parts.append(xml_text[cursor:])
    return ModifyResult(
        text="".join(parts),
        audit_rows=audit_rows,
        source_match_count=source_match_count,
        source_change_count=source_change_count,
        path_match_count=path_match_count,
        path_change_count=path_change_count,
    )


# -----------------------------------------------------------------------------
# Excel audit log
# -----------------------------------------------------------------------------
def excel_safe(value: object) -> object:
    """Prevent strings from accidentally being interpreted as Excel formulas."""
    if isinstance(value, str) and value and value[0] in ("=", "+", "-", "@"):
        return "'" + value
    return value


def save_audit_rows_atomic(log_path: Path, rows: Iterable[AuditRow]) -> None:
    rows = list(rows)
    if not rows:
        return

    log_path.parent.mkdir(parents=True, exist_ok=True)

    if log_path.exists():
        workbook = load_workbook(log_path)
        sheet = (
            workbook["ModificationLog"]
            if "ModificationLog" in workbook.sheetnames
            else workbook.create_sheet("ModificationLog")
        )

        existing_headers = [sheet.cell(1, col).value for col in range(1, len(AUDIT_HEADERS) + 1)]
        if sheet.max_row == 1 and sheet.cell(1, 1).value is None:
            sheet.delete_rows(1, 1)
            sheet.append(AUDIT_HEADERS)
        elif existing_headers != AUDIT_HEADERS:
            workbook.close()
            raise ValueError(
                "The existing audit workbook uses an older/different column layout. "
                "Rename or remove it once so the updater can create the current audit schema."
            )
    else:
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "ModificationLog"
        sheet.append(AUDIT_HEADERS)

    for row in rows:
        sheet.append([excel_safe(v) for v in row.as_list()])

    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    for cell in sheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions

    widths = {
        1: 24,   # Timestamp
        2: 30,   # File Name
        3: 30,   # Dataset Name
        4: 34,   # Modification Type
        5: 24,   # Status
        6: 36,   # Before Server
        7: 36,   # After Server
        8: 24,   # Before Warehouse
        9: 24,   # After Warehouse
        10: 52,  # Before Value
        11: 52,  # After Value
        12: 70,  # Before Code
        13: 70,  # After Code
        14: 45,  # Input Path
        15: 45,  # Output Path
        16: 66,  # Original SHA256
        17: 66,  # Modified SHA256
    }
    for col_idx, width in widths.items():
        sheet.column_dimensions[get_column_letter(col_idx)].width = width

    for row in sheet.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    fd, temp_name = tempfile.mkstemp(
        prefix=log_path.stem + "_",
        suffix=".xlsx",
        dir=str(log_path.parent),
    )
    os.close(fd)
    temp_path = Path(temp_name)
    try:
        workbook.save(temp_path)
        os.replace(temp_path, log_path)
    finally:
        workbook.close()
        if temp_path.exists():
            temp_path.unlink(missing_ok=True)


# -----------------------------------------------------------------------------
# File processing
# -----------------------------------------------------------------------------
def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def write_atomic(path: Path, data: bytes, overwrite: bool) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists() and not overwrite:
        raise FileExistsError(
            f"Output file already exists: {path}. Use --overwrite to replace it."
        )

    fd, temp_name = tempfile.mkstemp(
        prefix=path.stem + "_",
        suffix=path.suffix,
        dir=str(path.parent),
    )
    os.close(fd)
    temp_path = Path(temp_name)
    try:
        temp_path.write_bytes(data)
        os.replace(temp_path, path)
    finally:
        if temp_path.exists():
            temp_path.unlink(missing_ok=True)


def process_rdl_file(
    input_file: Path,
    output_dir: Path,
    log_path: Path,
    new_server: str,
    new_warehouse: str,
    overwrite: bool = False,
) -> FileProcessResult:
    output_file = output_dir / input_file.name
    original_bytes = input_file.read_bytes()
    original_hash = sha256_bytes(original_bytes)
    xml_text, encoding, bom = decode_xml_bytes(original_bytes)

    # XML validation + authoritative dataset list.
    dataset_names = validate_and_list_datasets(xml_text)

    print(f"\nFile: {input_file.name}")
    print("DataSets found:")
    if dataset_names:
        for index, name in enumerate(dataset_names, start=1):
            print(f"  {index}. {name}")
    else:
        print("  (none)")

    result = modify_rdl_text(
        xml_text=xml_text,
        new_server=new_server,
        new_warehouse=new_warehouse,
        input_file=input_file,
        output_file=output_file,
    )

    # No Snowflake.Databases source exists anywhere in this RDL.
    if result.source_match_count == 0:
        for row in result.audit_rows:
            row.original_sha256 = original_hash
        save_audit_rows_atomic(log_path, result.audit_rows)
        print("  No matching Snowflake.Databases(server, warehouse, ...) source was found.")
        print("  Original file was NOT deleted.")
        return FileProcessResult(
            input_file=input_file,
            output_file=None,
            dataset_names=dataset_names,
            source_match_count=0,
            source_change_count=0,
            path_change_count=0,
            deleted_original=False,
        )

    modified_bytes = encode_xml_text(result.text, encoding, bom)
    modified_hash = sha256_bytes(modified_bytes)

    # Validate the transformed RDL before permanent file operations.
    validate_and_list_datasets(result.text)

    for row in result.audit_rows:
        row.original_sha256 = original_hash
        row.modified_sha256 = modified_hash

    output_written = False
    try:
        # 1. Write modified/pass-through RDL atomically.
        write_atomic(output_file, modified_bytes, overwrite=overwrite)
        output_written = True

        # 2. Write audit log atomically.
        save_audit_rows_atomic(log_path, result.audit_rows)

        # 3. Remove original only after both writes succeeded.
        input_file.unlink()
    except Exception:
        # Keep the source file safe. Roll back an output created by this failed run.
        if output_written and output_file.exists():
            output_file.unlink(missing_ok=True)
        raise

    print(f"  Snowflake.Databases source(s) found: {result.source_match_count}")
    print(f"  Source value change(s): {result.source_change_count}")
    print(f"  ConnectionOverrides.Path match(es): {result.path_match_count}")
    print(f"  ConnectionOverrides.Path change(s): {result.path_change_count}")
    print(f"  Output: {output_file}")
    print(f"  Audit log: {log_path}")
    print("  Original input file deleted after successful output + audit save.")

    return FileProcessResult(
        input_file=input_file,
        output_file=output_file,
        dataset_names=dataset_names,
        source_match_count=result.source_match_count,
        source_change_count=result.source_change_count,
        path_change_count=result.path_change_count,
        deleted_original=True,
    )


def load_config(path: Path) -> dict:
    if not path.exists():
        return {}
    with path.open("r", encoding="utf-8") as handle:
        data = json.load(handle)
    if not isinstance(data, dict):
        raise ValueError("config.json must contain a JSON object.")
    return data


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Update Snowflake.Databases server/warehouse values in Power BI RDL files, "
            "synchronize matching ConnectionOverrides.Path metadata, write modified files "
            "to an output folder, and maintain an Excel audit log."
        )
    )
    parser.add_argument("--config", default="config.json", help="Path to JSON configuration file.")
    parser.add_argument("--input-dir", help="Folder containing incoming .rdl files.")
    parser.add_argument("--output-dir", help="Folder for modified .rdl files.")
    parser.add_argument("--audit-file", help="Excel audit log path (.xlsx).")
    parser.add_argument("--server", help="New Snowflake connection server name.")
    parser.add_argument("--warehouse", help="New Snowflake warehouse name.")
    parser.add_argument("--overwrite", action="store_true", help="Allow replacing an existing output RDL.")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    config = load_config(Path(args.config))

    input_dir = Path(args.input_dir or config.get("input_dir", "input"))
    output_dir = Path(args.output_dir or config.get("output_dir", "output"))
    audit_file = Path(args.audit_file or config.get("audit_file", "logs/rdl_modification_log.xlsx"))
    server = args.server or config.get("snowflake_server")
    warehouse = args.warehouse or config.get("snowflake_warehouse")
    overwrite = bool(args.overwrite or config.get("overwrite", False))

    if not server:
        print(
            "ERROR: Snowflake server is required. Set snowflake_server in config.json or use --server.",
            file=sys.stderr,
        )
        return 2
    if not warehouse:
        print(
            "ERROR: Snowflake warehouse is required. Set snowflake_warehouse in config.json or use --warehouse.",
            file=sys.stderr,
        )
        return 2

    validate_target_value(server, "Snowflake server")
    validate_target_value(warehouse, "Snowflake warehouse")

    input_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)
    audit_file.parent.mkdir(parents=True, exist_ok=True)

    rdl_files = sorted(input_dir.glob("*.rdl"))
    if not rdl_files:
        print(f"No .rdl files found in: {input_dir.resolve()}")
        return 0

    failures = 0
    print(f"Found {len(rdl_files)} RDL file(s).")
    print(f"Target Snowflake server: {server}")
    print(f"Target Snowflake warehouse: {warehouse}")

    for rdl_file in rdl_files:
        try:
            process_rdl_file(
                input_file=rdl_file,
                output_dir=output_dir,
                log_path=audit_file,
                new_server=server,
                new_warehouse=warehouse,
                overwrite=overwrite,
            )
        except Exception as exc:
            failures += 1
            print(f"\nERROR processing {rdl_file.name}: {exc}", file=sys.stderr)
            print("Original file was kept in the input folder.", file=sys.stderr)

    if failures:
        print(f"\nCompleted with {failures} failure(s).", file=sys.stderr)
        return 1

    print("\nAll RDL files processed successfully.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
