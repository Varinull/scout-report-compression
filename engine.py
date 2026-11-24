"""Deterministic ROI, CSV Reconciliation, XLSX Inspector & Local Folder Profiler for Niche 4.

local-only, local-first evaluation engine using only the Python standard library.
Produces the canonical Sprint 4 Diagnostic Capsule (`niche_id=4`) and routes diagnostic
"""
from __future__ import annotations

import base64
import csv
import io
import json
import math
import os
import re
import sys
import xml.etree.ElementTree as ET
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple

ROOT_DIR = Path(__file__).resolve().parent.parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

SIGNAL_ROUTING_RULES = {
    k: []
    for k in (
        "unclassified_pii_or_secrets",
        "overprivileged_agent_or_db_roles",
        "vector_or_saas_exposure",
        "shadow_ai_unapproved_tools",
        "audit_logging_or_retention_gap",
        "regulatory_or_policy_drift",
        "hallucination_or_ungrounded_output",
        "missing_observability_or_hitl",
        "broken_rag_schema_or_data_hygiene",
        "manual_csv_reconciliation_bottleneck",
        "unstructured_pdf_or_doc_intake",
        "multi_cloud_config_or_telemetry_sprawl",
    )
}


NICHE_ID = 4
NICHE_SLUG = "sprint-4-report-compression"
NICHE_TITLE = "Reporting & Data-Prep Bottleneck Compression (Typically 2-3 Weeks)"
CAPSULE_VERSION = "1.0"

KEY_COL_PATTERN = re.compile(
    r"(?:^|[_\-\s])(id|key|code|acct|account|inv|invoice|cust|customer|tenant|sku|order|ref|record|tx|txn|sub)(?:$|[_\-\s])"
    r"|^(id|record_key|account_id|acct_id|invoice_id|customer_id|tenant_id|order_id|sku)$",
    re.IGNORECASE,
)

LOOKUP_FORMULA_PATTERN = re.compile(
    r"\b(VLOOKUP|XLOOKUP|HLOOKUP|INDEX|MATCH|INDIRECT|OFFSET|SUMIFS|COUNTIFS)\s*\(",
    re.IGNORECASE,
)

SKIP_DIRS = {".git", ".hg", ".svn", "node_modules", "__pycache__", ".venv", "venv", ".mypy_cache", ".pytest_cache"}


def _iso_now() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def _safe_float(val: Any, default: float = 0.0) -> float:
    if val is None:
        return default
    if isinstance(val, (int, float)):
        return float(val) if math.isfinite(float(val)) else default
    text = str(val).strip()
    if not text:
        return default
    cleaned = re.sub(r"[$,%\s]", "", text)
    if cleaned.startswith("(") and cleaned.endswith(")"):
        cleaned = "-" + cleaned[1:-1]
    try:
        num = float(cleaned)
        return num if math.isfinite(num) else default
    except ValueError:
        return default


def _safe_bool(val: Any) -> bool:
    if isinstance(val, bool):
        return val
    return str(val or "").strip().lower() in {"true", "1", "yes", "y", "checked"}


def _normalize_key_stem(col_name: str) -> str:
    s = re.sub(r"[^a-z0-9]+", "_", col_name.strip().lower()).strip("_")
    synonyms = {
        "acct": "account",
        "cust": "customer",
        "inv": "invoice",
        "txn": "transaction",
        "tx": "transaction",
        "num": "id",
        "number": "id",
        "code": "id",
        "key": "id",
    }
    parts = [synonyms.get(p, p) for p in s.split("_") if p]
    return "_".join(parts)


def _is_candidate_key_column(col_name: str) -> bool:
    if not col_name or not col_name.strip():
        return False
    return bool(KEY_COL_PATTERN.search(col_name.strip()))


def _col_letters_to_index(cell_ref: str) -> int:
    letters = "".join(ch for ch in cell_ref if ch.isalpha()).upper()
    idx = 0
    for ch in letters:
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return max(0, idx - 1)


def _extract_xlsx_data(data: bytes, collect_keys: bool = False) -> Tuple[Dict[str, Any], Dict[str, Set[str]]]:
    """Inspect `.xlsx` bytes via stdlib `zipfile` + `xml.etree.ElementTree` (read-only)."""
    key_sets_by_col: Dict[str, Set[str]] = {}
    if not data:
        return {
            "valid_xlsx": False,
            "error": "Empty byte payload",
            "sheet_count": 0,
            "total_rows": 0,
            "formula_count": 0,
            "lookup_formula_count": 0,
            "sheets": [],
        }, key_sets_by_col

    try:
        zf = zipfile.ZipFile(io.BytesIO(data), mode="r")
    except (zipfile.BadZipFile, OSError) as exc:
        return {
            "valid_xlsx": False,
            "error": f"Invalid XLSX archive: {exc}",
            "sheet_count": 0,
            "total_rows": 0,
            "formula_count": 0,
            "lookup_formula_count": 0,
            "sheets": [],
        }, key_sets_by_col

    with zf:
        names = set(zf.namelist())
        if "xl/workbook.xml" not in names:
            return {
                "valid_xlsx": False,
                "error": "Archive missing xl/workbook.xml",
                "sheet_count": 0,
                "total_rows": 0,
                "formula_count": 0,
                "lookup_formula_count": 0,
                "sheets": [],
            }, key_sets_by_col

        shared_strings: List[str] = []
        if "xl/sharedStrings.xml" in names:
            try:
                ss_root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
                for si in ss_root.iter():
                    if si.tag.endswith("}si") or si.tag == "si":
                        chunks = [
                            (node.text or "")
                            for node in si.iter()
                            if node.tag.endswith("}t") or node.tag == "t"
                        ]
                        shared_strings.append("".join(chunks))
            except ET.ParseError:
                pass

        rels_map: Dict[str, str] = {}
        if "xl/_rels/workbook.xml.rels" in names:
            try:
                rels_root = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
                for rel in rels_root.iter():
                    if rel.tag.endswith("}Relationship") or rel.tag == "Relationship":
                        r_id = rel.attrib.get("Id", "")
                        target = rel.attrib.get("Target", "").lstrip("/")
                        if target and not target.startswith("xl/"):
                            target = f"xl/{target}"
                        if r_id and target:
                            rels_map[r_id] = target
            except ET.ParseError:
                pass

        sheet_entries: List[Tuple[str, str]] = []
        try:
            wb_root = ET.fromstring(zf.read("xl/workbook.xml"))
            sheet_idx = 0
            for elem in wb_root.iter():
                if elem.tag.endswith("}sheet") or elem.tag == "sheet":
                    sheet_idx += 1
                    s_name = elem.attrib.get("name", f"Sheet{sheet_idx}")
                    r_id = ""
                    for k, v in elem.attrib.items():
                        if k.endswith("}id") or k == "r:id" or k == "id":
                            r_id = v
                            break
                    target = rels_map.get(r_id, f"xl/worksheets/sheet{sheet_idx}.xml")
                    if target not in names:
                        alt = f"xl/worksheets/sheet{sheet_idx}.xml"
                        target = alt if alt in names else ""
                    sheet_entries.append((s_name, target))
        except ET.ParseError as exc:
            return {
                "valid_xlsx": False,
                "error": f"Corrupt workbook.xml: {exc}",
                "sheet_count": 0,
                "total_rows": 0,
                "formula_count": 0,
                "lookup_formula_count": 0,
                "sheets": [],
            }, key_sets_by_col

        sheets_out: List[Dict[str, Any]] = []
        total_rows = 0
        total_formulas = 0
        total_lookup_formulas = 0

        for s_name, target in sheet_entries:
            if not target or target not in names:
                sheets_out.append(
                    {
                        "name": s_name,
                        "row_count": 0,
                        "data_row_count": 0,
                        "column_count": 0,
                        "headers": [],
                        "formula_count": 0,
                        "lookup_formula_count": 0,
                        "null_rate_pct": 0.0,
                        "duplicate_key_counts": {},
                    }
                )
                continue

            try:
                ws_root = ET.fromstring(zf.read(target))
            except ET.ParseError:
                continue

            row_count = 0
            formula_count = 0
            lookup_formula_count = 0
            headers: List[str] = []
            non_empty_cells = 0
            key_col_indices: Dict[int, str] = {}
            seen_keys_per_col: Dict[str, Set[str]] = {}
            dup_counts_per_col: Dict[str, int] = {}

            for row_elem in ws_root.iter():
                if not (row_elem.tag.endswith("}row") or row_elem.tag == "row"):
                    continue
                row_count += 1
                row_vals: Dict[int, str] = {}
                col_cursor = 0
                for c_elem in row_elem:
                    if not (c_elem.tag.endswith("}c") or c_elem.tag == "c"):
                        continue
                    ref = c_elem.attrib.get("r", "")
                    c_idx = _col_letters_to_index(ref) if ref else col_cursor
                    col_cursor = c_idx + 1
                    c_type = c_elem.attrib.get("t", "")
                    val_str = ""
                    for child in c_elem:
                        tag = child.tag
                        if tag.endswith("}f") or tag == "f":
                            formula_count += 1
                            f_txt = child.text or ""
                            if LOOKUP_FORMULA_PATTERN.search(f_txt):
                                lookup_formula_count += 1
                        elif tag.endswith("}v") or tag == "v":
                            raw_v = child.text or ""
                            if c_type == "s":
                                try:
                                    s_i = int(raw_v)
                                    val_str = shared_strings[s_i] if 0 <= s_i < len(shared_strings) else raw_v
                                except ValueError:
                                    val_str = raw_v
                            else:
                                val_str = raw_v
                        elif tag.endswith("}is") or tag == "is":
                            val_str = "".join(
                                (t_node.text or "")
                                for t_node in child.iter()
                                if t_node.tag.endswith("}t") or t_node.tag == "t"
                            )
                    val_clean = val_str.strip()
                    if val_clean:
                        row_vals[c_idx] = val_clean

                if row_count == 1 and row_vals:
                    max_c = max(row_vals.keys())
                    headers = [row_vals.get(i, f"col_{i+1}") for i in range(max_c + 1)]
                    for i, h in enumerate(headers):
                        if _is_candidate_key_column(h):
                            key_col_indices[i] = h
                            seen_keys_per_col[h] = set()
                            dup_counts_per_col[h] = 0
                elif row_count > 1:
                    non_empty_cells += len(row_vals)
                    for c_i, col_label in key_col_indices.items():
                        kv = row_vals.get(c_i, "")
                        if kv:
                            if kv in seen_keys_per_col[col_label]:
                                dup_counts_per_col[col_label] += 1
                            else:
                                seen_keys_per_col[col_label].add(kv)
                            if collect_keys:
                                key_sets_by_col.setdefault(col_label, set()).add(kv)

            data_rows = max(0, row_count - 1) if headers else row_count
            total_expected_cells = data_rows * max(1, len(headers))
            null_rate = (
                round(max(0.0, 1.0 - (non_empty_cells / total_expected_cells)) * 100.0, 1)
                if data_rows > 0 and headers
                else 0.0
            )

            total_rows += data_rows
            total_formulas += formula_count
            total_lookup_formulas += lookup_formula_count
            sheets_out.append(
                {
                    "name": s_name,
                    "row_count": row_count,
                    "data_row_count": data_rows,
                    "column_count": len(headers),
                    "headers": headers[:40],
                    "formula_count": formula_count,
                    "lookup_formula_count": lookup_formula_count,
                    "null_rate_pct": null_rate,
                    "duplicate_key_counts": {k: v for k, v in dup_counts_per_col.items() if v > 0},
                }
            )

        return {
            "valid_xlsx": True,
            "sheet_count": len(sheets_out),
            "total_rows": total_rows,
            "formula_count": total_formulas,
            "lookup_formula_count": total_lookup_formulas,
            "sheets": sheets_out,
        }, key_sets_by_col


def inspect_xlsx_bytes(data: bytes) -> Dict[str, Any]:
    """Public stdlib helper to inspect `.xlsx` sheet names, row counts, headers, and lookup formulas."""
    summary, _ = _extract_xlsx_data(data, collect_keys=False)
    return summary


def _estimate_pdf_pages(raw_bytes: bytes) -> int:
    if not raw_bytes:
        return 0
    matches = re.findall(rb"/Type\s*/Page\b", raw_bytes)
    if matches:
        return len(matches)
    return max(1, math.ceil(len(raw_bytes) / 45000))


def scan_local_directory(path: str, max_files: int = 200) -> Dict[str, Any]:
    """Read-only profiler for a local folder of recurring CSV, TSV, XLSX, JSON, and PDF report exports.

    Strictly local-only and zero raw value persistence: returns structural metadata,
    column names, null rates, duplicate key counts, formula counts, and cross-file join-key
    overlap statistics without retaining any customer IDs or financial amounts.
    """
    if not path or not str(path).strip():
        raise ValueError("Please provide a non-empty local directory path.")
    target = Path(str(path).strip()).expanduser().resolve()
    if not target.exists():
        raise FileNotFoundError(f"Directory does not exist: {target}")
    if not target.is_dir():
        raise NotADirectoryError(f"Path is not a directory: {target}")

    supported_exts = {".csv", ".tsv", ".xlsx", ".json", ".pdf"}
    discovered: List[Path] = []
    for root, dirs, files in os.walk(target):
        dirs[:] = sorted(d for d in dirs if d not in SKIP_DIRS and not d.startswith("."))
        for fn in sorted(files):
            if fn.startswith(".") or fn.startswith("~$"):
                continue
            fp = Path(root) / fn
            if fp.suffix.lower() in supported_exts:
                discovered.append(fp)
                if len(discovered) >= max_files:
                    break
        if len(discovered) >= max_files:
            break

    files_profile: List[Dict[str, Any]] = []
    # Temporary in-memory key registry: list of (file_label, col_name, stem, Set[str])
    ephemeral_key_sets: List[Tuple[str, str, str, Set[str]]] = []
    tabular_schemas: List[Tuple[str, List[str]]] = []

    totals = {
        "total_files": 0,
        "csv_tsv_files": 0,
        "xlsx_files": 0,
        "json_files": 0,
        "pdf_files": 0,
        "total_size_bytes": 0,
        "total_tabular_rows": 0,
        "total_pdf_pages_est": 0,
        "total_duplicate_keys": 0,
        "total_excel_formulas": 0,
        "total_lookup_formulas": 0,
        "join_orphan_keys_detected": 0,
        "schema_drift_warnings": 0,
    }

    for fp in discovered:
        ext = fp.suffix.lower()
        rel_name = str(fp.relative_to(target)).replace("\\", "/")
        try:
            size_bytes = fp.stat().st_size
        except OSError:
            continue

        totals["total_files"] += 1
        totals["total_size_bytes"] += size_bytes

        if ext in {".csv", ".tsv"}:
            totals["csv_tsv_files"] += 1
            delim = "\t" if ext == ".tsv" else ","
            headers: List[str] = []
            data_rows = 0
            empty_cells = 0
            total_cells = 0
            key_cols: Dict[int, str] = {}
            seen_keys: Dict[str, Set[str]] = {}
            dup_counts: Dict[str, int] = {}

            try:
                with fp.open("r", encoding="utf-8-sig", errors="replace", newline="") as fh:
                    reader = csv.reader(fh, delimiter=delim)
                    for idx, row in enumerate(reader):
                        if idx == 0:
                            headers = [c.strip() for c in row]
                            for c_i, h in enumerate(headers):
                                if _is_candidate_key_column(h):
                                    key_cols[c_i] = h
                                    seen_keys[h] = set()
                                    dup_counts[h] = 0
                            continue
                        if not any(cell.strip() for cell in row):
                            continue
                        data_rows += 1
                        if data_rows <= 10000:
                            width = max(len(headers), len(row))
                            for c_i in range(len(headers)):
                                val = row[c_i].strip() if c_i < len(row) else ""
                                total_cells += 1
                                if not val:
                                    empty_cells += 1
                                elif c_i in key_cols:
                                    col_name = key_cols[c_i]
                                    if val in seen_keys[col_name]:
                                        dup_counts[col_name] += 1
                                    else:
                                        seen_keys[col_name].add(val)
                            if width > len(headers):
                                total_cells += width - len(headers)
            except OSError:
                continue

            null_rate_pct = round((empty_cells / total_cells) * 100.0, 1) if total_cells > 0 else 0.0
            dup_total = sum(dup_counts.values())
            totals["total_tabular_rows"] += data_rows
            totals["total_duplicate_keys"] += dup_total

            for col_name, kset in seen_keys.items():
                if kset:
                    ephemeral_key_sets.append((rel_name, col_name, _normalize_key_stem(col_name), kset))
            if headers:
                tabular_schemas.append((rel_name, headers))

            files_profile.append(
                {
                    "file": rel_name,
                    "type": ext.lstrip(".").upper(),
                    "size_bytes": size_bytes,
                    "row_count": data_rows,
                    "column_count": len(headers),
                    "headers": headers[:40],
                    "null_rate_pct": null_rate_pct,
                    "candidate_key_columns": list(key_cols.values()),
                    "duplicate_key_counts": {k: v for k, v in dup_counts.items() if v > 0},
                }
            )

        elif ext == ".xlsx":
            totals["xlsx_files"] += 1
            try:
                raw_bytes = fp.read_bytes()
            except OSError:
                continue
            xlsx_info, xlsx_keys = _extract_xlsx_data(raw_bytes, collect_keys=True)
            totals["total_tabular_rows"] += int(xlsx_info.get("total_rows", 0))
            totals["total_excel_formulas"] += int(xlsx_info.get("formula_count", 0))
            totals["total_lookup_formulas"] += int(xlsx_info.get("lookup_formula_count", 0))

            dup_merged: Dict[str, int] = {}
            all_headers: List[str] = []
            max_null_rate = 0.0
            for sh in xlsx_info.get("sheets", []):
                sh_headers = sh.get("headers", [])
                if sh_headers:
                    tabular_schemas.append((f"{rel_name}#{sh.get('name', 'Sheet')}", sh_headers))
                    for h in sh_headers:
                        if h not in all_headers:
                            all_headers.append(h)
                max_null_rate = max(max_null_rate, float(sh.get("null_rate_pct", 0.0)))
                for k, v in sh.get("duplicate_key_counts", {}).items():
                    dup_merged[k] = dup_merged.get(k, 0) + int(v)
                    totals["total_duplicate_keys"] += int(v)

            for col_name, kset in xlsx_keys.items():
                if kset:
                    ephemeral_key_sets.append((rel_name, col_name, _normalize_key_stem(col_name), kset))

            files_profile.append(
                {
                    "file": rel_name,
                    "type": "XLSX",
                    "size_bytes": size_bytes,
                    "sheet_count": xlsx_info.get("sheet_count", 0),
                    "sheet_names": [s.get("name", "") for s in xlsx_info.get("sheets", [])],
                    "row_count": xlsx_info.get("total_rows", 0),
                    "column_count": len(all_headers),
                    "headers": all_headers[:40],
                    "formula_count": xlsx_info.get("formula_count", 0),
                    "lookup_formula_count": xlsx_info.get("lookup_formula_count", 0),
                    "null_rate_pct": max_null_rate,
                    "duplicate_key_counts": dup_merged,
                }
            )

        elif ext == ".json":
            totals["json_files"] += 1
            record_count = 0
            headers = []
            empty_cells = 0
            total_cells = 0
            seen_keys_json: Dict[str, Set[str]] = {}
            dup_counts_json: Dict[str, int] = {}
            try:
                parsed = json.loads(fp.read_text(encoding="utf-8", errors="replace"))
                records: List[Dict[str, Any]] = []
                if isinstance(parsed, list):
                    records = [r for r in parsed if isinstance(r, dict)]
                elif isinstance(parsed, dict):
                    for candidate_list_key in ("records", "data", "items", "rows", "entries", "line_items"):
                        if isinstance(parsed.get(candidate_list_key), list):
                            records = [r for r in parsed[candidate_list_key] if isinstance(r, dict)]
                            break
                    if not records:
                        records = [parsed]
                record_count = len(records)
                header_set: List[str] = []
                for r in records[:2000]:
                    for k in r.keys():
                        if k not in header_set:
                            header_set.append(str(k))
                headers = header_set
                key_fields = [h for h in headers if _is_candidate_key_column(h)]
                for kf in key_fields:
                    seen_keys_json[kf] = set()
                    dup_counts_json[kf] = 0
                for r in records[:5000]:
                    for h in headers:
                        total_cells += 1
                        v = r.get(h)
                        if v is None or (isinstance(v, str) and not v.strip()):
                            empty_cells += 1
                        elif h in seen_keys_json and isinstance(v, (str, int, float)):
                            sv = str(v).strip()
                            if sv in seen_keys_json[h]:
                                dup_counts_json[h] += 1
                            else:
                                seen_keys_json[h].add(sv)
            except (OSError, json.JSONDecodeError):
                pass

            null_rate_pct = round((empty_cells / total_cells) * 100.0, 1) if total_cells > 0 else 0.0
            totals["total_tabular_rows"] += record_count
            totals["total_duplicate_keys"] += sum(dup_counts_json.values())
            for col_name, kset in seen_keys_json.items():
                if kset:
                    ephemeral_key_sets.append((rel_name, col_name, _normalize_key_stem(col_name), kset))
            if headers:
                tabular_schemas.append((rel_name, headers))

            files_profile.append(
                {
                    "file": rel_name,
                    "type": "JSON",
                    "size_bytes": size_bytes,
                    "row_count": record_count,
                    "column_count": len(headers),
                    "headers": headers[:40],
                    "null_rate_pct": null_rate_pct,
                    "candidate_key_columns": list(seen_keys_json.keys()),
                    "duplicate_key_counts": {k: v for k, v in dup_counts_json.items() if v > 0},
                }
            )

        elif ext == ".pdf":
            totals["pdf_files"] += 1
            try:
                raw_pdf = fp.read_bytes()
            except OSError:
                raw_pdf = b""
            pages_est = _estimate_pdf_pages(raw_pdf)
            totals["total_pdf_pages_est"] += pages_est
            files_profile.append(
                {
                    "file": rel_name,
                    "type": "PDF",
                    "size_bytes": size_bytes,
                    "page_count_est": pages_est,
                }
            )

    # Cross-file join-key overlap & orphan detection (discarding raw key values immediately)
    join_key_overlaps: List[Dict[str, Any]] = []
    schema_warnings: List[str] = []

    for i in range(len(ephemeral_key_sets)):
        f_a, col_a, stem_a, set_a = ephemeral_key_sets[i]
        for j in range(i + 1, len(ephemeral_key_sets)):
            f_b, col_b, stem_b, set_b = ephemeral_key_sets[j]
            if f_a == f_b:
                continue
            inter = len(set_a & set_b)
            same_stem = (stem_a == stem_b) or (col_a.lower() == col_b.lower())
            if inter == 0 and not same_stem:
                continue
            union_sz = len(set_a | set_b)
            if union_sz == 0:
                continue
            overlap_pct = round((inter / union_sz) * 100.0, 1)
            # Only report pairs that either actually share key values or share the same key name/stem
            if inter > 0 or same_stem:
                orphans_a = len(set_a - set_b)
                orphans_b = len(set_b - set_a)
                if inter > 0:
                    totals["join_orphan_keys_detected"] += orphans_a + orphans_b
                header_mismatch = col_a != col_b
                if header_mismatch and inter > 0:
                    schema_warnings.append(
                        f"Join-key header drift between `{f_a}` (`{col_a}`) and `{f_b}` (`{col_b}`), {inter} shared IDs require canonical key aliasing."
                    )
                join_key_overlaps.append(
                    {
                        "source_a": f"{f_a} [{col_a}]",
                        "source_b": f"{f_b} [{col_b}]",
                        "header_name_mismatch": header_mismatch,
                        "source_a_unique_keys": len(set_a),
                        "source_b_unique_keys": len(set_b),
                        "matched_keys": inter,
                        "orphan_in_a": orphans_a,
                        "orphan_in_b": orphans_b,
                        "jaccard_overlap_pct": overlap_pct,
                    }
                )

    # Check partial schema drift across similar tabular exports
    for i in range(len(tabular_schemas)):
        name_a, cols_a = tabular_schemas[i]
        norm_a = {c.lower(): c for c in cols_a if c}
        if len(norm_a) < 2:
            continue
        for j in range(i + 1, len(tabular_schemas)):
            name_b, cols_b = tabular_schemas[j]
            norm_b = {c.lower(): c for c in cols_b if c}
            if len(norm_b) < 2:
                continue
            shared = set(norm_a.keys()) & set(norm_b.keys())
            union_cols = set(norm_a.keys()) | set(norm_b.keys())
            ratio = len(shared) / len(union_cols) if union_cols else 0.0
            if 0.45 <= ratio < 1.0:
                only_a = [norm_a[k] for k in sorted(set(norm_a.keys()) - set(norm_b.keys()))][:5]
                only_b = [norm_b[k] for k in sorted(set(norm_b.keys()) - set(norm_a.keys()))][:5]
                schema_warnings.append(
                    f"Partial column drift between `{name_a}` and `{name_b}` ({len(shared)}/{len(union_cols)} columns match; only in first: {only_a}; only in second: {only_b})."
                )

    totals["schema_drift_warnings"] = len(schema_warnings)
    ephemeral_key_sets.clear()

    # Infer helpful suggested calculator inputs from the scanned files
    suggested_formats: List[str] = []
    if totals["csv_tsv_files"] > 0 or totals["xlsx_files"] > 0:
        suggested_formats.append("csv_excel_exports")
    if totals["pdf_files"] > 0:
        suggested_formats.append("pdf_invoices_or_contracts")
    lower_names = " ".join(f["file"].lower() for f in files_profile)
    if totals["json_files"] > 0 or any(
        kw in lower_names for kw in ("aws", "gcp", "azure", "cloud", "billing", "snowflake", "datadog", "stripe", "usage")
    ):
        suggested_formats.append("cloud_or_saas_billing")

    suggested_bottlenecks: List[str] = []
    if totals["csv_tsv_files"] + totals["xlsx_files"] >= 2 or totals["total_lookup_formulas"] > 0:
        suggested_bottlenecks.append("manual_copy_paste_joins")
    if totals["join_orphan_keys_detected"] > 0 or totals["total_duplicate_keys"] > 0:
        suggested_bottlenecks.append("reconciliation_mismatches")
    if totals["schema_drift_warnings"] > 0:
        suggested_bottlenecks.append("schema_drift_or_dirty_exports")
    suggested_bottlenecks.append("no_signoff_audit_trail")

    return {
        "scanned_path": str(target),
        "scanned_at": _iso_now(),
        "zero_raw_values_exported": True,
        "totals": totals,
        "files": files_profile,
        "join_key_overlaps": join_key_overlaps[:25],
        "schema_warnings": schema_warnings[:20],
        "suggested_inputs": {
            "exports_per_run": totals["total_files"],
            "source_formats": suggested_formats,
            "bottleneck_types": suggested_bottlenecks,
        },
    }


def parse_workflow_steps_csv(csv_text: str, target_reduction_pct: float = 50.0) -> List[Dict[str, Any]]:
    """Parse `workflow_steps_csv_text` (`step_order,step_name,manual_minutes,is_copy_paste,error_prone,source_type`)."""
    text = (csv_text or "").strip()
    if not text:
        return []
    reader = csv.DictReader(io.StringIO(text))
    if not reader.fieldnames:
        return []

    target_ratio = max(0.10, min(0.85, target_reduction_pct / 100.0))
    steps: List[Dict[str, Any]] = []
    for idx, row in enumerate(reader, start=1):
        norm = {(k or "").strip().lower(): (v or "").strip() for k, v in row.items()}
        step_name = norm.get("step_name") or norm.get("step") or norm.get("task") or ""
        minutes = _safe_float(norm.get("manual_minutes") or norm.get("minutes"), 0.0)
        if not step_name and minutes <= 0:
            continue
        is_copy_paste = _safe_bool(norm.get("is_copy_paste"))
        error_prone = _safe_bool(norm.get("error_prone"))
        source_type = (norm.get("source_type") or "csv_excel_exports").lower()
        order = int(_safe_float(norm.get("step_order"), float(idx)))

        if "human_review" in source_type or (not is_copy_paste and not error_prone):
            comp_rate = 0.0
            automation_mode = "Human Sign-Off Step (Time Kept)"
        elif is_copy_paste and error_prone:
            comp_rate = target_ratio
            automation_mode = "Automated Import & ID Matching"
        elif is_copy_paste and not error_prone:
            comp_rate = target_ratio
            automation_mode = "Automated File Formatting"
        else:
            comp_rate = target_ratio
            if "pdf" in source_type or "doc" in source_type or "contract" in source_type:
                automation_mode = "Structured PDF Field Extraction + Check"
            else:
                automation_mode = "Pre-Calculated Summary Table + Human Review"

        saved_min = round(minutes * comp_rate, 2)
        remaining_min = round(max(0.0, minutes - saved_min), 2)
        steps.append(
            {
                "step_order": order,
                "step_name": step_name or f"Step {order}",
                "manual_minutes": round(minutes, 1),
                "is_copy_paste": is_copy_paste,
                "error_prone": error_prone,
                "source_type": source_type,
                "compression_rate": comp_rate,
                "saved_minutes": saved_min,
                "post_sprint_minutes": remaining_min,
                "automation_pattern": automation_mode,
            }
        )
    return steps


def parse_reconciliation_csv(csv_text: str, tolerance: float = 0.01) -> Dict[str, Any]:
    """Parse `reconciliation_csv_text` (`record_key,source_a_amount,source_b_amount,status_note`)."""
    text = (csv_text or "").strip()
    if not text:
        return {
            "rows_checked": 0,
            "matched_count": 0,
            "mismatch_count": 0,
            "missing_in_a_count": 0,
            "missing_in_b_count": 0,
            "amount_variance_count": 0,
            "missing_key_count": 0,
            "duplicate_key_count": 0,
            "total_abs_variance_usd": 0.0,
            "discrepancies": [],
        }

    reader = csv.DictReader(io.StringIO(text))
    if not reader.fieldnames:
        return {
            "rows_checked": 0,
            "matched_count": 0,
            "mismatch_count": 0,
            "missing_in_a_count": 0,
            "missing_in_b_count": 0,
            "amount_variance_count": 0,
            "missing_key_count": 0,
            "duplicate_key_count": 0,
            "total_abs_variance_usd": 0.0,
            "discrepancies": [],
        }

    rows_checked = 0
    matched_count = 0
    mismatch_count = 0
    missing_a = 0
    missing_b = 0
    variance_cnt = 0
    missing_keys = 0
    dup_keys = 0
    seen_keys: Set[str] = set()
    total_abs_variance = 0.0
    discrepancies: List[Dict[str, Any]] = []

    for idx, row in enumerate(reader, start=1):
        norm = {(k or "").strip().lower(): (v or "").strip() for k, v in row.items()}
        if not any(norm.values()):
            continue
        rows_checked += 1
        rec_key = norm.get("record_key") or norm.get("id") or norm.get("key") or ""
        raw_a = norm.get("source_a_amount", "")
        raw_b = norm.get("source_b_amount", "")
        note = norm.get("status_note", "")

        if not rec_key:
            missing_keys += 1
            rec_key_display = f"<BLANK_KEY_ROW_{idx}>"
        else:
            rec_key_display = rec_key
            if rec_key in seen_keys:
                dup_keys += 1
            else:
                seen_keys.add(rec_key)

        has_a = bool(raw_a)
        has_b = bool(raw_b)
        amt_a = _safe_float(raw_a, 0.0) if has_a else 0.0
        amt_b = _safe_float(raw_b, 0.0) if has_b else 0.0
        delta = round(amt_a - amt_b, 2)

        if not has_a and has_b:
            mismatch_count += 1
            missing_a += 1
            total_abs_variance += abs(amt_b)
            discrepancies.append(
                {
                    "record_key": rec_key_display,
                    "issue": "MISSING_IN_SOURCE_A",
                    "source_a_amount": None,
                    "source_b_amount": amt_b,
                    "delta_usd": round(-amt_b, 2),
                    "status_note": note or "Present in Source B only",
                }
            )
        elif has_a and not has_b:
            mismatch_count += 1
            missing_b += 1
            total_abs_variance += abs(amt_a)
            discrepancies.append(
                {
                    "record_key": rec_key_display,
                    "issue": "MISSING_IN_SOURCE_B",
                    "source_a_amount": amt_a,
                    "source_b_amount": None,
                    "delta_usd": round(amt_a, 2),
                    "status_note": note or "Present in Source A only",
                }
            )
        elif abs(delta) > tolerance:
            mismatch_count += 1
            variance_cnt += 1
            total_abs_variance += abs(delta)
            discrepancies.append(
                {
                    "record_key": rec_key_display,
                    "issue": "AMOUNT_VARIANCE",
                    "source_a_amount": amt_a,
                    "source_b_amount": amt_b,
                    "delta_usd": delta,
                    "status_note": note or f"Unreconciled delta of ${abs(delta):,.2f}",
                }
            )
        else:
            matched_count += 1

    return {
        "rows_checked": rows_checked,
        "matched_count": matched_count,
        "mismatch_count": mismatch_count,
        "missing_in_a_count": missing_a,
        "missing_in_b_count": missing_b,
        "amount_variance_count": variance_cnt,
        "missing_key_count": missing_keys,
        "duplicate_key_count": dup_keys,
        "total_abs_variance_usd": round(total_abs_variance, 2),
        "discrepancies": discrepancies[:25],
    }


def _compute_runs_per_year(runs_per_month: float) -> float:
    if abs(runs_per_month - 4.33) <= 0.03:
        return 52.0
    if abs(runs_per_month - 2.17) <= 0.03:
        return 26.0
    if abs(runs_per_month - 21.67) <= 0.06:
        return 260.0
    return max(0.0, runs_per_month * 12.0)


def evaluate_intake(payload: Dict[str, Any], mode: str = "production") -> Dict[str, Any]:
    """Deterministically evaluate a Sprint 4 intake payload and emit the canonical Capsule JSON."""
    raw_prospect = payload.get("prospect") if isinstance(payload.get("prospect"), dict) else {}
    raw_inputs = payload.get("inputs") if isinstance(payload.get("inputs"), dict) else payload

    mode_clean = "demo" if str(mode or payload.get("mode") or "").strip().lower() == "demo" else "production"

    prospect = {
        "company_name": str(raw_prospect.get("company_name") or "").strip(),
        "contact_name": str(raw_prospect.get("contact_name") or "").strip(),
        "contact_email": str(raw_prospect.get("contact_email") or "").strip(),
        "industry": str(raw_prospect.get("industry") or "general").strip() or "general",
        "team_size": str(raw_prospect.get("team_size") or "50-250").strip() or "50-250",
    }

    workflow_name = str(raw_inputs.get("workflow_name") or "Recurring Reporting & Reconciliation Workflow").strip()
    runs_per_month = max(0.25, _safe_float(raw_inputs.get("runs_per_month"), 4.0))
    hours_per_run = max(0.25, _safe_float(raw_inputs.get("hours_per_run"), 8.0))
    hourly_rate_usd = max(15.0, _safe_float(raw_inputs.get("hourly_rate_usd"), 85.0))
    people_involved = max(1, int(_safe_float(raw_inputs.get("people_involved"), 2.0)))
    exports_per_run = max(1, int(_safe_float(raw_inputs.get("exports_per_run"), 3.0)))

    source_formats = [
        str(s).strip()
        for s in (raw_inputs.get("source_formats") or [])
        if str(s).strip()
    ]
    bottleneck_types = [
        str(b).strip()
        for b in (raw_inputs.get("bottleneck_types") or [])
        if str(b).strip()
    ]
    workflow_steps_csv_text = str(raw_inputs.get("workflow_steps_csv_text") or "")
    reconciliation_csv_text = str(raw_inputs.get("reconciliation_csv_text") or "")

    # Optional XLSX summary (from `xlsx_base64` or pre-parsed `xlsx_summary`)
    xlsx_summary = raw_inputs.get("xlsx_summary") if isinstance(raw_inputs.get("xlsx_summary"), dict) else None
    if not xlsx_summary and raw_inputs.get("xlsx_base64"):
        try:
            decoded_xlsx = base64.b64decode(str(raw_inputs["xlsx_base64"]))
            xlsx_summary = inspect_xlsx_bytes(decoded_xlsx)
        except Exception:
            xlsx_summary = None

    local_scan_summary = (
        raw_inputs.get("local_scan_summary")
        if isinstance(raw_inputs.get("local_scan_summary"), dict)
        else (payload.get("local_scan_summary") if isinstance(payload.get("local_scan_summary"), dict) else None)
    )

    target_reduction_pct = int(round(max(10.0, min(85.0, _safe_float(raw_inputs.get("target_reduction_pct"), 50.0)))))
    steps = parse_workflow_steps_csv(workflow_steps_csv_text, float(target_reduction_pct))
    recon = parse_reconciliation_csv(reconciliation_csv_text)

    # Use the user's explicit What-If Target Reduction % (default 50%) — never a hidden hardcoded percentage
    compression_pct = target_reduction_pct
    compression_ratio = compression_pct / 100.0

    runs_per_year = _compute_runs_per_year(runs_per_month)
    current_annual_hours = int(hours_per_run * runs_per_year * 10 + 0.5) / 10.0
    current_annual_cost = int(round(current_annual_hours * hourly_rate_usd))
    exact_reclaimed_hours = current_annual_hours * compression_ratio
    annual_hours_reclaimed = int(exact_reclaimed_hours + 0.5)
    post_sprint_hours_per_run = int(hours_per_run * (1.0 - compression_ratio) * 10 + 0.5) / 10.0
    annual_savings_usd = int(exact_reclaimed_hours * hourly_rate_usd + 0.5)

    # Checklist Priority Score (0 - 100) based on current annual time cost + data complexity + scanned issues
    if current_annual_cost >= 75000:
        savings_pts = 35
    elif current_annual_cost >= 45000:
        savings_pts = 30
    elif current_annual_cost >= 25000:
        savings_pts = 24
    elif current_annual_cost >= 12000:
        savings_pts = 17
    elif current_annual_cost >= 5000:
        savings_pts = 10
    else:
        savings_pts = 5

    comp_pts = min(20, len(bottleneck_types) * 5)
    complexity_pts = min(16, len(source_formats) * 5) + min(10, exports_per_run * 2)

    evidence_pts = 0
    if len(steps) >= 2:
        evidence_pts += 4
    elif local_scan_summary and local_scan_summary.get("totals", {}).get("total_files", 0) >= 2:
        evidence_pts += 4

    local_orphans = int((local_scan_summary or {}).get("totals", {}).get("join_orphan_keys_detected", 0))
    local_dups = int((local_scan_summary or {}).get("totals", {}).get("total_duplicate_keys", 0))
    xlsx_lookups = int((xlsx_summary or {}).get("lookup_formula_count", 0)) + int(
        (local_scan_summary or {}).get("totals", {}).get("total_lookup_formulas", 0)
    )
    schema_warn_count = int((local_scan_summary or {}).get("totals", {}).get("schema_drift_warnings", 0))

    if recon["mismatch_count"] >= 3 or local_orphans >= 3:
        evidence_pts += 7
    elif recon["mismatch_count"] >= 1 or local_orphans >= 1 or local_dups >= 1 or xlsx_lookups >= 3:
        evidence_pts += 4
    elif "reconciliation_mismatches" in bottleneck_types:
        evidence_pts += 3

    headline_score = max(15, min(98, savings_pts + comp_pts + complexity_pts + evidence_pts))

    if headline_score >= 88:
        severity_tier = "CRITICAL"
        tier_prefix = "High-Volume Manual Bottleneck"
    elif headline_score >= 70:
        severity_tier = "HIGH"
        tier_prefix = "Strong Automation Candidate"
    elif headline_score >= 45:
        severity_tier = "MEDIUM"
        tier_prefix = "Moderate Automation Candidate"
    else:
        severity_tier = "LOW"
        tier_prefix = "Light Reporting Cleanup"

    has_pdf_fmt = "pdf_invoices_or_contracts" in source_formats
    if headline_score >= 82 or exports_per_run >= 6 or (len(source_formats) >= 3 and has_pdf_fmt):
        estimated_sprint_days = "5-7 Weeks (Phased) · Phase 1: 1-Week Discovery · Phase 2: 4-6 Weeks (~160-240 hrs at 40 hrs/wk)"
    elif headline_score >= 62 or exports_per_run >= 4 or has_pdf_fmt:
        estimated_sprint_days = "3-5 Weeks · ~120-200 hrs (at 40 hrs/wk)"
    elif headline_score >= 40:
        estimated_sprint_days = "2-3 Weeks · ~80-120 hrs (at 40 hrs/wk)"
    else:
        estimated_sprint_days = "1-2 Weeks · ~40-80 hrs (at 40 hrs/wk)"
    payback_label = estimated_sprint_days

    score_label = f"{tier_prefix} ({headline_score}/100 Checklist Score · Current Time Cost: ${current_annual_cost:,}/yr)"

    # Build strictly client-safe, engineering-grade Top Findings (RC-01 .. RC-06)
    findings: List[Dict[str, str]] = []
    copy_paste_minutes = sum(s["manual_minutes"] for s in steps if s["is_copy_paste"])

    if (
        "manual_copy_paste_joins" in bottleneck_types
        or copy_paste_minutes > 0
        or "csv_excel_exports" in source_formats
        or xlsx_lookups > 0
    ):
        lookup_note = (
            f" Detected {xlsx_lookups} spreadsheet lookup/conditional formula(s) (`VLOOKUP`/`XLOOKUP`/`SUMIFS`) in workbook inspection."
            if xlsx_lookups > 0
            else ""
        )
        step_note = (
            f" ({round(copy_paste_minutes / 60.0, 1)} hrs/run spent directly in copy-paste steps)"
            if copy_paste_minutes > 0
            else ""
        )
        findings.append(
            {
                "id": f"RC-0{len(findings) + 1}",
                "severity": "HIGH",
                "title": "Manual Spreadsheet Copy-Pasting & VLOOKUP Bottleneck",
                "detail": (
                    f"Based on your inputs, your team spends {hours_per_run:g} hrs/run{step_note} assembling data across recurring exports "
                    f"({current_annual_hours:g} hrs/yr, or ${current_annual_cost:,}/yr at ${hourly_rate_usd:g}/hr).{lookup_note} "
                    f"At a {compression_pct}% what-if reduction target, that represents ~{annual_hours_reclaimed} hrs/yr (~${annual_savings_usd:,}/yr) of manual prep."
                ),
                "client_next_step": (
                    "Lock standard export column names up front and replace spreadsheet VLOOKUP chains with an automated join step that alerts you immediately if columns are missing or IDs are duplicated."
                ),
            }
        )

    if recon["mismatch_count"] > 0 or "reconciliation_mismatches" in bottleneck_types or local_orphans > 0:
        if recon["rows_checked"] > 0:
            recon_detail = (
                f"Your uploaded sample CSV showed {recon['mismatch_count']} mismatch(es) across {recon['rows_checked']} rows "
                f"(${recon['total_abs_variance_usd']:,.2f} total difference)."
            )
        elif local_orphans > 0:
            recon_detail = (
                f"Your local folder scan found {local_orphans} unmatched ID(s) and {local_dups} duplicate ID(s) "
                f"across export files, requiring manual row-by-row hunting."
            )
        else:
            recon_detail = (
                "Mismatched IDs, timing cutoffs, and missing credits between systems require manual line-by-line hunting every cycle."
            )
        findings.append(
            {
                "id": f"RC-0{len(findings) + 1}",
                "severity": "HIGH" if (recon["mismatch_count"] >= 2 or local_orphans >= 2) else "MEDIUM",
                "title": "Numbers & IDs Don't Match Cleanly Between Systems",
                "detail": recon_detail,
                "client_next_step": (
                    "Set up an automated System A vs. System B comparison step that groups exact matches, rounding differences, missing rows, and true variances before totals are finalized."
                ),
            }
        )

    if "pdf_invoices_or_contracts" in source_formats or any("pdf" in s["source_type"] for s in steps):
        pdf_pages = int((local_scan_summary or {}).get("totals", {}).get("total_pdf_pages_est", 0))
        pdf_note = f" (~{pdf_pages} PDF pages counted in your folder scan)" if pdf_pages > 0 else ""
        findings.append(
            {
                "id": f"RC-0{len(findings) + 1}",
                "severity": "MEDIUM",
                "title": "Manual Copying From PDF Invoices, Contracts, or Statements",
                "detail": (
                    f"Important totals, rebates, or billing terms sit inside PDFs{pdf_note}, "
                    "leading to re-typing mistakes or missed adjustments in Excel trackers."
                ),
                "client_next_step": (
                    "Extract PDF line items into structured fields with page references and an automatic math check that confirms line items add up to the invoice total."
                ),
            }
        )

    if "cloud_or_saas_billing" in source_formats:
        findings.append(
            {
                "id": f"RC-0{len(findings) + 1}",
                "severity": "MEDIUM",
                "title": "Inconsistent Cloud or Software Billing Exports",
                "detail": (
                    "Different cloud and software billing exports use different account, department, or cost-center labels, causing repetitive cleanup work each cycle."
                ),
                "client_next_step": (
                    "Map vendor billing exports into one consistent table with clear fallback rules for untagged charges."
                ),
            }
        )

    if "manual_summary_drafting" in bottleneck_types:
        findings.append(
            {
                "id": f"RC-0{len(findings) + 1}",
                "severity": "MEDIUM",
                "title": "Re-Writing Recurring Summary Notes by Hand",
                "detail": (
                    "Staff rewrite recurring weekly or monthly commentary from scratch, and pasting raw spreadsheets into generic AI chat tools risks wrong numbers."
                ),
                "client_next_step": (
                    "Draft summary notes strictly from pre-calculated totals, paired with an automatic check that confirms every number in the text matches the table."
                ),
            }
        )

    if "no_signoff_audit_trail" in bottleneck_types or recon["mismatch_count"] > 0:
        findings.append(
            {
                "id": f"RC-0{len(findings) + 1}",
                "severity": "HIGH",
                "title": "No Clear Record of Who Checked & Approved Adjustments",
                "detail": (
                    "Last-minute spreadsheet edits and Slack approvals leave no clear log of who verified each mismatch before reports go out."
                ),
                "client_next_step": (
                    "Keep every calculated total and adjustment in a 'Needs Human Sign-Off' step until a reviewer approves or notes a correction."
                ),
            }
        )

    schema_warn_count = int((local_scan_summary or {}).get("totals", {}).get("schema_drift_warnings", 0))
    if (
        "schema_drift_or_dirty_exports" in bottleneck_types
        or recon["missing_key_count"] > 0
        or recon["duplicate_key_count"] > 0
        or schema_warn_count > 0
    ):
        findings.append(
            {
                "id": f"RC-0{len(findings) + 1}",
                "severity": "MEDIUM",
                "title": "Shifting Column Names, Blank IDs, or Duplicate Rows",
                "detail": (
                    f"Found {schema_warn_count} column-mismatch warning(s), "
                    f"{recon['missing_key_count']} blank ID(s), and {recon['duplicate_key_count'] + local_dups} duplicate ID(s)."
                ),
                "client_next_step": (
                    "Check required column names, non-blank unique IDs, and date formats automatically before running report calculations."
                ),
            }
        )

    if not findings:
        findings.append(
            {
                "id": "RC-01",
                "severity": "MEDIUM",
                "title": "Recurring Report Time-Saving Opportunity",
                "detail": (
                    f"Based on your inputs, this report takes {current_annual_hours:g} hrs/yr (${current_annual_cost:,}/yr); "
                    f"at your {compression_pct}% what-if target, automating imports saves ~{annual_hours_reclaimed} hrs/yr."
                ),
                "client_next_step": "Standardize input CSV columns and automate recurring roll-ups behind a simple human sign-off step.",
            }
        )

    # Pipeline Verification Checklist (`safe_to_deploy_checklist`)
    checklist = [
        {
            "step": 1,
            "category": "Consistent Input Columns",
            "item": "Agree on standard column names, unique record IDs, and date/currency formats across all recurring export files.",
            "status": "ACTION_REQUIRED",
        },
        {
            "step": 2,
            "category": "Automatic File Matching",
            "item": "Replace manual VLOOKUPs and copy-pasting with an automated System A vs. System B match that flags missing or mismatched rows.",
            "status": "ACTION_REQUIRED"
            if ("manual_copy_paste_joins" in bottleneck_types or "reconciliation_mismatches" in bottleneck_types or recon["mismatch_count"] > 0)
            else "RECOMMENDED",
        },
        {
            "step": 3,
            "category": "PDF & Summary Fact-Check",
            "item": "Verify any numbers pulled from PDFs or included in draft summary notes against the underlying table totals.",
            "status": "ACTION_REQUIRED"
            if ("pdf_invoices_or_contracts" in source_formats or "manual_summary_drafting" in bottleneck_types)
            else "RECOMMENDED",
        },
        {
            "step": 4,
            "category": "Human Sign-Off Step",
            "item": "Keep every final total and mismatch adjustment in a 'Needs Human Sign-Off' state until a reviewer approves it.",
            "status": "ACTION_REQUIRED",
        },
        {
            "step": 5,
            "category": "Side-by-Side Comparison",
            "item": "Run the automated workflow alongside 1–2 past reporting cycles to verify numbers match before retiring manual spreadsheets.",
            "status": "ACTION_REQUIRED",
        },
    ]

    # Emit only registered diagnostic signal keys
    raw_signals: List[str] = []
    if (
        "csv_excel_exports" in source_formats
        or "manual_copy_paste_joins" in bottleneck_types
        or "reconciliation_mismatches" in bottleneck_types
        or recon["mismatch_count"] > 0
        or any(s["is_copy_paste"] for s in steps)
    ):
        raw_signals.append("manual_csv_reconciliation_bottleneck")

    if "pdf_invoices_or_contracts" in source_formats or any("pdf" in s["source_type"] for s in steps):
        raw_signals.append("unstructured_pdf_or_doc_intake")

    if "cloud_or_saas_billing" in source_formats or "multi_cloud_telemetry" in source_formats:
        raw_signals.append("multi_cloud_config_or_telemetry_sprawl")

    if "manual_summary_drafting" in bottleneck_types:
        raw_signals.append("hallucination_or_ungrounded_output")

    if (
        "no_signoff_audit_trail" in bottleneck_types
        or recon["mismatch_count"] > 0
        or any(s["error_prone"] for s in steps)
    ):
        raw_signals.append("missing_observability_or_hitl")

    if (
        "schema_drift_or_dirty_exports" in bottleneck_types
        or recon["missing_key_count"] > 0
        or recon["duplicate_key_count"] > 0
        or schema_warn_count > 0
    ):
        raw_signals.append("broken_rag_schema_or_data_hygiene")

    if "no_signoff_audit_trail" in bottleneck_types and (
        "compliance_audit_exposure" in bottleneck_types
        or prospect["industry"].lower() in {"healthcare", "banking", "insurance", "regulated"}
    ):
        raw_signals.append("audit_logging_or_retention_gap")

    if not raw_signals:
        raw_signals.append("manual_csv_reconciliation_bottleneck")

    routing_signals = [sig for sig in dict.fromkeys(raw_signals) if sig in SIGNAL_ROUTING_RULES]

    inputs_out: Dict[str, Any] = {
        "workflow_name": workflow_name,
        "runs_per_month": runs_per_month,
        "people_involved": people_involved,
        "exports_per_run": exports_per_run,
        "hours_per_run": hours_per_run,
        "hourly_rate_usd": hourly_rate_usd,
        "target_reduction_pct": compression_pct,
        "source_formats": source_formats,
        "bottleneck_types": bottleneck_types,
        "workflow_steps_csv_text": workflow_steps_csv_text,
        "reconciliation_csv_text": reconciliation_csv_text,
    }
    if xlsx_summary:
        inputs_out["xlsx_summary"] = xlsx_summary
    if local_scan_summary:
        inputs_out["local_scan_summary"] = local_scan_summary

    return {
        "capsule_version": CAPSULE_VERSION,
        "niche_id": NICHE_ID,
        "niche_slug": NICHE_SLUG,
        "niche_title": NICHE_TITLE,
        "mode": mode_clean,
        "generated_at": _iso_now(),
        "prospect": prospect,
        "inputs": inputs_out,
        "client_assessment": {
            "headline_score": headline_score,
            "score_label": score_label,
            "severity_tier": severity_tier,
            "estimated_sprint_days": estimated_sprint_days,
            "key_metrics": [
                {"label": "Current Annual Time", "value": f"{current_annual_hours} hrs/yr"},
                {"label": "Current Annual Time Cost", "value": f"${current_annual_cost:,} / yr"},
                {"label": f"What-If Hours Saved ({compression_pct}%)", "value": f"~{annual_hours_reclaimed} hrs/yr"},
                {"label": f"What-If Cost Saved ({compression_pct}%)", "value": f"~${annual_savings_usd:,} / yr"},
                {"label": "Estimated Effort", "value": estimated_sprint_days},
            ],
            "roi_breakdown": {
                "runs_per_year": round(runs_per_year, 1),
                "current_annual_hours": current_annual_hours,
                "current_annual_cost_usd": current_annual_cost,
                "compression_pct": compression_pct,
                "annual_hours_reclaimed": annual_hours_reclaimed,
                "post_sprint_hours_per_run": post_sprint_hours_per_run,
                "annual_savings_usd": annual_savings_usd,
                "payback_period": payback_label,
                "reconciliation_summary": {
                    "rows_checked": recon["rows_checked"],
                    "matched_count": recon["matched_count"],
                    "mismatch_count": recon["mismatch_count"],
                    "missing_key_count": recon["missing_key_count"],
                    "duplicate_key_count": recon["duplicate_key_count"],
                    "total_abs_variance_usd": recon["total_abs_variance_usd"],
                    "discrepancies": recon["discrepancies"],
                },
                "step_breakdown": steps,
            },
            "top_findings": findings[:6],
            "safe_to_deploy_checklist": checklist,
        },
        "routing_signals": routing_signals,
    }


def sanitize_capsule_for_share(capsule: Dict[str, Any]) -> Dict[str, Any]:
    """Produce a strictly structural capsule with zero financial cell values, customer IDs, or private paths."""
    cloned = json.loads(json.dumps(capsule))
    inp = cloned.get("inputs", {})
    inp["reconciliation_csv_text"] = ""
    if inp.get("workflow_steps_csv_text"):
        steps = parse_workflow_steps_csv(inp["workflow_steps_csv_text"])
        sanitized_lines = ["step_order,step_name,manual_minutes,is_copy_paste,error_prone,source_type"]
        for s in steps:
            sanitized_lines.append(
                f"{s['step_order']},Step_{s['step_order']}_{s['source_type']},{s['manual_minutes']},{str(s['is_copy_paste']).lower()},{str(s['error_prone']).lower()},{s['source_type']}"
            )
        inp["workflow_steps_csv_text"] = "\n".join(sanitized_lines)

    if isinstance(inp.get("local_scan_summary"), dict):
        scan = inp["local_scan_summary"]
        scan["scanned_path"] = "<REDACTED_LOCAL_PATH>"
        for idx, f_entry in enumerate(scan.get("files", []), start=1):
            f_entry["file"] = f"source_{idx}.{str(f_entry.get('type', 'dat')).lower()}"

    roi = cloned.get("client_assessment", {}).get("roi_breakdown", {})
    if isinstance(roi.get("reconciliation_summary"), dict):
        roi["reconciliation_summary"]["discrepancies"] = []
        roi["reconciliation_summary"]["total_abs_variance_usd"] = 0.0

    cloned["sanitized_structural_only"] = True
    return cloned


def render_markdown_report(capsule: Dict[str, Any]) -> str:
    """Render a printable Markdown ROI & Deterministic Compression Pipeline Specification."""
    ca = capsule.get("client_assessment", {})
    inp = capsule.get("inputs", {})
    prospect = capsule.get("prospect", {})
    roi = ca.get("roi_breakdown", {})
    recon = roi.get("reconciliation_summary", {})
    steps = roi.get("step_breakdown", [])

    org_line = prospect.get("company_name") or "Local Workspace Assessment"
    lines: List[str] = [
        f"# {inp.get('workflow_name', 'Reporting Workflow')}, Compression & Pipeline Spec",
        "",
        f"- **Organization / Workspace:** {org_line}",
        f"- **Generated At:** `{capsule.get('generated_at', '')}` (`mode={capsule.get('mode', 'production')}`)",
        f"- **Headline Score:** **{ca.get('score_label', '')}**",
        f"- **Estimated Implementation Window:** `{ca.get('estimated_sprint_days', 'Typically 2-3 Weeks')}`",
        "",
        "## 1. Executive Hours-Saved & ROI Summary",
        "",
        "| Metric | Baseline / What-If Projection |",
        "| :--- | :--- |",
        f"| Current Effort per Run (Your Input) | `{inp.get('hours_per_run', 0)} hrs/run` ({inp.get('runs_per_month', 0)} runs/mo) |",
        f"| Current Annual Time & Cost (Exact Math) | `{roi.get('current_annual_hours', 0)} hrs/yr` (`${roi.get('current_annual_cost_usd', 0):,}/yr` @ `${inp.get('hourly_rate_usd', 0)}/hr`) |",
        f"| Selected What-If Reduction Target | **`{roi.get('compression_pct', 50)}%`** (illustrative target: `~{roi.get('post_sprint_hours_per_run', 0)} hrs/run`) |",
        f"| **What-If Annual Hours Saved** | **`~{roi.get('annual_hours_reclaimed', 0)} hrs/yr`** |",
        f"| **What-If Annual Cost Saved** | **`~${roi.get('annual_savings_usd', 0):,} / yr`** |",
        "",
    ]

    if steps:
        lines.extend(
            [
                "## 2. Step-by-Step Workflow Compression Blueprint",
                "",
                "| # | Workflow Step | Current (min) | Post-Sprint (min) | Target Automation Pattern |",
                "| :--- | :--- | ---: | ---: | :--- |",
            ]
        )
        for s in steps:
            lines.append(
                f"| {s['step_order']} | {s['step_name']} | {s['manual_minutes']} | {s['post_sprint_minutes']} | {s['automation_pattern']} |"
            )
        lines.append("")

    if recon.get("rows_checked", 0) > 0:
        lines.extend(
            [
                "## 3. Sample Reconciliation Drift Audit",
                "",
                f"- **Records Checked:** `{recon['rows_checked']}` | **Matched:** `{recon['matched_count']}` | **Discrepancies:** `{recon['mismatch_count']}`",
                f"- **Total Unreconciled Absolute Variance:** `${recon.get('total_abs_variance_usd', 0.0):,.2f}`",
                "",
            ]
        )
        if recon.get("discrepancies"):
            lines.extend(
                [
                    "| Record Key | Issue Type | Source A | Source B | Delta ($) | Note |",
                    "| :--- | :--- | ---: | ---: | ---: | :--- |",
                ]
            )
            for d in recon["discrepancies"][:10]:
                sa = f"${d['source_a_amount']:,.2f}" if d["source_a_amount"] is not None else "*(missing)*"
                sb = f"${d['source_b_amount']:,.2f}" if d["source_b_amount"] is not None else "*(missing)*"
                lines.append(
                    f"| `{d['record_key']}` | `{d['issue']}` | {sa} | {sb} | `${d['delta_usd']:,.2f}` | {d['status_note']} |"
                )
            lines.append("")

    local_scan = inp.get("local_scan_summary")
    if isinstance(local_scan, dict) and local_scan.get("totals"):
        tot = local_scan["totals"]
        lines.extend(
            [
                "## 4. Local Folder Structural Profile (Zero Raw Values)",
                "",
                f"- **Files Profiled:** `{tot.get('total_files', 0)}` (`{tot.get('csv_tsv_files', 0)}` CSV/TSV, `{tot.get('xlsx_files', 0)}` XLSX, `{tot.get('json_files', 0)}` JSON, `{tot.get('pdf_files', 0)}` PDF)",
                f"- **Tabular Rows / PDF Pages:** `{tot.get('total_tabular_rows', 0):,}` rows | `~{tot.get('total_pdf_pages_est', 0)}` PDF pages",
                f"- **Spreadsheet Formulas:** `{tot.get('total_excel_formulas', 0)}` total (`{tot.get('total_lookup_formulas', 0)}` brittle lookup/conditional formulas)",
                f"- **Key Hygiene:** `{tot.get('total_duplicate_keys', 0)}` duplicate keys | `{tot.get('join_orphan_keys_detected', 0)}` cross-file orphan keys | `{tot.get('schema_drift_warnings', 0)}` schema warnings",
                "",
            ]
        )

    lines.extend(["## Prioritized Bottleneck Findings", ""])
    for f in ca.get("top_findings", []):
        lines.extend(
            [
                f"### [{f['id']}] ({f['severity']}) {f['title']}",
                f"- **Observation:** {f['detail']}",
                f"- **Recommended Engineering Fix:** {f['client_next_step']}",
                "",
            ]
        )

    lines.extend(
        [
            "## Pipeline Verification & `'Not checked'` Governance Checklist",
            "",
            "| Step | Category | Verification Gate | Status |",
            "| :--- | :--- | :--- | :--- |",
        ]
    )
    for item in ca.get("safe_to_deploy_checklist", []):
        lines.append(f"| {item['step']} | **{item['category']}** | {item['item']} | `{item['status']}` |")
    lines.append("")

    lines.extend(
        [
            "---",
            "> **Next Steps (100% Optional):** Automating this internally? Use the plan above freely with your team. Want VARINULL to turn this into a clean, repeatable workflow (typically 2–3 weeks, scoped upfront)? Export a Summary File (zero financial numbers or customer rows) whenever you want a hand.",
            "",
        ]
    )

    return "\n".join(lines)
