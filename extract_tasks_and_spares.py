import argparse
import os
import re
from collections import OrderedDict
from copy import copy
from pathlib import Path

import pytesseract
from pdf2image import convert_from_path
from openpyxl import Workbook
from PIL import ImageOps

# More tolerant pattern for part-number-like strings (handles "1550595-9960" etc.)
PART_CANDIDATE_RE = re.compile(r"\d{4,}[- ]?\d{3,4}")


def _normalize_windows_path(path: str | None) -> str | None:
    """Convert MSYS/Cygwin style /c/ paths to Windows style C:/ for poppler."""
    if not path:
        return path
    if path.startswith("/"):
        parts = path.lstrip("/").split("/", 1)
        if len(parts) == 2 and len(parts[0]) == 1:
            drive, rest = parts
            return f"{drive.upper()}:/{rest}"
    return path


def configure_ocr_from_env() -> str | None:
    """
    Configure pytesseract and Poppler from environment variables.

    TESSERACT_CMD / TESSERACT_CMD_WIN: optional, full path or command name
    POPPLER_PATH / POPPLER_PATH_WIN : optional, path to Poppler 'bin' directory

    Returns POPPLER_PATH (or None if not set).
    """
    env_file = Path(".env.ocr")
    if env_file.exists():
        for line in env_file.read_text().splitlines():
            if "=" in line:
                k, v = line.split("=", 1)
                if k not in os.environ:
                    os.environ[k] = v

    tess_cmd = os.environ.get("TESSERACT_CMD") or os.environ.get("TESSERACT_CMD_WIN")
    if tess_cmd:
        pytesseract.pytesseract.tesseract_cmd = tess_cmd

    poppler_path = os.environ.get("POPPLER_PATH") or os.environ.get("POPPLER_PATH_WIN")
    poppler_path = _normalize_windows_path(poppler_path) if poppler_path else None
    return poppler_path


def pdf_to_text_lines(pdf_path: Path, poppler_path: str | None):
    """
    Convert the entire PDF to text lines using OCR.

    We process all pages; task and spare-part detection is done purely
    from content (no page ranges needed).
    """
    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF not found: {pdf_path}")

    images = convert_from_path(
        str(pdf_path),
        dpi=300,
        poppler_path=poppler_path,
    )

    all_lines: list[str] = []
    for i, img in enumerate(images, start=1):
        print(f"OCR on page {i}...")
        gray = img.convert("L")
        enhanced = ImageOps.autocontrast(gray, cutoff=2)
        processed = enhanced.point(lambda x: 0 if x < 180 else 255, "1")

        # Multiple OCR passes to better capture shaded rows and long lines
        ocr_inputs = [
            ("--oem 3 --psm 6", img),
            ("--oem 3 --psm 4", processed),
            ("--oem 3 --psm 6", processed),
        ]

        combined: list[str] = []
        for cfg, source_img in ocr_inputs:
            text = pytesseract.image_to_string(source_img, config=cfg)
            combined.extend([ln.rstrip() for ln in text.splitlines() if ln.strip()])

        # Deduplicate while preserving order (normalize spaces for matching)
        seen = OrderedDict()
        for ln in combined:
            key = " ".join(ln.split())
            if key not in seen:
                seen[key] = ln

        page_lines = list(seen.values())
        all_lines.extend(page_lines)

    # dump OCR text for debugging
    dump_path = pdf_path.with_suffix(".ocr_all_pages.txt")
    dump_path.write_text("\n".join(all_lines), encoding="utf-8")
    print(f"OCR dump written to: {dump_path}")

    return all_lines


def is_header_line(line: str) -> bool:
    """Detect the header row (very tolerant)."""
    lower = " ".join(line.lower().split())
    return (
        "task code" in lower
        and "trade" in lower
        and "task action" in lower
        and "task description" in lower
    )


def parse_grey_row(line: str):
    """
    From a grey row line, extract:
    - Location 1
    - Location 2
    - setTypeCode (number inside parentheses)
    - ComponentPath (full line, unchanged)
    """
    component_path = line.strip()

    if "\\" in line:
        parts = line.split("\\", 1)
    else:
        parts = [line, ""]

    location1 = parts[0].strip()
    location2 = parts[1].strip()

    # setTypeCode: digits inside parentheses (prefer last occurrence)
    all_codes = re.findall(r"\((\d+)\)", line)
    set_type_code = all_codes[-1] if all_codes else ""

    return location1, location2, set_type_code, component_path


def looks_like_component_line(line: str) -> bool:
    """
    Heuristic for grey/component rows that carry location info.
    They usually contain backslashes and/or bracketed part numbers.
    """
    stripped = line.strip()
    if not stripped:
        return False

    if "\\" in line or "[" in line or stripped.startswith("("):
        return True

    return False


def strip_status_prefix(line: str) -> str:
    """
    Remove OCR bullet/status noise before the task code (e.g. 'oO ', 'OO ', 'O ', 'i ').
    Returns the substring starting at the first digit or '*'.
    """
    match = re.search(r"[\*\d]", line)
    if match:
        return line[match.start() :].strip()
    return line.strip()


def is_metadata_line(line: str) -> bool:
    """Filter out page/asset/database footers and headers."""
    lower = line.lower()
    return lower.startswith(("asset:", "database:", "printed by", "page "))


def looks_like_task_row(line: str) -> bool:
    """
    Heuristic to identify task rows, even with noisy prefixes.
    """
    if is_metadata_line(line):
        return False

    # component/grey lines are handled separately
    if "\\" in line and "(" in line and ")" in line:
        return False

    normalized = strip_status_prefix(line)
    tokens = normalized.split()
    if len(tokens) < 3:
        return False

    code_token = tokens[0]
    if "/" in code_token:
        return False

    has_code = bool(re.match(r"\*?\d{6,8}$", code_token))
    trade_token = tokens[1]
    trade_ok = trade_token.isalpha() and trade_token.isupper()
    return has_code and trade_ok


def normalize_task_code(code: str) -> str:
    """Strip leading '*' or non-digit noise for consistent lookup keys."""
    return re.sub(r"^\*", "", code)


def pick_task_code_from_tokens(tokens: list[str]) -> str:
    """Find the first token that looks like a task code."""
    for tok in tokens:
        if re.fullmatch(r"\*?\d{6,8}", tok):
            return normalize_task_code(tok)
    return ""


def _split_interval(tokens: list[str]) -> tuple[list[str], str]:
    """Split tokens into (body, interval) using common interval endings."""
    if not tokens:
        return [], ""

    lower_tail = [t.lower() for t in tokens]

    if len(tokens) >= 2 and lower_tail[-2:] == ["no", "interval"]:
        return tokens[:-2], "No Interval"

    if lower_tail[-1] in {
        "hours",
        "hour",
        "week",
        "weeks",
        "month",
        "months",
        "day",
        "days",
    }:
        if len(tokens) >= 2 and re.match(r"^\d+$", tokens[-2]):
            return tokens[:-2], f"{tokens[-2]} {tokens[-1]}"

    if lower_tail[-1] == "recap":
        return tokens[:-1], "Recap"

    if len(tokens) >= 2:
        return tokens[:-2], " ".join(tokens[-2:])

    return [], tokens[-1]


def _split_doc_ref_and_description(tokens: list[str]) -> tuple[str, str]:
    """
    Split remaining tokens into (description, doc_ref).
    Doc ref is assumed to be the trailing block containing short uppercase words,
    digits or punctuation.
    """
    if not tokens:
        return "", ""

    if len(tokens) >= 2 and [tok.lower() for tok in tokens[-2:]] == ["no", "reference"]:
        return " ".join(tokens[:-2]).strip(), "No reference"

    lookback = tokens[-5:] if len(tokens) > 5 else tokens
    lookback_start = len(tokens) - len(lookback)
    doc_start = len(tokens) - 1
    for offset, tok in enumerate(lookback):
        if re.search(r"\d", tok) or re.search(r"[:;/.\-]", tok) or (
            tok.isupper() and len(tok) <= 4
        ):
            doc_start = lookback_start + offset
            break

    desc_tokens = tokens[:doc_start]
    doc_tokens = tokens[doc_start:]

    description = " ".join(desc_tokens).strip()
    doc_ref = " ".join(doc_tokens).strip()
    return description, doc_ref


def parse_task_row(line: str):
    """
    Parse a task row into:
    Task Code, Trade, Task Action, Task Description, Doc Ref, Interval
    """
    normalized = strip_status_prefix(line)
    tokens = normalized.split()
    if len(tokens) < 3:
        return "", "", "", "", "", ""

    task_code = tokens[0]
    trade = tokens[1] if len(tokens) > 1 else ""
    task_action = tokens[2] if len(tokens) > 2 else ""

    remaining = tokens[3:]
    body_tokens, interval = _split_interval(remaining)
    task_description, doc_ref = _split_doc_ref_and_description(body_tokens)

    return task_code, trade, task_action, task_description, doc_ref, interval


def build_workbook(task_rows: list[dict], spare_rows: list[dict]) -> Workbook:
    """
    Create workbook with Tasks and SpareParts sheets.
    """
    wb = Workbook()

    # --- Tasks sheet ---
    ws_tasks = wb.active
    ws_tasks.title = "Tasks"
    ws_tasks.delete_rows(1, ws_tasks.max_row)

    task_headers = [
        "Sort",
        "TaskCode",
        "TaskAction",
        "TaskDescription",
        "TypeOfWork",
        "MotionType",
        "Duration",
        "DurationCalc",
        "DurationUOM",
        "Interval",
        "MTBMPredicted",
        "CostCode",
        "IncludeInME",
        "TaskDependency",
        "FollowUpTasks",
        "LocationDependency",
        "Active",
        "Trade",
        "Section",
        "DocRef",
        "Location1",
        "Location2",
        "ComponentPath",
        "AssetType",
        "AssetTypeCode",
    ]

    ws_tasks.append(task_headers)
    for cell in ws_tasks[1]:
        bold_font = copy(cell.font)
        bold_font.bold = True
        cell.font = bold_font
    ws_tasks.row_dimensions[1].height = 22

    for idx, row in enumerate(task_rows, start=1):
        row["Sort"] = idx
        ws_tasks.append([row.get(h, "") for h in task_headers])

    # --- SpareParts sheet ---
    ws_spares = wb.create_sheet("SpareParts")
    spare_headers = [
        "TaskCode",
        "PartNo",
        "PartDescription",
        "MU_TL",
        "QtyRequired",
        "UOM",
        "ItemDependency",
        "Location1",
        "Location2",
        "AssetType",
        "AssetTypeCode",
    ]
    ws_spares.append(spare_headers)
    for cell in ws_spares[1]:
        bold_font = copy(cell.font)
        bold_font.bold = True
        cell.font = bold_font
    ws_spares.row_dimensions[1].height = 22

    for row in spare_rows:
        ws_spares.append([row.get(h, "") for h in spare_headers])

    return wb


def extract_tasks_and_assets_from_lines(lines: list[str]):
    """
    Scan all OCR lines and extract task rows + asset context.
    """
    print("Parsing task lines across entire document...")
    rows: list[dict] = []
    rows_by_code: dict[str, dict] = {}
    current_location1 = ""
    current_location2 = ""
    current_set_type_code = ""
    current_component_path = ""
    last_task_code = ""
    asset_type = ""
    asset_code = ""

    for line in lines:
        if not line.strip():
            continue

        if line.lower().startswith("asset:"):
            parts = line.split(":", 1)[1].strip().split()
            if parts:
                asset_code = parts[0]
                asset_type = " ".join(parts[1:]) if len(parts) > 1 else ""
            continue

        if is_header_line(line):
            print(f"Found header line: {line}")
            continue

        if is_metadata_line(line):
            continue

        if looks_like_component_line(line):
            (
                current_location1,
                current_location2,
                current_set_type_code,
                current_component_path,
            ) = parse_grey_row(line)
            continue

        if looks_like_task_row(line):
            (
                task_code,
                trade,
                task_action,
                task_description,
                doc_ref,
                interval,
            ) = parse_task_row(line)

            norm_code = normalize_task_code(task_code)

            row = {
                "TaskCode": norm_code,
                "Trade": trade,
                "TaskAction": task_action,
                "TaskDescription": task_description,
                "DocRef": doc_ref,
                "Interval": interval,
                "Location1": current_location1,
                "Location2": current_location2,
                "setTypeCode": current_set_type_code,
                "ComponentPath": current_component_path,
                "TypeOfWork": "",
                "MotionType": "",
                "Duration": "",
                "DurationCalc": "",
                "DurationUOM": "",
                "MTBMPredicted": "",
                "CostCode": "",
                "IncludeInME": "",
                "TaskDependency": "",
                "FollowUpTasks": "",
                "LocationDependency": "",
                "Active": "Y",
                "Section": "",
                "AssetType": "",
                "AssetTypeCode": current_set_type_code,
            }
            if norm_code in rows_by_code:
                existing = rows_by_code[norm_code]
                for key, value in row.items():
                    if key == "TaskDescription":
                        if len(value) > len(existing.get(key, "")):
                            existing[key] = value
                        continue
                    if existing.get(key, "") == "" and value != "":
                        existing[key] = value
                rows_by_code[norm_code] = existing
            else:
                rows.append(row)
                rows_by_code[norm_code] = row

            last_task_code = norm_code
            continue

        # Continuation lines for description
        if last_task_code and not is_metadata_line(line):
            target = rows_by_code.get(last_task_code)
            if target is not None:
                target["TaskDescription"] = (
                    f"{target['TaskDescription']} {line.strip()}"
                ).strip()

    print(f"Parsed {len(rows)} task rows.")
    return rows, rows_by_code, asset_type, asset_code


def looks_like_part_line(line: str) -> bool:
    """
    Tolerant heuristic for part lines: look for something that resembles
    a long digit block with optional hyphen / space.
    """
    return bool(PART_CANDIDATE_RE.search(line))


def parse_part_block(lines: list[str], idx: int):
    """
    Parse a spare part record starting at lines[idx].
    Returns (record_dict, next_index).
    """
    line = lines[idx]
    tokens = line.split()
    if not tokens:
        return None, idx + 1

    # Locate part number (tolerant to OCR noise)
    part_idx = None
    part_no = ""

    for j, tok in enumerate(tokens):
        clean = re.sub(r"[^\d\-]", "", tok)
        if PART_CANDIDATE_RE.fullmatch(clean):
            part_idx = j
            part_no = clean
            break

    # handle split part numbers, e.g. "1550595" "9960"
    if part_idx is None and len(tokens) >= 2:
        for j in range(len(tokens) - 1):
            t1, t2 = tokens[j], tokens[j + 1]
            c1 = re.sub(r"\D", "", t1)
            c2 = re.sub(r"\D", "", t2)
            if c1.isdigit() and len(c1) >= 4 and c2.isdigit() and len(c2) >= 3:
                part_idx = j
                part_no = f"{c1}-{c2}"
                break

    if part_idx is None:
        return None, idx + 1

    # Task code and action
    task_idx = None
    for j in range(part_idx + 1, len(tokens)):
        if re.fullmatch(r"\*?\d{6,8}", tokens[j]):
            task_idx = j
            break
    task_code = normalize_task_code(tokens[task_idx]) if task_idx is not None else ""
    task_action = (
        tokens[task_idx + 1] if task_idx is not None and task_idx + 1 < len(tokens) else ""
    )

    # Description between part and task code
    desc_tokens = tokens[part_idx + 1 : task_idx if task_idx else len(tokens)]
    # Component / qty / uom after task_action
    comp_tokens = tokens[(task_idx + 2) if task_idx is not None else len(tokens) :]

    # Detect qty/uom at end
    uom = ""
    qty = ""
    if comp_tokens:
        last = comp_tokens[-1]
        if re.fullmatch(r"[A-Za-z]{1,4}", last):
            uom = last
            comp_tokens = comp_tokens[:-1]
    if comp_tokens:
        last = comp_tokens[-1]
        if re.fullmatch(r"[0-9]+(?:\.[0-9]+)?", last):
            qty = last
            comp_tokens = comp_tokens[:-1]

    component_path = " ".join(comp_tokens).strip()
    # Pull in continuation lines that don't start a new part or task header
    next_idx = idx + 1
    while next_idx < len(lines):
        nxt = lines[next_idx]
        if (
            looks_like_part_line(nxt)
            or looks_like_task_row(nxt)
            or is_header_line(nxt)
            or is_metadata_line(nxt)
        ):
            break
        if nxt.strip():
            component_path = (component_path + " " + nxt.strip()).strip()
        next_idx += 1

    part_description = " ".join(desc_tokens).strip()
    return (
        {
            "TaskCode": task_code,
            "TaskAction": task_action,
            "PartNo": part_no,
            "PartDescription": part_description,
            "QtyRequired": qty,
            "UOM": uom,
            "ComponentPath": component_path,
        },
        next_idx,
    )


def extract_spare_parts_from_lines(
    lines: list[str],
    task_lookup: dict[str, dict],
):
    """
    Scan all OCR lines and extract spare-part rows.

    Because we pass the complete task_lookup (built from the same full
    lines list), this works even if spare parts appear before the task
    rows in the PDF: the TaskCode references will still match tasks.
    """
    print("Parsing spare part lines across entire document...")
    spare_rows: list[dict] = []
    current_task_code = ""
    current_location1 = ""
    current_location2 = ""
    current_set_type_code = ""
    current_component_path = ""
    seen_keys: set[tuple] = set()

    idx = 0
    while idx < len(lines):
        line = lines[idx]
        idx += 1

        if not line.strip():
            continue

        lower = line.lower()
        if is_metadata_line(line) or lower.startswith("asset:"):
            continue

        # 1) Spare part lines FIRST so they aren't eaten by component heuristic
        if looks_like_part_line(line):
            parsed, next_idx = parse_part_block(lines, idx - 1)
            idx = next_idx
            if not parsed:
                continue

            task_code = parsed.get("TaskCode") or current_task_code or ""
            if not task_code:
                # If we don't have a task code at all, we can't map it properly.
                continue

            task_ctx = task_lookup.get(task_code, {})
            key = (task_code, parsed["PartNo"], parsed["PartDescription"])
            if key in seen_keys:
                continue
            seen_keys.add(key)

            # prefer current component context, otherwise task context
            loc1 = current_location1 or task_ctx.get("Location1", "")
            loc2 = current_location2 or task_ctx.get("Location2", "")
            set_type_code = current_set_type_code or task_ctx.get("setTypeCode", "")
            if not set_type_code:
                match_codes = re.findall(
                    r"\((\d+)\)", parsed.get("ComponentPath", "")
                )
                if match_codes:
                    set_type_code = match_codes[-1]

            spare_rows.append(
                {
                    "TaskCode": task_code,
                    "PartNo": parsed["PartNo"],
                    "PartDescription": parsed["PartDescription"],
                    "MU_TL": "",
                    "QtyRequired": parsed["QtyRequired"],
                    "UOM": parsed["UOM"],
                    "ItemDependency": "",
                    "Location1": loc1,
                    "Location2": loc2,
                    "AssetType": "",
                    "AssetTypeCode": set_type_code,
                }
            )
            continue  # handled this line

        # 2) capture component context for following parts
        if looks_like_component_line(line) and not looks_like_part_line(line):
            (
                current_location1,
                current_location2,
                current_set_type_code,
                current_component_path,
            ) = parse_grey_row(line)
            continue

        # 3) task line (context for subsequent parts)
        if looks_like_task_row(line):
            task_code, *_ = strip_status_prefix(line).split(maxsplit=1)
            current_task_code = normalize_task_code(task_code)
            continue

    print(f"Parsed {len(spare_rows)} spare part rows.")
    return spare_rows


def main():
    parser = argparse.ArgumentParser(
        description="Extract task table and spare parts from scanned PDF into Excel."
    )
    parser.add_argument(
        "--pdf",
        required=True,
        help="Path to input PDF file.",
    )
    parser.add_argument(
        "--out",
        help="Output XLSX path. Defaults to <pdf_stem>_tasks_spares.xlsx",
    )

    args = parser.parse_args()

    pdf_path = Path(args.pdf)

    if args.out:
        output_xlsx = Path(args.out)
    else:
        output_xlsx = pdf_path.with_name(f"{pdf_path.stem}_tasks_spares.xlsx")

    poppler_path = configure_ocr_from_env()
    print(f"Using POPPLER_PATH={poppler_path!r}")
    print("Reading and OCR'ing entire PDF...")

    # Single OCR pass over entire document
    lines = pdf_to_text_lines(pdf_path, poppler_path)

    # First pass: tasks (so task_lookup is complete)
    task_rows, task_lookup, asset_type, asset_code = extract_tasks_and_assets_from_lines(
        lines
    )

    # Second pass: spare parts (can safely reference tasks, even if parts appear before them)
    spare_rows = extract_spare_parts_from_lines(
        lines,
        task_lookup=task_lookup,
    )

    wb = build_workbook(task_rows, spare_rows)
    wb.save(output_xlsx)
    print(f"Saved Excel file: {output_xlsx}")


if __name__ == "__main__":
    main()
