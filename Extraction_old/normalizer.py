"""
Financial Statement Normalizer — Generic OpenAI Version
=========================================================
One script for ALL financial statement types.
The ONLY thing that changes per statement type is the --prompt file.

Usage:
    python normalizer.py \
        --pdf    path/to/statement.pdf \
        --xlsx   path/to/standard_coa_master.xlsx \
        --prompt path/to/SNP_Normalize_With_COA_Prompt.txt \
        --output output.json

Other statement types (just swap --prompt):
    python normalizer.py --prompt BS_Normalize_Prompt.txt   ...  # Balance Sheet
    python normalizer.py --prompt IS_Normalize_Prompt.txt   ...  # Income Statement
    python normalizer.py --prompt CF_Normalize_Prompt.txt   ...  # Cash Flow

Optional flags:
    --model          gpt-4o (default) | gpt-4o-mini | gpt-4-turbo
    --max-tokens     max output tokens (default: 8000)
    --temperature    0.0 to 1.0 (default: 0 for deterministic output)
    --reporting-columns "Col1" "Col2" ...   override column detection

Requirements:
    pip install openai openpyxl

Environment variable:
    export OPENAI_API_KEY=sk-...
"""

import argparse
import base64
import json
import os
import re
import sys
from pathlib import Path

import openpyxl
from openai import OpenAI, APIError, APITimeoutError, RateLimitError

# Matches the page-number tag appended by financial_statement_extractor.py
# Examples: _SNP_p15  _SOA_p16-17  _BS_p22
_PAGE_TAG_RE = re.compile(
    r"_(?:SNP|SOA|BS|IS|CF)_p(\d+)(?:-(\d+))?$",
    re.IGNORECASE,
)


# ─────────────────────────────────────────────────────────────────────────────
# 1.  ARGUMENT PARSING
# ─────────────────────────────────────────────────────────────────────────────

def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="Generic financial statement normalizer using OpenAI GPT-4o.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    p.add_argument("--pdf",    required=True,  help="Path to the financial statement PDF.")
    p.add_argument("--xlsx",   required=True,  help="Path to standard_coa_master.xlsx.")
    p.add_argument("--prompt", required=True,  help="Path to the statement-specific prompt .txt file.")
    p.add_argument("--output", default="normalized_output.json",
                   help="Output JSON file path (default: normalized_output.json).")
    p.add_argument("--model",  default="gpt-4o-mini",
                   help="OpenAI model (default: gpt-4o-mini).")
    p.add_argument("--max-tokens", type=int, default=8000,
                   help="Max completion tokens (default: 8000). Raise for very long statements.")
    p.add_argument("--temperature", type=float, default=0.0,
                   help="Sampling temperature (default: 0.0 for deterministic output).")
    p.add_argument("--reporting-columns", nargs="*", default=None, metavar="COL",
                   help=(
                       "Optional: lock exact column names to override PDF header detection. "
                       "Example: --reporting-columns 'Governmental Activities' 'Business-type Activities' 'Totals'"
                   ))
    return p.parse_args()


# ─────────────────────────────────────────────────────────────────────────────
# 2.  FILE LOADERS
# ─────────────────────────────────────────────────────────────────────────────

def load_text_file(path: str, label: str) -> str:
    p = Path(path)
    if not p.exists():
        sys.exit(f"[ERROR] {label} not found: {path}")
    content = p.read_text(encoding="utf-8")
    print(f"[INFO] {label:<22} {len(content):>8,} chars   {path}")
    return content


def load_pdf_as_data_uri(path: str) -> tuple[str, str]:
    p = Path(path)
    if not p.exists():
        sys.exit(f"[ERROR] PDF not found: {path}")
    raw = p.read_bytes()
    b64 = base64.b64encode(raw).decode("utf-8")
    print(f"[INFO] {'PDF':<22} {len(raw):>8,} bytes   {path}")
    return f"data:application/pdf;base64,{b64}", p.name


def load_xlsx_as_pipe_text(path: str) -> str:
    p = Path(path)
    if not p.exists():
        sys.exit(f"[ERROR] XLSX not found: {path}")

    wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
    ws = wb.worksheets[0]

    lines = []
    row_count = 0
    for row in ws.iter_rows(values_only=True):
        if all(c is None or str(c).strip() == "" for c in row):
            continue
        cells = [str(c).strip() if c is not None else "" for c in row]
        lines.append(" | ".join(cells))
        row_count += 1

    wb.close()
    print(f"[INFO] {'XLSX (sheet 1)':<22} {row_count:>8,} rows    {path}")
    return "\n".join(lines)


# ─────────────────────────────────────────────────────────────────────────────
# 3.  PAGE INFO EXTRACTION FROM FILENAME
# ─────────────────────────────────────────────────────────────────────────────

def extract_page_info_from_filename(pdf_path: str) -> dict | None:
    """
    Parse the page-number tag that financial_statement_extractor.py embeds in
    the filename, e.g.:

        CityReport_SNP_p15.pdf      → {"start": 15, "end": 15, "label": "page 15"}
        CityReport_SOA_p16-17.pdf   → {"start": 16, "end": 17, "label": "pages 16–17"}

    Returns None if no tag is found (graceful fallback for old-style filenames).
    """
    stem = Path(pdf_path).stem
    m = _PAGE_TAG_RE.search(stem)
    if not m:
        return None

    start = int(m.group(1))
    end   = int(m.group(2)) if m.group(2) else start

    if start == end:
        label = f"page {start}"
        pages = [start]
    else:
        label = f"pages {start}–{end}"
        pages = list(range(start, end + 1))

    return {"start": start, "end": end, "pages": pages, "label": label}


# ─────────────────────────────────────────────────────────────────────────────
# 4.  BUILD OPENAI MESSAGES
# ─────────────────────────────────────────────────────────────────────────────

def build_messages(
    system_prompt: str,
    pdf_data_uri: str,
    pdf_filename: str,
    coa_text: str,
    reporting_columns: list[str] | None,
    page_info: dict | None,
) -> list[dict]:
    """
    Construct the two-message list for OpenAI Chat Completions.

    Message 1 — system:
        The full statement-specific normalization prompt.

    Message 2 — user (multi-block content):
        Block A : PDF file (native file input, base64 data URI)
        Block B : COA master as plain-text document
        Block C : SOURCE PAGE REFERENCE (injected when page_info is available)
        Block D : Instruction + optional [[REPORTING COLUMNS]] override
    """

    # ── Block A: PDF ─────────────────────────────────────────────────────────
    block_pdf = {
        "type": "file",
        "file": {
            "filename": pdf_filename,
            "file_data": pdf_data_uri,
        },
    }

    # ── Block B: COA Master ──────────────────────────────────────────────────
    block_coa = {
        "type": "text",
        "text": (
            "STANDARD COA MASTER (pipe-delimited table):\n"
            "Format of each row: COA Flag | COA Datapoint | Statement | Section\n"
            "Use this table exactly as instructed in the system prompt for COA Datapoint mapping.\n\n"
            + coa_text
        ),
    }

    user_content = [block_pdf, block_coa]

    # ── Block C: SOURCE PAGE REFERENCE (NEW) ─────────────────────────────────
    # Tells the LLM which pages of the original full report this extracted PDF
    # corresponds to, so it can populate Metadata.Page No correctly (Section B
    # of the SNP prompt) and any per-row Source Page fields.
    if page_info:
        if page_info["start"] == page_info["end"]:
            page_note = (
                f"SOURCE PAGE REFERENCE:\n"
                f"This extracted PDF corresponds to {page_info['label']} of the "
                f"original full financial report (1-based page numbering).\n"
                f"When populating the 'Page No' field in Metadata, use: {page_info['start']}.\n"
                f"When populating any 'Source Page' or 'Page Reference' field, "
                f"use: {page_info['start']}."
            )
        else:
            pages_str = ", ".join(str(p) for p in page_info["pages"])
            page_note = (
                f"SOURCE PAGE REFERENCE:\n"
                f"This extracted PDF corresponds to {page_info['label']} of the "
                f"original full financial report (1-based page numbering).\n"
                f"The content spans pages: {pages_str}.\n"
                f"When populating the 'Page No' field in Metadata, use a "
                f"comma-separated list like \"{pages_str}\".\n"
                f"When populating any 'Source Page' or 'Page Reference' field, "
                f"use the page where each specific data point appears."
            )
        user_content.append({"type": "text", "text": page_note})

    # ── Block D: Instruction ─────────────────────────────────────────────────
    instruction_lines = [
        "Please normalize the financial statement from the attached PDF.",
        "Follow all steps in the system prompt exactly.",
        "Use the COA Master table provided above for COA Datapoint mapping.",
        "Return valid JSON only — no markdown, no extra text outside the JSON.",
    ]

    if reporting_columns:
        instruction_lines += [
            "",
            "[[REPORTING COLUMNS]]",
            *reporting_columns,
            "",
            "The [[REPORTING COLUMNS]] list above is the sole authority for column names "
            "and count. Do NOT detect or infer columns from the PDF header.",
        ]

    user_content.append({"type": "text", "text": "\n".join(instruction_lines)})

    return [
        {"role": "system", "content": system_prompt},
        {"role": "user",   "content": user_content},
    ]


# ─────────────────────────────────────────────────────────────────────────────
# 5.  CALL OPENAI
# ─────────────────────────────────────────────────────────────────────────────

def call_openai(
    messages: list[dict],
    model: str,
    max_tokens: int,
    temperature: float,
) -> str:
    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        sys.exit(
            "\n[ERROR] OPENAI_API_KEY environment variable is not set.\n"
            "        Run:  export OPENAI_API_KEY=sk-...\n"
        )

    client = OpenAI(api_key=api_key)

    print(f"\n[INFO] Model            : {model}")
    print(f"[INFO] Max tokens       : {max_tokens:,}")
    print(f"[INFO] Temperature      : {temperature}")
    print("[INFO] Sending request  : please wait (30–120s for large PDFs)...")

    try:
        response = client.chat.completions.create(
            model=model,
            max_tokens=max_tokens,
            temperature=temperature,
            messages=messages,
        )
    except APITimeoutError:
        sys.exit("\n[ERROR] Request timed out. Try --max-tokens 4000 or a smaller PDF.")
    except RateLimitError:
        sys.exit("\n[ERROR] Rate limit hit. Wait a moment and retry.")
    except APIError as e:
        sys.exit(f"\n[ERROR] OpenAI API error: {e}")

    u = response.usage
    print(f"\n[INFO] Token usage:")
    print(f"       Prompt tokens     : {u.prompt_tokens:,}")
    print(f"       Completion tokens : {u.completion_tokens:,}")
    print(f"       Total tokens      : {u.total_tokens:,}")

    return (response.choices[0].message.content or "").strip()


# ─────────────────────────────────────────────────────────────────────────────
# 6.  PARSE JSON RESPONSE
# ─────────────────────────────────────────────────────────────────────────────

def parse_json(raw: str) -> dict:
    text = raw.strip()

    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass

    fence = re.search(r"```(?:json)?\s*([\s\S]+?)\s*```", text)
    if fence:
        try:
            return json.loads(fence.group(1))
        except json.JSONDecodeError:
            pass

    try:
        start = text.index("{")
        end   = text.rindex("}") + 1
        return json.loads(text[start:end])
    except (ValueError, json.JSONDecodeError):
        pass

    debug = Path("_raw_response_debug.txt")
    debug.write_text(raw, encoding="utf-8")
    sys.exit(
        f"\n[ERROR] Could not parse response as valid JSON.\n"
        f"        Debug file saved : {debug}\n"
        f"        First 500 chars  :\n\n{raw[:500]}\n"
    )


# ─────────────────────────────────────────────────────────────────────────────
# 7.  VALIDATE OUTPUT
# ─────────────────────────────────────────────────────────────────────────────

_TOP_KEYS = {"Metadata", "Reporting Columns", "Sections"}

_EXPECTED_SECTIONS = {
    "Assets",
    "Deferred Outflows of Resources",
    "Liabilities",
    "Deferred Inflows of Resources",
    "Net Position",
}


def validate(data: dict) -> None:
    missing_top = _TOP_KEYS - set(data.keys())
    if missing_top:
        print(f"[WARN] Missing top-level keys  : {missing_top}")

    sections      = data.get("Sections", {})
    present       = set(sections.keys())
    missing_sec   = _EXPECTED_SECTIONS - present
    if missing_sec:
        print(f"[WARN] Missing sections        : {missing_sec}")

    total_rows = sum(len(v) for v in sections.values())
    meta       = data.get("Metadata", {})
    cols       = data.get("Reporting Columns", [])

    print(f"\n[INFO] ── Output Summary ──────────────────────────────")
    print(f"[INFO] Issuer Name       : {meta.get('Issuer Name', 'N/A')}")
    print(f"[INFO] Statement         : {meta.get('Statement', 'N/A')}")
    print(f"[INFO] FYE               : {meta.get('FYE', 'N/A')}")
    print(f"[INFO] Page No           : {meta.get('Page No', 'N/A')}")
    print(f"[INFO] Currency          : {meta.get('Currency reported', 'N/A')}")
    print(f"[INFO] Reporting columns : {cols}")
    print(f"[INFO] Sections          : {list(sections.keys())}")
    print(f"[INFO] Total rows        : {total_rows}")


# ─────────────────────────────────────────────────────────────────────────────
# 8.  MAIN
# ─────────────────────────────────────────────────────────────────────────────

def main() -> None:
    args = parse_args()

    banner = "Financial Statement Normalizer — OpenAI"
    print("=" * 60)
    print(f"  {banner}")
    print("=" * 60)
    print(f"[INFO] PDF              : {args.pdf}")
    print(f"[INFO] XLSX             : {args.xlsx}")
    print(f"[INFO] Prompt           : {args.prompt}")
    print(f"[INFO] Output           : {args.output}")
    print()

    # ── Step 1: Load all inputs ───────────────────────────────────────────────
    print("[STEP 1/5] Loading inputs...")
    system_prompt          = load_text_file(args.prompt, "Prompt")
    pdf_data_uri, pdf_name = load_pdf_as_data_uri(args.pdf)
    coa_text               = load_xlsx_as_pipe_text(args.xlsx)

    # Parse page numbers from the PDF filename (works for both pipeline-produced
    # files like CityReport_SNP_p15.pdf and old-style CityReport_SNP.pdf)
    page_info = extract_page_info_from_filename(args.pdf)
    if page_info:
        print(f"[INFO] {'Source pages':<22} {page_info['label']}   (from filename)")
    else:
        print(f"[INFO] {'Source pages':<22} [none] — no page tag in filename; "
              f"Metadata.Page No will rely on PDF content only")
    print()

    # ── Step 2: Build messages ────────────────────────────────────────────────
    print("[STEP 2/5] Building OpenAI messages...")
    messages = build_messages(
        system_prompt     = system_prompt,
        pdf_data_uri      = pdf_data_uri,
        pdf_filename      = pdf_name,
        coa_text          = coa_text,
        reporting_columns = args.reporting_columns,
        page_info         = page_info,
    )
    n_blocks   = len(messages[1]["content"])
    block_desc = "PDF file · COA master"
    if page_info:
        block_desc += " · source page reference"
    block_desc += " · instruction"
    if args.reporting_columns:
        block_desc += " · reporting columns"

    print(f"[INFO] Messages          : 2 (system + user)")
    print(f"[INFO] User content      : {n_blocks} blocks ({block_desc})")
    print()

    # ── Step 3: Call OpenAI ───────────────────────────────────────────────────
    print("[STEP 3/5] Calling OpenAI API...")
    raw_response = call_openai(
        messages    = messages,
        model       = args.model,
        max_tokens  = args.max_tokens,
        temperature = args.temperature,
    )
    print()

    # ── Step 4: Parse JSON ────────────────────────────────────────────────────
    print("[STEP 4/5] Parsing JSON response...")
    result = parse_json(raw_response)
    print("[INFO] JSON parsed successfully.")

    # ── Step 5: Validate and write output ─────────────────────────────────────
    print("\n[STEP 5/5] Validating and writing output...")
    validate(result)

    out_path = Path(args.output)
    out_path.write_text(
        json.dumps(result, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )

    print()
    print("=" * 60)
    print(f"  Done!  Output saved → {args.output}")
    print("=" * 60)


if __name__ == "__main__":
    main()