"""
backfill_coords.py
==================
One-shot script to add coordinates to ALL existing JSON outputs
without re-running any LLM calls.

Usage:
  python backfill_coords.py <output_folder> <raw_pdf_folder>

Example:
  python backfill_coords.py ./04_Validated_output ./03_Validated_Report

Logic:
  For each deal sub-folder in output_folder:
    For each .json file:
      Find the matching raw PDF in raw_pdf_folder by deal name
      Run attach_coordinates() and overwrite the JSON in-place
      (original backed up as .json.bak before overwrite)
"""

import json
import os
import re
import shutil
import sys
from pathlib import Path

from coordinate_extractor import attach_coordinates
from compact_schema import STATEMENT_SUFFIX_ALTERNATION


# ── Helpers ────────────────────────────────────────────────────────────────────

_STMT_SUFFIX_RE = re.compile(
    rf"_({STATEMENT_SUFFIX_ALTERNATION})_p\d+",
    re.IGNORECASE,
)


def get_deal_name_from_json(json_stem: str) -> str:
    """Strip statement-type + page suffix to get the deal base name."""
    return _STMT_SUFFIX_RE.split(json_stem)[0]


def find_raw_pdf(deal_name: str, raw_folder: str) -> str | None:
    """Locate the raw (un-sliced) PDF for a deal in the raw folder."""
    for f in os.listdir(raw_folder):
        if not f.lower().endswith(".pdf"):
            continue
        stem = Path(f).stem
        # Match stem start — deal_name is a prefix of the raw filename
        if stem.startswith(deal_name) or Path(f).stem == deal_name:
            return os.path.join(raw_folder, f)
    return None


def find_extracted_pdf(json_path: str, raw_folder: str) -> str | None:
    """
    Find the sliced/extracted PDF that matches this JSON.
    Strategy: look for same stem + .pdf in the deal sub-folder first,
    then fall back to the raw folder.
    """
    json_path = Path(json_path)
    stem = json_path.stem  # e.g. LG_CIT_WI_600005841_2024_SNP_p4-5

    # 1. Same directory as the JSON
    candidate = json_path.parent / (stem + ".pdf")
    if candidate.exists():
        return str(candidate)

    # 2. Raw folder — exact filename match
    candidate2 = Path(raw_folder) / (stem + ".pdf")
    if candidate2.exists():
        return str(candidate2)

    # 3. Raw folder — search by deal name + statement suffix
    deal_name = get_deal_name_from_json(stem)
    for f in os.listdir(raw_folder):
        if f.lower().endswith(".pdf") and stem.lower() in f.lower():
            return os.path.join(raw_folder, f)

    return None


# ── Main ────────────────────────────────────────────────────────────────────────

def backfill(output_folder: str, raw_folder: str, backup: bool = True) -> None:
    output_folder = Path(output_folder)
    raw_folder    = Path(raw_folder)

    if not output_folder.is_dir():
        print(f"[ERROR] output_folder not found: {output_folder}")
        sys.exit(1)
    if not raw_folder.is_dir():
        print(f"[ERROR] raw_folder not found: {raw_folder}")
        sys.exit(1)

    # Collect all deal sub-folders
    deal_dirs = [d for d in output_folder.iterdir() if d.is_dir()]
    if not deal_dirs:
        # Flat output structure — treat output_folder itself
        deal_dirs = [output_folder]

    total_json = 0
    enriched   = 0
    skipped    = 0

    for deal_dir in sorted(deal_dirs):
        json_files = list(deal_dir.glob("*.json"))
        if not json_files:
            continue

        print(f"\n[DEAL] {deal_dir.name}")

        for json_file in sorted(json_files):
            total_json += 1
            stem = json_file.stem

            # Find matching extracted PDF
            pdf_path = find_extracted_pdf(str(json_file), str(raw_folder))

            # If not found, also check the deal_dir itself
            if pdf_path is None:
                candidate = deal_dir / (stem + ".pdf")
                if candidate.exists():
                    pdf_path = str(candidate)

            if pdf_path is None:
                print(f"  [SKIP] {json_file.name} — no matching PDF found")
                skipped += 1
                continue

            print(f"  [OK]   {json_file.name}")
            print(f"         PDF → {Path(pdf_path).name}")

            # Load JSON
            try:
                with open(json_file, encoding="utf-8") as f:
                    data = json.load(f)
            except Exception as e:
                print(f"  [SKIP] JSON parse error: {e}")
                skipped += 1
                continue

            # Enrich
            try:
                enriched_data = attach_coordinates(data, pdf_path)
            except Exception as e:
                print(f"  [SKIP] Coordinate extraction error: {e}")
                skipped += 1
                continue

            # Back up original
            if backup:
                bak_path = json_file.with_suffix(".json.bak")
                shutil.copy2(json_file, bak_path)

            # Write enriched
            with open(json_file, "w", encoding="utf-8") as f:
                json.dump(enriched_data, f, indent=2, ensure_ascii=False)

            enriched += 1

    print(f"\n{'='*60}")
    print(f"  Backfill complete.")
    print(f"  JSON files found : {total_json}")
    print(f"  Enriched         : {enriched}")
    print(f"  Skipped          : {skipped}")
    if backup:
        print(f"  Originals backed up as .json.bak")
    print(f"{'='*60}")


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python backfill_coords.py <output_folder> <raw_pdf_folder>")
        sys.exit(1)

    backfill(
        output_folder = sys.argv[1],
        raw_folder    = sys.argv[2],
        backup        = True,
    )