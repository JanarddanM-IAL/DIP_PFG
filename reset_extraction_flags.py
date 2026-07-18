# -*- coding: utf-8 -*-
"""
reset_extraction_flags.py — reset the extraction-stage columns of TProcessStatus rows
back to pending, so PFG_Extraction.py will re-pick them. Test/ops utility only.

Clears (to NULL) for the targeted rows:
    ExtractionStatus, ExtractionFlag, DataValidationStatus, DataValidationFlag,
    CompletionStatus, OutputPath, Remarks

It does NOT touch sourcing / sourcing-validation columns, so a reset row is still
"extraction-ready" (SourcingValidationStatus=1, SourcingValidationFlag='c').

Usage:
    python reset_extraction_flags.py 90 91      # reset these TProcessStatus.Id values
    python reset_extraction_flags.py --all       # reset every active COAID IN (1,2) row
    python reset_extraction_flags.py             # print usage (no changes)
"""

import sys
from db import Database, build_in_clause

_RESET_SETS = """
    ExtractionStatus = NULL,
    ExtractionFlag = NULL,
    DataValidationStatus = NULL,
    DataValidationFlag = NULL,
    CompletionStatus = NULL,
    OutputPath = NULL,
    Remarks = NULL,
    ModifiedOn = GETDATE()
"""


def reset_ids(row_ids):
    clause, params = build_in_clause("Id", row_ids)
    with Database() as db:
        res = db.update(f"UPDATE TProcessStatus SET {_RESET_SETS} "
                        f"WHERE IsActive = 1 AND {clause}", params)
        print(f"[RESET] rows affected: {res.rowcount}" if res.success
              else f"[ERROR] {res.error}")


def reset_all():
    with Database() as db:
        res = db.update(f"""
            UPDATE TProcessStatus SET {_RESET_SETS}
            WHERE IsActive = 1
              AND COAID IN (1, 2)
              AND SourcingValidationStatus = 1
              AND SourcingValidationFlag = 'c'
        """)
        print(f"[RESET] rows affected: {res.rowcount}" if res.success
              else f"[ERROR] {res.error}")


if __name__ == "__main__":
    argv = sys.argv[1:]
    if "--all" in argv:
        reset_all()
    elif argv and all(a.isdigit() for a in argv):
        reset_ids([int(a) for a in argv])
    else:
        print(__doc__)
