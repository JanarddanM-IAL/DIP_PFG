# -*- coding: utf-8 -*-
"""
seed_test_data.py — insert test rows into TProcessStatus (+ TProcessingAdditionalInfo)
so PFG_Sourcing.py / PFG_Validation.py have a work list to process.

Config (TModuleMaster/TSegmentMaster/TCOAMaster/TFileType) and TCompanyMaster are already
populated in the DIP database; this only seeds the empty execution tables.

Relation (current schema): TProcessingAdditionalInfo links to TProcessStatus by its own
`ID` column (FK -> TProcessStatus.Id, UNIQUE 1:1); its PK is `AdditionalInfoId` (identity).
So each seeded AdditionalInfo row is inserted with ID = the TProcessStatus.Id we just created.
TProcessStatus.ProcessingId is an independent deliverable value (does NOT need to equal Id).

Usage:
    python seed_test_data.py            # insert seed rows
    python seed_test_data.py --clear    # delete previous seed rows first, then insert
    python seed_test_data.py --clear-only
"""

import sys
from db import get_db_connection

# Companies to seed (LG sector, have UEI/State — valid FAC search inputs).
COMPANY_IDS = [3, 8, 9]

COAID        = 1      # PFG-COA-A (FK -> TCOAMaster.COAID; work list filters COAID IN (1,2))
FILE_TYPE_ID = 1      # Annual Report (AR)
PROCESS_YEAR = 2025
SEED_TAG     = "seed-test"   # CreatedBy marker so we can find/clear these rows


def clear_seed(cur):
    # Delete children first (FK), then the work rows — only our seed-tagged rows.
    cur.execute("""
        DELETE ai
        FROM TProcessingAdditionalInfo ai
        JOIN TProcessStatus ps ON ai.ID = ps.Id
        WHERE ps.CreatedBy = ?
    """, SEED_TAG)
    n_ai = cur.rowcount
    cur.execute("DELETE FROM TProcessStatus WHERE CreatedBy = ?", SEED_TAG)
    print(f"[CLEAR] Removed {cur.rowcount} TProcessStatus + {n_ai} TProcessingAdditionalInfo seed rows.")


def seed(cur):
    for i, cid in enumerate(COMPANY_IDS):
        processing_id = 101 + i   # independent deliverable id (need not equal Id)

        # 1) Insert the work row; let Id auto-generate and capture it via OUTPUT.
        cur.execute(
            """
            INSERT INTO TProcessStatus
                (ProcessingId, CompanyId, COAID, FileTypeId, ProcessYear, ProcessingCode,
                 SourcingStatus, SourcingFlag, CompletionStatus, IsActive, CreatedBy, CreatedOn)
            OUTPUT INSERTED.Id
            VALUES (?, ?, ?, ?, ?, NEWID(),
                    0, NULL, 0, 1, ?, GETDATE())
            """,
            processing_id, cid, COAID, FILE_TYPE_ID, PROCESS_YEAR, SEED_TAG,
        )
        new_id = cur.fetchone()[0]

        # 2) Matching AdditionalInfo row — linked by ID = TProcessStatus.Id
        #    (FK TProcessingAdditionalInfo.ID -> TProcessStatus.Id; AdditionalInfoId is identity).
        cur.execute(
            """
            INSERT INTO TProcessingAdditionalInfo (ID, CreatedBy, CreatedOn, IsActive)
            VALUES (?, ?, GETDATE(), 1)
            """,
            new_id, SEED_TAG,
        )
        print(f"[SEED] CompanyId={cid} -> TProcessStatus.Id={new_id} (ProcessingId={processing_id}), "
              f"AdditionalInfo.ID={new_id}, COAID={COAID}, FileTypeId={FILE_TYPE_ID}, Year={PROCESS_YEAR}")


def main():
    do_clear      = "--clear" in sys.argv or "--clear-only" in sys.argv
    do_seed       = "--clear-only" not in sys.argv

    conn = get_db_connection()
    cur = conn.cursor()
    try:
        if do_clear:
            clear_seed(cur)
        if do_seed:
            seed(cur)
        conn.commit()
        print("[DONE] Seed committed.")
    except Exception as e:  # noqa: BLE001
        conn.rollback()
        print(f"[ERROR] Rolled back: {e}")
        raise
    finally:
        conn.close()


if __name__ == "__main__":
    main()
