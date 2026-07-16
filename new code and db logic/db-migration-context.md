# DIP Schema — Context Primer (for AI-assisted development / Excel→DB migration)

> **How to use:** Paste this file into Claude (or any LLM) at the start of a session that touches the DIP
> extraction schema or the Excel→DB migration, so it has full context without re-explaining. It is the
> compact, self-contained companion to `schema-onboarding.md` / `schema-onboarding.html`.
> Source of truth for keys = `DB_schema.xlsx` (9 original tables) **+ the new `TSegmentFileTypeMapping`** = 10 tables.
> **Naming convention: PascalCase** for all tables/columns; acronyms (`COA`, `UEI`, `EIN`) stay UPPERCASE; ids are `…Id`.
> *(Tip: if you use Claude Code, renaming/copying this to `CLAUDE.md` in the project root makes it auto-load every session.)*

---

## 1. What the system is

A **multi-tenant document-extraction platform** (financial + ESG data). Clients subscribe to **output formats
(COAs = charts of accounts)** for a set of **companies**. A client-agnostic background pipeline sources each
company's documents (Annual Report, Sustainability Report, FAC…), extracts them into each requested COA format,
and tracks every unit of work in `TProcessStatus`.

**Three layers:**
- **① Config** — what can be produced & how: `TModuleMaster → TSegmentMaster → TCOAMaster`, plus `TSegmentFileTypeMapping` (which file types each segment needs)
- **② Subscription** — who wants what, for which companies: `TClientMaster`, `TClientCOAMapping`, `TClientCOACompanyMapping`, `TCompanyMaster`
- **③ Execution** — the flattened, deduplicated work list: `TFileType`, `TProcessStatus`

Migration note: current code is **Excel-based**; the tables below are the target DB structure. Tables marked
**[existing]** already exist in the product DB; **[new]** are being introduced for this feature.

---

## 2. Tables & columns (migration reference)

Key legend: **PK** primary key · **FK→T** foreign key · audit block = `CreatedBy, CreatedOn, ModifiedBy, ModifiedOn, IsActive` (present on all tables except `TFileType`, which omits it — verify during migration).

### TModuleMaster  [existing]
`ModuleId` PK · `ModuleName` · `Sourcing` bit · `SourcingValidation` bit · `Extraction` bit · `ExtractionValidation` bit · `NumberOfThread` int · + audit block
- Drives pipeline behavior: the 4 bit flags gate which stages run; `NumberOfThread` = parallelism.

### TSegmentMaster  [new]
`SegmentId` PK · `SegmentName` · `ModuleId` FK→TModuleMaster · + audit block

### TCOAMaster  [new]
`COAId` PK · `COAName` · `SegmentId` FK→TSegmentMaster · `IsMatric` bit · + audit block
- `IsMatric` (likely intended "IsMetric"): `1` = metric/KPI extraction, `0` = financial-statement extraction.

### TSegmentFileTypeMapping  [new]
`SegmentFileTypeMappingId` PK · `SegmentId` FK→TSegmentMaster · `FileTypeId` FK→TFileType · `SequenceOrder` int · + audit block
- Which file types each **segment** requires, and their **merge order** (`SequenceOrder`). Drives row generation
  (see §4.7). `UNIQUE(SegmentId, FileTypeId) WHERE IsActive=1`.

### TClientMaster  [existing]
`ClientId` PK · `ClientUniqueId` · `ContactPerson` · `ClientName` · `CompanyName` · `EmailId` · `ClientAddress` · `CountryId` · `CityId` · `ZipCode` · `ContactNo` · `NoOfNodes` · `CreatedBy` · `CreatedOn` · `ModifiedBy` · `ModifiedOn` · `IsActive` · `IsDeactivateByAdmin` · `DeactivateByAdminBy` · `IsNotificationSubscribed` · `UnsubscribedDate` · `AuthenticationTypeId`
- Note: has TWO deactivation fields (`IsActive`, `IsDeactivateByAdmin`); confirm precedence in migration.
- (These four were non-standard in the source Excel — `Isactive`, `Isdeactivatebyadmin`, `DeactivatebyadminBy`, `IsnotificationSubscribed` — normalized here to PascalCase; renaming existing columns is itself a migration step.)

### TClientCOAMapping  [new]
`ClientCOAMappingId` PK · `COAId` FK→TCOAMaster · `ClientId` FK→TClientMaster · + audit block
- One row = one subscription ("this client wants this COA").

### TCompanyMaster  [existing]
`CompanyId` PK · `IssuerName` · `State` · `Sector` · `SubSector` · `UEI` · `EIN` · + audit block

### TClientCOACompanyMapping  [new]
`ClientCOACompanyMappingId` PK · `ClientCOAMappingId` FK→TClientCOAMapping · `CompanyId` FK→TCompanyMaster · + audit block
- One row = one company inside a subscription. **This table is preconfigured by an upstream team.**

### TFileType  [existing]
`FileTypeId` PK · `FileType` · `Tag` · (+ CreatedBy/CreatedOn/ModifiedBy/ModifiedOn/IsActive present here too)

### TProcessStatus  [existing] — the pipeline work list
`Id` PK · `ProcessingId` · `COAId` FK→TCOAMaster · `CompanyId` FK→TCompanyMaster · `FileTypeId` FK→TFileType · `ProcessYear` · `ProcessingCode` (GUID) · `SourcingStatus` · `SourcingFlag` · `SourcingValidationStatus` · `SourcingValidationFlag` · `ExtractionStatus` · `ExtractionFlag` · `DataValidationStatus` · `DataValidationFlag` · `CompletionStatus` · `PdfDownloadLink` · `PdfFilePath` · `OutputPath` · `FyeDate` · `ReleaseDate` · `DownloadDate` · `Remarks` · + audit block
- **No `ClientId` column** — pipeline is client-agnostic (see §4).
- Stage tracking = a `…Status` (int) + `…Flag` (char) pair per stage. Naming alias: **`DataValidationStatus/Flag` = the module's `ExtractionValidation` stage.**

### FK relationship map
```
TModuleMaster 1─< TSegmentMaster 1─< TCOAMaster 1─< TClientCOAMapping 1─< TClientCOACompanyMapping >─1 TCompanyMaster
                        │                                    │
                        │ 1                                  1
                        └─< TSegmentFileTypeMapping >─1 TFileType        TClientMaster 1───────────────┘
                                                                          (TClientMaster 1─< TClientCOAMapping)

TCOAMaster 1─< TProcessStatus >─1 TCompanyMaster       TFileType 1─< TProcessStatus
```
`TProcessStatus` is **materialized** from `TClientCOACompanyMapping` (+ file types from the COA's segment via
`TSegmentFileTypeMapping`); it is NOT foreign-keyed to any client.

---

## 3. The two keys in TProcessStatus (most important concept)

- **`ProcessingId`** = one **logical deliverable** = `COAId + CompanyId + ProcessYear`.
  When a COA needs several source files merged into one output, its rows share ONE `ProcessingId` with different `FileTypeId`.
- **`ProcessingCode`** (a GUID) = one **physical source file** = `CompanyId + FileTypeId + ProcessYear`.
  Shared by every COA that consumes that file; used for file naming. Source once, reuse many.

Example (Company 15, 2025): the Annual Report GUID `CAF3D14C…` is shared by COA 1, 2 AND 3 (one download, 3 rows).
COA 3 also needs a Sustainability Report (own GUID) — COA 3's two rows share `ProcessingId 109` and merge into one output.

---

## 4. Business rules confirmed with the user (not all derivable from data)

1. **Row generation is upload-triggered.** `TClientCOACompanyMapping` (and the mappings above it) are
   **preconfigured by a separate/upstream team**. When a **new Excel is uploaded**, rows are inserted into
   `TProcessStatus` per that mapping.
2. **COAId is stamped from the subscription.** Generation walks `TClientCOACompanyMapping` → parent
   `ClientCOAMappingId` → `TClientCOAMapping.COAId`, and stamps that `COAId` onto each work row.
3. **Company set is per-(client, COA).** A client is NOT assumed to want every company for every COA.
   Real data: Client 1 → COA 1 for {11,12,13,14,15} but COA 2 for {11,12,15}.
4. **Client-agnostic dedup.** Two clients requesting the same (COA, Company, Year) collapse into ONE row/compute.
   That's why there is no `ClientId` in `TProcessStatus`.
5. **Resolution walk** (how a row knows what to do): `COAId → TCOAMaster.SegmentId → TSegmentMaster.ModuleId
   → TModuleMaster` (stage flags + `NumberOfThread`). `IsMatric` picks metric vs financial-statement extraction.
6. **Sourcing dedup rule.** Before sourcing, a row checks for an existing row with the same
   `(CompanyId, FileTypeId, ProcessYear)` where `SourcingValidationStatus = 1`; if found, it copies the paths
   (`PdfDownloadLink`, `PdfFilePath`) + sourcing statuses and reuses that row's GUID instead of re-downloading.
7. **COA → required FileType(s): segment-level junction `TSegmentFileTypeMapping` (DECIDED).** The required file
   types are a property of the **segment** (`LG` → Annual Report; `Global` → Annual Report + Sustainability Report),
   stored in `TSegmentFileTypeMapping(SegmentId, FileTypeId, SequenceOrder)`. Generation resolves
   `COAId → TCOAMaster.SegmentId → TSegmentFileTypeMapping` (ORDER BY `SequenceOrder`) → the file types to create
   rows for — order-safe, FK-enforced, no string parsing. See `TSegmentFileTypeMapping.sql`. Chosen over a per-COA
   comma-separated column because the requirement is uniform per segment. **Revisit → move mapping to COA level**
   only if two COAs in one segment need different file types; add `IsRequired`/effective-dating only if
   optional/per-year/per-company variance appears.

---

## 5. Sample data (as in DB_schema.xlsx — concrete reference)

**Modules** (flags S/SV/E/EV, threads):
`1 FinanceIQ` 1/1/1/0 t3 · `2 ClimateIQ` 1/1/0/1 t2 (no segments/COAs yet — idle) · `3 SustainIQ` 1/1/1/1 t3

**Segments:** `1 LG`→M1 · `2 Non-LG`→M1 (unused) · `3 Indian`→M3 (unused) · `4 Global`→M3

**COAs:** `1 LG-Global`→Seg1, IsMatric0 · `2 LG-ABC`→Seg1, IsMatric0 · `3 Sustainability Global`→Seg4, IsMatric1

**TSegmentFileTypeMapping (new — seed):** Seg 1 (LG) → FileType 1 (AR) seq1 · Seg 4 (Global) → FileType 1 (AR) seq1, FileType 2 (SR) seq2. *(Segments 2 & 3 unseeded — no COAs yet.)*

**Clients:** `1` (ClientName 'y') · `2` ('C4F_USER'). *(Client IDs/names are placeholder sample data; both rows share one ClientUniqueId — treat as sample noise.)*

**TClientCOAMapping:** `m1`=(Client1,COA1) · `m2`=(Client1,COA2) · `m3`=(Client2,COA3)

**Companies (all Sector=LG):** 11 CITY OF MOUNTAIN VIEW (CA,CIT) · 12 PINELLAS COUNTY (FL,CNT) · 13 CHEROKEE COUNTY (NC,CNT) · 14 ASCENSION PARISH SCHOOL BOARD (LA,SD) · 15 CITY AND BOROUGH OF JUNEAU (AK,CIT) · 16 CITY OF OTTUMWA (IA,CIT) · 17 TOWN OF LITTLETON (MA,CIT)

**TClientCOACompanyMapping:** m1→{11,12,13,14,15} · m2→{11,12,15} · m3→{15,16,17}

**FileTypes:** `1` Annual Report (AR) · `2` Sustainability Report (SR) · `3` FAC (unused in sample)

**TProcessStatus (14 rows; all statuses=1, flags='c', year=2025):**
`Id | ProcessingId | COA | Company | FileType | ProcessingCode`
```
 1 | 101 | 1 | 11 | AR | D1CD4E29
 2 | 102 | 2 | 11 | AR | D1CD4E29   (reuses row 1's file)
 3 | 103 | 1 | 12 | AR | E707826A
 4 | 104 | 2 | 12 | AR | E707826A   (reuses row 3's file)
 5 | 105 | 1 | 13 | AR | 33AF5DB3
 6 | 106 | 1 | 14 | AR | BC050486
 7 | 107 | 1 | 15 | AR | CAF3D14C
 8 | 108 | 2 | 15 | AR | CAF3D14C   (reuses row 7's file)
 9 | 109 | 3 | 15 | AR | CAF3D14C   (reuses row 7's file)
10 | 109 | 3 | 15 | SR | 8CD43CE0   (merges into ProcessingId 109)
11 | 110 | 3 | 16 | AR | 79AE6223
12 | 110 | 3 | 16 | SR | 19F6ABCD   (merges into ProcessingId 110)
13 | 111 | 3 | 17 | AR | A14D9B8E
14 | 111 | 3 | 17 | SR | 53D9C70D   (merges into ProcessingId 111)
```
14 rows = 11 distinct ProcessingIds (deliverables) + 10 distinct GUIDs (source files). 4 rows reuse a file already sourced.

---

## 6. Known gaps / decisions to weigh during migration

The static model is sound (keep the two-key design, client-agnostic sharing, config-driven gating). Runtime/integrity
gaps to address as the DB is built:

1. **Split physical file into its own table** `TSourceFile` (one row per `CompanyId+FileTypeId+ProcessYear`, owns
   GUID/paths/sourcing status); `TProcessStatus` references it by FK. Removes the copy-from-sibling dedup hack and
   the stale-copy risk. *(Critical)*
2. **Concurrency:** no claim/lease today; the dedup is racy and won't fire on the initial bulk pass. Add
   `ClaimedBy/LeaseUntil/AttemptCount` (or a real queue); make source a lookup-or-create (`INSERT … ON CONFLICT`). *(Critical)*
3. **Unique constraints:** none enforced. Add `UNIQUE(ClientUniqueId)`, `UNIQUE(COAId,CompanyId,FileTypeId,ProcessYear)`
   on the work list (so re-uploads don't duplicate), unique source-file identity, unique keys on both mapping tables,
   unique UEI/EIN. Use partial indexes (`WHERE IsActive=1`) to coexist with soft-delete. *(Critical)*
4. **Stage state:** replace `Status(int)+Flag(char)` pairs with a single enum per stage
   (`Pending/Running/Succeeded/Failed/Skipped/DeadLettered`) + an append-only event/history table (per-stage
   timestamps, attempts, error). Distinguishes "skipped" from "done". *(High)*
5. **Multi-tenancy:** add a client fulfillment layer on top of the shared work list for per-client SLA / GDPR-deletion
   / billing / notifications (which the client-agnostic sharing otherwise makes impossible). *(High)*
6. **COA→FileType is now data-driven** via `TSegmentFileTypeMapping` (see §4.7) — done, not a gap. Move to COA-level
   granularity only if two COAs in one segment ever diverge.
7. **Naming cleanup:** `ExtractionValidation` vs `DataValidation*`; `IsMatric`→`IsMetric`;
   `NumberOfThread` arguably per-stage not per-module. (Table/column names are now standardized to PascalCase.)

---

## 7. Quick glossary
- **COA** — Chart of Accounts = an output format a client subscribes to (the leaf of the config chain).
- **Module** — processing engine (FinanceIQ/ClimateIQ/SustainIQ) defining enabled stages + threads.
- **Segment** — grouping of COAs under a module; also carries the required file-type set (`TSegmentFileTypeMapping`).
- **ProcessingId** — one deliverable (COA+Company+Year); groups merged multi-file rows.
- **ProcessingCode** — GUID for one physical source file (Company+FileType+Year); shared across COAs; file-naming key.
- **Materialization** — generating TProcessStatus rows from the mappings on Excel upload.
