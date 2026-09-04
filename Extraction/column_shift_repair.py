# column_shift_repair.py
from decimal import Decimal, InvalidOperation

def _to_dec(v):
    if v is None: return None
    s = str(v).strip()
    if s in ("", "-", "–", "—", "null"): return None
    s = s.replace(",", "").replace("$", "").strip()
    if s.startswith("(") and s.endswith(")"):
        s = "-" + s[1:-1]
    try:
        return Decimal(s)
    except InvalidOperation:
        return None

def _items_key(sections):
    for rows in sections.values():
        for r in rows:
            for k in r:
                if k.endswith(" Items") or k == "Items":
                    return k
    return "Items"

def repair_column_shifts(data: dict, stmt_type: str) -> dict:
    sections = data.get("Sections", {}) or {}
    cols     = data.get("Reporting Columns", []) or []
    if len(cols) < 3:
        return data  # need GA + BTA + Total minimum

    ikey = _items_key(sections)
    log  = []

    # ── Identify addend cols and sum col once ──────────────────────────────
    # Sum col = first col whose header contains "total" (case-insensitive)
    # Addend cols = all cols to the LEFT of the sum col
    sum_col_name = next(
        (c for c in cols if "total" in c.lower()), None
    )
    if sum_col_name is None:
        return data  # no Total column — cannot detect swap

    sum_col_idx   = cols.index(sum_col_name)
    addend_cols   = cols[:sum_col_idx]          # GA, BTA (everything left of Total)

    if len(addend_cols) < 2:
        return data  # need at least 2 addends to have a swap

    for sec_name, rows in sections.items():
        dp_rows = [r for r in rows if r.get("COA Flag") == "DP"]
        cp_rows = [r for r in rows if r.get("COA Flag") == "CP"]
        if not dp_rows or not cp_rows:
            continue

        for cp in cp_rows:
            label = (cp.get(ikey) or "").strip().lower()
            if not label.startswith("total"):
                continue

            # ── Find swap candidates ───────────────────────────────────────
            # A candidate has:
            #   - exactly one addend filled, one addend blank
            #   - filled addend value == Total column value  (identity trap)
            candidates = []
            for dp in dp_rows:
                vals = {c: _to_dec(dp.get(c)) for c in cols}

                addend_vals   = {c: vals[c] for c in addend_cols}
                filled_addends = [c for c, v in addend_vals.items() if v is not None]
                blank_addends  = [c for c, v in addend_vals.items() if v is None]

                if len(filled_addends) != 1 or len(blank_addends) != 1:
                    continue  # not a single-addend row

                filled_col = filled_addends[0]
                blank_col  = blank_addends[0]
                v          = addend_vals[filled_col]
                total_val  = vals.get(sum_col_name)

                # Identity trap: filled addend == Total → ambiguous assignment
                if total_val is not None and v == total_val:
                    candidates.append((dp, filled_col, blank_col, v))

            if not candidates:
                continue

            # ── Compute current DP sums per addend column ──────────────────
            sums = {c: Decimal(0) for c in addend_cols}
            for dp in dp_rows:
                for c in addend_cols:
                    v = _to_dec(dp.get(c))
                    if v is not None:
                        sums[c] += v

            reported = {c: _to_dec(cp.get(c)) for c in addend_cols}
            if any(reported[c] is None for c in addend_cols):
                continue

            deltas = {c: sums[c] - reported[c] for c in addend_cols}

            # ── Apply swap if exactly one candidate reconciles both addends ─
            for dp, fcol, bcol, val in candidates:
                if deltas[fcol] == val and deltas[bcol] == -val:
                    dp[bcol] = dp[fcol]
                    dp[fcol] = "-"
                    log.append({
                        "section":    sec_name,
                        "row":        dp.get(ikey, ""),
                        "moved_value": f"{val:,}",
                        "from_col":   fcol,
                        "to_col":     bcol,
                        "cp_anchor":  cp.get(ikey, ""),
                        "stmt_type":  stmt_type,
                    })
                    sums[fcol] -= val
                    sums[bcol] += val
                    deltas = {c: sums[c] - reported[c] for c in addend_cols}
                    break

    if log:
        data.setdefault("_column_shift_repairs", []).extend(log)
        for entry in log:
            print(f"   [REPAIR] {entry['section']} :: '{entry['row']}' "
                  f"moved {entry['moved_value']} from {entry['from_col']} "
                  f"→ {entry['to_col']} (anchor: {entry['cp_anchor']})")
    return data