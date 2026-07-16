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
    """
    Detect & fix single-value DP rows where the LLM placed the value in the
    wrong reporting column. Strategy:
      For each CP 'Total ...' row in a section:
        1) sum all DPs per column
        2) compute delta = (sum - reported) for each column
        3) if exactly ONE single-value DP has |value| == |delta_filled|
           AND delta_filled == +value AND delta_other == -value,
           swap that DP's value to the other column.
    Mutates `data` in place. Logs every change under data["_column_shift_repairs"].
    """
    sections = data.get("Sections", {}) or {}
    cols     = data.get("Reporting Columns", []) or []
    if len(cols) < 2:
        return data  # nothing to swap on single-period statements

    ikey = _items_key(sections)
    log  = []

    for sec_name, rows in sections.items():
        dp_rows = [r for r in rows if r.get("COA Flag") == "DP"]
        cp_rows = [r for r in rows if r.get("COA Flag") == "CP"]
        if not dp_rows or not cp_rows:
            continue

        for cp in cp_rows:
            label = (cp.get(ikey) or "").strip().lower()
            if not label.startswith("total"):
                continue

            # Candidate = DP with exactly one filled column, one blank column
            candidates = []
            for dp in dp_rows:
                vals = {c: _to_dec(dp.get(c)) for c in cols}
                filled = [c for c, v in vals.items() if v is not None]
                blanks = [c for c, v in vals.items() if v is None]
                if len(filled) == 1 and len(blanks) == 1:
                    candidates.append((dp, filled[0], blanks[0], vals[filled[0]]))
            if not candidates:
                continue

            # Current DP sums per column
            sums = {c: Decimal(0) for c in cols}
            for dp in dp_rows:
                for c in cols:
                    v = _to_dec(dp.get(c))
                    if v is not None:
                        sums[c] += v

            reported = {c: _to_dec(cp.get(c)) for c in cols}
            if any(reported[c] is None for c in cols):
                continue

            deltas = {c: sums[c] - reported[c] for c in cols}

            # Look for the unique candidate whose swap reconciles BOTH columns
            for dp, fcol, bcol, val in candidates:
                if deltas[fcol] == val and deltas[bcol] == -val:
                    dp[bcol] = dp[fcol]   # move value to the right column
                    dp[fcol] = "-"
                    log.append({
                        "section": sec_name,
                        "row": dp.get(ikey, ""),
                        "moved_value": f"{val:,}",
                        "from_col": fcol,
                        "to_col":   bcol,
                        "cp_anchor": cp.get(ikey, ""),
                        "stmt_type": stmt_type,
                    })
                    # update running state in case another CP in this section
                    # also depends on the same DP
                    sums[fcol] -= val
                    sums[bcol] += val
                    deltas = {c: sums[c] - reported[c] for c in cols}
                    break

    if log:
        data.setdefault("_column_shift_repairs", []).extend(log)
        for entry in log:
            print(f"   [REPAIR] {entry['section']} :: '{entry['row']}' "
                  f"moved {entry['moved_value']} from {entry['from_col']} "
                  f"→ {entry['to_col']} (anchor: {entry['cp_anchor']})")
    return data
