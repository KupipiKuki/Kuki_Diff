"""
Excel Diff & Compare engine -- Unified Package.

Compares two .xlsx/.xlsm/.csv/.tsv datasets or directories and generates styled
reports highlighting character-level differences in red rich text. Matches rows
by identity independent of row order, with duplicate-safe tiered keys.

Features:
  - Single file comparison (generate_diff_report)
  - Batch directory comparison (compare_directories, find_file_pairs)
  - Paired detail tabs ("(Added)" / "(Removed)") and Stacked Submittal layout
  - Summary dashboard with PASS/FAIL integrity audit and structural checks
  - Fully configurable arguments for sheet names, header rows, skiprows,
    usecols, column names, na values, match keys/tiers, and group labels
  - Graceful autofilter error recovery and rich-text/strikethrough inspection
"""

import os
import re
import difflib
from datetime import datetime

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.styles.colors import Color
from openpyxl.cell.text import InlineFont
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.utils import get_column_letter

try:
    import string_checker as sc_mod
    import theme_color as tc_mod
except ImportError:
    try:
        from . import string_checker as sc_mod
        from . import theme_color as tc_mod
    except Exception:
        sc_mod = None
        tc_mod = None

# ---------------------------------------------------------------------------
# Palette / styling constants
# ---------------------------------------------------------------------------
ADDED_COLOR = "00FF0000"    # red   -- inserted / replaced text
REMOVED_COLOR = "0000B050"  # green -- deleted / replaced text (struck through)

FILL_CHANGED = PatternFill(patternType="solid", fgColor=Color(rgb="00FBE2D5"))  # light orange
FILL_ROW_ADDED = PatternFill(patternType="solid", fgColor=Color(rgb="00E2EFDA"))  # light green
FILL_ROW_REMOVED = PatternFill(patternType="solid", fgColor=Color(rgb="00E2EFDA"))  # light green
FILL_HEADER = PatternFill(patternType="solid", fgColor=Color(rgb="00D9D9D9"))
FILL_TITLE = PatternFill(patternType="solid", fgColor=Color(rgb="00244062"))
FILL_BANNER = PatternFill(patternType="solid", fgColor=Color(rgb="00163CF0"))  # royal blue for submittal banner
FILL_PASS = PatternFill(patternType="solid", fgColor=Color(rgb="00C6EFCE"))
FILL_FAIL = PatternFill(patternType="solid", fgColor=Color(rgb="00FFC7CE"))

FONT_TITLE = Font(bold=True, size=14, color="00FFFFFF")
FONT_BANNER = Font(bold=True, size=11, color="00FFFFFF")
FONT_HEADER = Font(bold=True)
FONT_PASS = Font(bold=True, color="00006100")
FONT_FAIL = Font(bold=True, color="009C0006")

THIN = Side(style="thin", color="00BFBFBF")
BORDER_ALL = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

STATUS_COLUMN = "Diff Status"

# Summary-table column indices (1-based)
COL_PCT = 16
COL_INTEGRITY = 17
COL_NOTES = 18

# ---------------------------------------------------------------------------
# Cell-level rich text helpers
# ---------------------------------------------------------------------------


def _norm(value):
    """Normalize any cell value to a plain string."""
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    return str(value)


def diff_strings_rich_text(old_text, new_text, highlight_color=ADDED_COLOR,
                           old_highlight_color=REMOVED_COLOR):
    """
    ADDED view. Returns the NEW string with inserted/replaced characters in red
    bold as an openpyxl CellRichText. Returns a plain str when identical.
    """
    old_text = _norm(old_text)
    new_text = _norm(new_text)

    if old_text == new_text:
        return new_text

    matcher = difflib.SequenceMatcher(None, old_text, new_text)
    highlight_font = InlineFont(color=highlight_color, b=True)

    rt = CellRichText()
    for tag, _i1, _i2, j1, j2 in matcher.get_opcodes():
        segment = new_text[j1:j2]
        if not segment:
            continue
        if tag == "equal":
            rt.append(segment)
        elif tag in ("insert", "replace"):
            rt.append(TextBlock(highlight_font, segment))

    if not len(rt):
        return new_text
    return rt


def diff_strings_removed_rich_text(old_text, new_text, highlight_color=REMOVED_COLOR):
    """
    REMOVED view. Returns the OLD string with deleted/replaced characters in
    green strikethrough as an openpyxl CellRichText. Returns a plain str when
    identical.
    """
    old_text = _norm(old_text)
    new_text = _norm(new_text)

    if old_text == new_text:
        return old_text

    matcher = difflib.SequenceMatcher(None, old_text, new_text)
    strike_font = InlineFont(color=highlight_color, b=True, strike=True)

    rt = CellRichText()
    for tag, i1, i2, _j1, _j2 in matcher.get_opcodes():
        segment = old_text[i1:i2]
        if not segment:
            continue
        if tag == "equal":
            rt.append(segment)
        elif tag in ("delete", "replace"):
            rt.append(TextBlock(strike_font, segment))

    if not len(rt):
        return old_text
    return rt


def strikethrough_rich_text(text, highlight_color=REMOVED_COLOR):
    """Render an entire string in green strikethrough (fully removed rows)."""
    text = _norm(text)
    if not text:
        return ""
    rt = CellRichText()
    rt.append(TextBlock(InlineFont(color=highlight_color, b=True, strike=True), text))
    return rt


# ---------------------------------------------------------------------------
# Alignment helpers
# ---------------------------------------------------------------------------


def _union_columns(df_old, df_new):
    """New-file column order first, then columns that exist only in the old file."""
    cols = list(df_new.columns)
    cols += [c for c in df_old.columns if c not in df_new.columns]
    return cols


def _safe_sheet_title(base, suffix, used):
    """Excel caps sheet names at 31 chars; truncate the base and de-duplicate."""
    room = 31 - len(suffix)
    title = f"{_norm(base)[:room]}{suffix}"
    counter = 2
    while title in used:
        tag = f"_{counter}"
        title = f"{_norm(base)[:room - len(tag)]}{tag}{suffix}"
        counter += 1
    used.add(title)
    return title


# ---------------------------------------------------------------------------
# Order-independent row matching
# ---------------------------------------------------------------------------

MATCH_SIM_THRESHOLD = 0.62


def _is_int_like(text):
    """True when the value is an integer, tolerating '12', '12.0', ' 12 '."""
    s = str(text).strip()
    if s == "":
        return False
    try:
        return float(s).is_integer()
    except (TypeError, ValueError):
        return False


def _int_token(text):
    return str(int(float(str(text).strip())))


def _valid(value, rule):
    """Apply a single tier validator to one cell value."""
    s = _norm(value).strip()
    if rule in (None, "nonempty", "str", "string"):
        return s != ""
    if rule == "int":
        return _is_int_like(s)
    if isinstance(rule, (list, tuple, set, frozenset)):
        return s.lower() in {str(a).strip().lower() for a in rule}
    return s != ""


def _tier_spec(tier):
    """Normalize a tier into (columns, require-map)."""
    if isinstance(tier, dict):
        return list(tier.get("columns", [])), dict(tier.get("require", {}) or {})
    if isinstance(tier, str):
        return [tier], {}
    return list(tier), {}


def _tier_key(df, idx, columns, require):
    """
    Build the tier key for one row, or None when the row does not qualify.
    """
    parts = []
    for col in columns:
        if col not in df.columns:
            return None
        raw = df.iloc[idx][col]
        rule = require.get(col)
        if not _valid(raw, rule):
            return None
        s = _norm(raw).strip()
        parts.append(_int_token(s) if rule == "int" else s.lower())
    return tuple(parts) if parts else None


def _row_signature(df, idx, columns):
    return tuple(_cell(df, idx, c) for c in columns)


def _similarity(df_old, old_idx, df_new, new_idx, columns):
    """Mean per-column string similarity across shared columns."""
    if not columns:
        return 0.0
    total = 0.0
    for col in columns:
        o = _cell(df_old, old_idx, col)
        n = _cell(df_new, new_idx, col)
        if o == n:
            total += 1.0
        elif not o and not n:
            total += 1.0
        elif not o or not n:
            total += 0.0
        else:
            total += difflib.SequenceMatcher(None, o, n).ratio()
    return total / len(columns)


def _auto_tiers(df_old, df_new, shared_cols):
    """
    Rank shared columns for use as identity keys.
    """
    scored = []
    for col in shared_cols:
        o = [_cell(df_old, i, col) for i in range(len(df_old))]
        n = [_cell(df_new, i, col) for i in range(len(df_new))]
        o_nb = [v for v in o if v.strip()]
        n_nb = [v for v in n if v.strip()]
        if not o_nb or not n_nb:
            continue
        so, sn = set(o_nb), set(n_nb)
        uniq = (len(so) / len(o_nb)) * (len(sn) / len(n_nb))
        overlap = len(so & sn) / max(1, min(len(so), len(sn)))
        coverage = min(len(o_nb) / len(o), len(n_nb) / len(n))
        score = uniq * overlap * coverage
        if uniq >= 0.5 and overlap >= 0.3:
            scored.append((score, col))
    scored.sort(reverse=True)
    return [[col] for _score, col in scored[:4]]


def _greedy_similarity_pairs(df_old, df_new, old_left, new_left, shared_cols,
                             threshold=MATCH_SIM_THRESHOLD):
    """Best-first pairing of leftover rows; only pairs above the threshold."""
    if not old_left or not new_left or not shared_cols:
        return []
    cands = []
    for oi in old_left:
        for ni in new_left:
            sim = _similarity(df_old, oi, df_new, ni, shared_cols)
            if sim >= threshold:
                cands.append((-sim, oi, ni))
    cands.sort()
    used_o, used_n, pairs = set(), set(), []
    for _neg, oi, ni in cands:
        if oi in used_o or ni in used_n:
            continue
        used_o.add(oi)
        used_n.add(ni)
        pairs.append((oi, ni))
    return pairs


def _match_rows(df_old, df_new, match_keys, shared_cols):
    """Match rows across files independent of row order."""
    old_left = list(range(len(df_old)))
    new_left = list(range(len(df_new)))
    old_to_new = {}
    tier_log = []
    ambiguous = {"old": [], "new": []}

    def consume(pairs, label):
        if not pairs:
            return
        for oi, ni in pairs:
            old_to_new[oi] = ni
        done_o = {o for o, _ in pairs}
        done_n = {n for _, n in pairs}
        old_left[:] = [i for i in old_left if i not in done_o]
        new_left[:] = [i for i in new_left if i not in done_n]
        tier_log.append((label, len(pairs)))

    # --- Tier 1: identical rows, any order -----------------------------------
    if shared_cols:
        buckets = {}
        for ni in new_left:
            buckets.setdefault(_row_signature(df_new, ni, shared_cols), []).append(ni)
        pairs = []
        for oi in old_left:
            sig = _row_signature(df_old, oi, shared_cols)
            if buckets.get(sig):
                pairs.append((oi, buckets[sig].pop(0)))
        consume(pairs, "identical row")

    # --- Tier 2..n: key tiers ------------------------------------------------
    for tier in match_keys:
        if not old_left or not new_left:
            break
        columns, require = _tier_spec(tier)
        if not columns:
            continue
        label = "+".join(str(c) for c in columns)

        o_buckets, n_buckets = {}, {}
        for oi in old_left:
            k = _tier_key(df_old, oi, columns, require)
            if k is not None:
                o_buckets.setdefault(k, []).append(oi)
        for ni in new_left:
            k = _tier_key(df_new, ni, columns, require)
            if k is not None:
                n_buckets.setdefault(k, []).append(ni)

        exact, deferred = [], []
        for k, o_rows in o_buckets.items():
            n_rows = n_buckets.get(k)
            if not n_rows:
                continue
            if len(o_rows) == 1 and len(n_rows) == 1:
                exact.append((o_rows[0], n_rows[0]))
            else:
                deferred.append((k, o_rows, n_rows))
        consume(exact, label)

        dup_pairs = []
        for k, o_rows, n_rows in deferred:
            o_rows = [i for i in o_rows if i in old_left]
            n_rows = [i for i in n_rows if i in new_left]
            if not o_rows or not n_rows:
                continue
            key_txt = "|".join(k)
            if len(o_rows) > 1:
                ambiguous["old"].append(f"{label}={key_txt}")
            if len(n_rows) > 1:
                ambiguous["new"].append(f"{label}={key_txt}")
            inner = _greedy_similarity_pairs(df_old, df_new, o_rows, n_rows,
                                             shared_cols, threshold=0.0)
            dup_pairs.extend(inner)
        consume(dup_pairs, f"{label} (duplicate key)")

    # --- Final tier: similarity fallback -------------------------------------
    consume(_greedy_similarity_pairs(df_old, df_new, list(old_left),
                                     list(new_left), shared_cols),
            "similarity")

    ambiguous["old"] = sorted(set(ambiguous["old"]))
    ambiguous["new"] = sorted(set(ambiguous["new"]))
    return old_to_new, tier_log, ambiguous


def _plan_from_matches(df_old, df_new, old_to_new):
    """Order report rows: new-file order with removals placed in context."""
    plan = []
    last_new = -1
    for oi in range(len(df_old)):
        if oi in old_to_new:
            last_new = old_to_new[oi]
        else:
            plan.append(((last_new, 1, oi), {"key": None, "old_idx": oi, "new_idx": None}))
    new_to_old = {n: o for o, n in old_to_new.items()}
    for ni in range(len(df_new)):
        plan.append(((ni, 0, 0), {"key": None, "old_idx": new_to_old.get(ni), "new_idx": ni}))
    plan.sort(key=lambda p: p[0])
    return [row for _sort, row in plan]


def _align(df_old, df_new, key_column, match_keys="auto", shared_cols=None):
    """Build aligned row plan."""
    dup_keys = {"old": [], "new": []}
    meta = {"mode": "positional", "tiers": []}

    if key_column and key_column in df_old.columns and key_column in df_new.columns:
        old_keys = df_old[key_column].map(_norm)
        new_keys = df_new[key_column].map(_norm)

        dup_keys["old"] = sorted(set(old_keys[old_keys.duplicated()].tolist()))
        dup_keys["new"] = sorted(set(new_keys[new_keys.duplicated()].tolist()))

        old_pos, new_pos = {}, {}
        for pos, k in enumerate(old_keys):
            old_pos.setdefault(k, pos)
        for pos, k in enumerate(new_keys):
            new_pos.setdefault(k, pos)

        ordered = list(new_pos.keys()) + [k for k in old_pos if k not in new_pos]
        rows = [
            {"key": k, "old_idx": old_pos.get(k), "new_idx": new_pos.get(k)}
            for k in ordered
        ]
        meta["mode"] = f"key column '{key_column}'"
        return rows, dup_keys, meta

    if match_keys == "positional" or match_keys is None:
        rows = []
        for i in range(max(len(df_old), len(df_new))):
            rows.append({
                "key": None,
                "old_idx": i if i < len(df_old) else None,
                "new_idx": i if i < len(df_new) else None,
            })
        return rows, dup_keys, meta

    if shared_cols is None:
        shared_cols = [c for c in df_new.columns if c in df_old.columns]

    if match_keys == "auto":
        tiers = _auto_tiers(df_old, df_new, shared_cols)
        meta["mode"] = "order-independent (auto keys)"
    else:
        tiers = list(match_keys)
        meta["mode"] = "order-independent (explicit keys)"

    old_to_new, tier_log, ambiguous = _match_rows(df_old, df_new, tiers, shared_cols)
    rows = _plan_from_matches(df_old, df_new, old_to_new)

    dup_keys["old"] = ambiguous["old"]
    dup_keys["new"] = ambiguous["new"]
    meta["tiers"] = tier_log
    meta["tier_specs"] = ["+".join(_tier_spec(t)[0]) for t in tiers]
    meta["matched"] = len(old_to_new)
    return rows, dup_keys, meta


def _cell(df, idx, col):
    if idx is None or col not in df.columns:
        return ""
    return _norm(df.iloc[idx][col])


def _finish_sheet(ws, n_cols, freeze="D2"):
    ws.freeze_panes = freeze
    if n_cols:
        ws.auto_filter.ref = f"A1:{get_column_letter(n_cols)}{max(ws.max_row, 1)}"
    widths = {1: 14, 2: 9, 3: 9}
    for c in range(4, n_cols + 1):
        widths[c] = 22
    for c, w in widths.items():
        ws.column_dimensions[get_column_letter(c)].width = w


# ---------------------------------------------------------------------------
# Ingestion & Data Preparation
# ---------------------------------------------------------------------------


def _param_for(sheet, param):
    """Retrieve parameter for a specific sheet if passed as dict, else param."""
    if isinstance(param, dict):
        return param.get(sheet)
    return param


def _drop_ignored(df, ignore_columns):
    """Remove ignored columns (case-insensitive) before comparison."""
    if not ignore_columns:
        return df
    drop = {c.lower() for c in ignore_columns}
    keep = [c for c in df.columns if _norm(c).strip().lower() not in drop]
    return df[keep]


def _drop_blank_rows(df):
    """Drop rows that are entirely empty after string normalization."""
    if df.empty:
        return df
    mask = df.apply(lambda r: any(_norm(v).strip() != "" for v in r), axis=1)
    return df[mask].reset_index(drop=True)


def _read_excel_clean(path, sheet_name, header, skiprows, usecols, fill_na):
    """Read excel sheet with fallback if active filters cause pandas/openpyxl errors."""
    try:
        return pd.read_excel(path, sheet_name=sheet_name, dtype=str,
                             header=header, skiprows=skiprows, usecols=usecols).fillna(fill_na)
    except Exception as e:
        err_msg = str(e)
        if 'numerical or a string containing a wildcard' in err_msg or 'filter' in err_msg.lower():
            # Load with openpyxl, remove autofilter, and extract data cleanly
            wb = load_workbook(path, data_only=True)
            ws = wb[sheet_name]
            ws.auto_filter.ref = None
            data = []
            for row in ws.iter_rows(values_only=True):
                data.append(list(row))
            if not data:
                return pd.DataFrame()
            raw_df = pd.DataFrame(data)
            if skiprows:
                if isinstance(skiprows, int):
                    raw_df = raw_df.iloc[skiprows:].reset_index(drop=True)
                elif isinstance(skiprows, (list, tuple)):
                    raw_df = raw_df.drop(index=list(skiprows)).reset_index(drop=True)
            if header is not None and header < len(raw_df):
                col_names = list(raw_df.iloc[header])
                raw_df = raw_df.iloc[header + 1:].reset_index(drop=True)
                raw_df.columns = col_names
            if usecols:
                if isinstance(usecols, (list, tuple)):
                    raw_df = raw_df.iloc[:, [c for c in usecols if c < raw_df.shape[1]]]
            return raw_df.fillna(fill_na)
        raise


def _read_rich_markup_sheet(path, sheet_name, header, skiprows, usecols, fill_na, ignore_strikethrough):
    """Read sheet using openpyxl with rich text and strikethrough inspection."""
    wb = load_workbook(path, rich_text=True)
    ws = wb[sheet_name]
    sc = sc_mod.StringChecker(workbook=wb) if sc_mod else None

    data = []
    for row in ws.iter_rows():
        row_vals = []
        for cell in row:
            if sc:
                val = sc.extract_effective_text(cell.value, cell.font, ignore_strikethrough=ignore_strikethrough)
            else:
                val = _norm(cell.value)
            row_vals.append(val if val != "" else fill_na)
        data.append(row_vals)

    if not data:
        return pd.DataFrame()
    df = pd.DataFrame(data)
    if skiprows:
        if isinstance(skiprows, int):
            df = df.iloc[skiprows:].reset_index(drop=True)
        elif isinstance(skiprows, (list, tuple)):
            df = df.drop(index=list(skiprows)).reset_index(drop=True)
    if header is not None and header < len(df):
        cols = list(df.iloc[header])
        df = df.iloc[header + 1:].reset_index(drop=True)
        df.columns = cols
    if usecols and isinstance(usecols, (list, tuple)):
        df = df.iloc[:, [c for c in usecols if c < df.shape[1]]]
    return df.fillna(fill_na)


def _read_sheets(path, header_row=None, skiprows=None, usecols=None,
                 column_names=None, fill_na="", parse_rich_markup=False,
                 ignore_strikethrough=True, strip_whitespace=False):
    """Read every sheet of an .xlsx/.xlsm/.csv/.tsv into string DataFrames with overrides."""
    fill_na_str = str(fill_na) if fill_na is not None else ""

    if str(path).lower().endswith((".csv", ".txt", ".tsv")):
        sep = "\t" if str(path).lower().endswith(".tsv") else ","
        hdr = _param_for("Sheet1", header_row)
        skip = _param_for("Sheet1", skiprows)
        cols = _param_for("Sheet1", usecols)
        cnames = _param_for("Sheet1", column_names)

        hdr_idx = None if hdr is False else ((hdr - 1) if hdr is not None and hdr > 0 else 0)
        df = pd.read_csv(path, dtype=str, sep=sep, header=hdr_idx,
                         skiprows=skip, usecols=cols).fillna(fill_na_str)
        if cnames:
            df.columns = list(cnames)[:df.shape[1]]
        if strip_whitespace:
            df.columns = [_norm(c).strip() for c in df.columns]
            df = df.map(lambda v: _norm(v).strip() if isinstance(v, str) else v)
        return {"Sheet1": _drop_blank_rows(df)}

    xls = pd.ExcelFile(path)
    out = {}
    for name in xls.sheet_names:
        hdr = _param_for(name, header_row)
        skip = _param_for(name, skiprows)
        cols = _param_for(name, usecols)
        cnames = _param_for(name, column_names)

        hdr_idx = None if hdr is False else ((hdr - 1) if hdr is not None and hdr > 0 else 0)
        try:
            if parse_rich_markup and sc_mod:
                df = _read_rich_markup_sheet(path, name, hdr_idx, skip, cols, fill_na_str, ignore_strikethrough)
            else:
                df = _read_excel_clean(path, name, hdr_idx, skip, cols, fill_na_str)

            if cnames:
                df.columns = list(cnames)[:df.shape[1]]
            else:
                # Drop unnamed/blank columns
                df = df[[c for c in df.columns if not str(c).startswith("Unnamed:")]]

            if strip_whitespace:
                df.columns = [_norm(c).strip() for c in df.columns]
                df = df.map(lambda v: _norm(v).strip() if isinstance(v, str) else v)

            out[name] = _drop_blank_rows(df)
        except Exception as exc:
            # Fallback to top-row header
            try:
                df = pd.read_excel(path, sheet_name=name, dtype=str).fillna(fill_na_str)
                df = df[[c for c in df.columns if not str(c).startswith("Unnamed:")]]
                if strip_whitespace:
                    df.columns = [_norm(c).strip() for c in df.columns]
                    df = df.map(lambda v: _norm(v).strip() if isinstance(v, str) else v)
                out[name] = _drop_blank_rows(df)
            except Exception as exc2:
                print(f"Could not read sheet {name} from {path}: {exc2}")
    return out


# ---------------------------------------------------------------------------
# Output Writing: Paired Tabs, Stacked Tabs, and Summary
# ---------------------------------------------------------------------------


def _write_trace(ws, row, old_idx, new_idx, offset):
    """Record source row numbers."""
    for col, idx in ((2, old_idx), (3, new_idx)):
        cell = ws.cell(row=row, column=col,
                       value=("" if idx is None else idx + offset))
        cell.border = BORDER_ALL
        cell.alignment = Alignment(horizontal="center", vertical="top")
        cell.font = Font(color="00808080")


def _write_status(ws, row, status, fill):
    cell = ws.cell(row=row, column=1, value=status)
    cell.border = BORDER_ALL
    cell.alignment = Alignment(vertical="top")
    if status == "Added":
        cell.font = Font(bold=True, color=ADDED_COLOR)
    elif status == "Removed":
        cell.font = Font(bold=True, color=REMOVED_COLOR, strike=True)
    elif status == "Modified":
        cell.font = Font(bold=True)
    if fill:
        cell.fill = fill


def _write_sheet_pair(wb, sheet, df_old, df_new, columns, rows, dup_keys,
                      key_column, used_titles, include_unchanged, meta=None):
    """Write the (Added) and (Removed) tabs for one source sheet."""
    ws_add = wb.create_sheet(_safe_sheet_title(sheet, " (Added)", used_titles))
    ws_rem = wb.create_sheet(_safe_sheet_title(sheet, " (Removed)", used_titles))

    header = [STATUS_COLUMN, "Old Row", "New Row"] + [_norm(c) for c in columns]
    for ws in (ws_add, ws_rem):
        for col_idx, name in enumerate(header, 1):
            cell = ws.cell(row=1, column=col_idx, value=name)
            cell.font = FONT_HEADER
            cell.fill = FILL_HEADER
            cell.border = BORDER_ALL
            cell.alignment = Alignment(vertical="center", wrap_text=True)

    shared_cols = [c for c in columns if c in df_old.columns and c in df_new.columns]

    counts = {
        "sheet": sheet,
        "rows_old": len(df_old),
        "rows_new": len(df_new),
        "rows_added": 0,
        "rows_removed": 0,
        "rows_modified": 0,
        "rows_unchanged": 0,
        "cols_old": len(df_old.columns),
        "cols_new": len(df_new.columns),
        "cols_added": [_norm(c) for c in df_new.columns if c not in df_old.columns],
        "cols_removed": [_norm(c) for c in df_old.columns if c not in df_new.columns],
        "cells_compared": 0,
        "cells_changed": 0,
        "cells_changed_shared": 0,
        "rows_modified_shared": 0,
        "rows_whitespace_only": 0,
        "rows_case_only": 0,
        "rows_reordered": 0,
        "dup_keys_old": dup_keys["old"],
        "dup_keys_new": dup_keys["new"],
        "alignment": (meta or {}).get("mode", "positional"),
        "match_tiers": (meta or {}).get("tiers", []),
        "rows_matched": (meta or {}).get("matched", 0),
    }

    header_offset = (meta or {}).get("source_row_offset", 2)
    r_add, r_rem = 2, 2
    for row in rows:
        old_idx, new_idx = row["old_idx"], row["new_idx"]
        old_vals = [_cell(df_old, old_idx, c) for c in columns]
        new_vals = [_cell(df_new, new_idx, c) for c in columns]

        if new_idx is not None and old_idx is None:
            status = "Added"
            counts["rows_added"] += 1
        elif old_idx is not None and new_idx is None:
            status = "Removed"
            counts["rows_removed"] += 1
        elif old_vals != new_vals:
            status = "Modified"
            counts["rows_modified"] += 1
        else:
            status = "Unchanged"
            counts["rows_unchanged"] += 1

        counts["cells_compared"] += len(columns)
        counts["cells_changed"] += sum(1 for o, n in zip(old_vals, new_vals) if o != n)

        if old_idx is not None and new_idx is not None:
            if old_idx != new_idx:
                counts["rows_reordered"] += 1

        if old_idx is not None and new_idx is not None:
            shared_old = [_cell(df_old, old_idx, c) for c in shared_cols]
            shared_new = [_cell(df_new, new_idx, c) for c in shared_cols]
            changed_shared = [(o, n) for o, n in zip(shared_old, shared_new) if o != n]
            counts["cells_changed_shared"] += len(changed_shared)
            if changed_shared:
                counts["rows_modified_shared"] += 1
                if all(o.strip() == n.strip() for o, n in changed_shared):
                    counts["rows_whitespace_only"] += 1
                elif all(o.strip().lower() == n.strip().lower() for o, n in changed_shared):
                    counts["rows_case_only"] += 1

        row["_status"] = status
        row["_old_vals"] = old_vals
        row["_new_vals"] = new_vals

        if status == "Unchanged" and not include_unchanged:
            continue

        # ADDED tab (new values, red bold insertions)
        if new_idx is not None:
            _write_status(ws_add, r_add, status,
                          FILL_ROW_ADDED if status == "Added" else None)
            _write_trace(ws_add, r_add, old_idx, new_idx, header_offset)
            for col_idx, (old_v, new_v) in enumerate(zip(old_vals, new_vals), 4):
                cell = ws_add.cell(row=r_add, column=col_idx)
                if status == "Added":
                    cell.value = new_v
                    cell.fill = FILL_ROW_ADDED
                else:
                    cell.value = diff_strings_rich_text(old_v, new_v)
                    if old_v != new_v:
                        cell.fill = FILL_CHANGED
                cell.border = BORDER_ALL
                cell.alignment = Alignment(vertical="top", wrap_text=True)
            r_add += 1

        # REMOVED tab (old values, green strikethrough)
        if old_idx is not None:
            _write_status(ws_rem, r_rem, status,
                          FILL_ROW_REMOVED if status == "Removed" else None)
            _write_trace(ws_rem, r_rem, old_idx, new_idx, header_offset)
            for col_idx, (old_v, new_v) in enumerate(zip(old_vals, new_vals), 4):
                cell = ws_rem.cell(row=r_rem, column=col_idx)
                if status == "Removed":
                    cell.value = strikethrough_rich_text(old_v)
                    cell.fill = FILL_ROW_REMOVED
                else:
                    cell.value = diff_strings_removed_rich_text(old_v, new_v)
                    if old_v != new_v:
                        cell.fill = FILL_CHANGED
                cell.border = BORDER_ALL
                cell.alignment = Alignment(vertical="top", wrap_text=True)
            r_rem += 1

    for ws in (ws_add, ws_rem):
        _finish_sheet(ws, len(header))

    pct = (counts["cells_changed"] / counts["cells_compared"] * 100) if counts["cells_compared"] else 0.0
    counts["pct_cells_changed"] = round(pct, 2)
    counts["pct_cells_changed_shared"] = round(
        (counts["cells_changed_shared"] / counts["cells_compared"] * 100)
        if counts["cells_compared"] else 0.0, 2)

    reasons = []
    if counts.get("rows_reordered"):
        reasons.append(f"{counts['rows_reordered']} matched row(s) changed position "
                       "(matched by identity, not row order)")
    if counts["rows_added"]:
        reasons.append(f"{counts['rows_added']} row(s) added")
    if counts["rows_removed"]:
        reasons.append(f"{counts['rows_removed']} row(s) removed")
    if counts["cols_added"]:
        reasons.append("columns added: " + ", ".join(counts["cols_added"]))
    if counts["cols_removed"]:
        reasons.append("columns removed: " + ", ".join(counts["cols_removed"]))
    if counts["rows_whitespace_only"]:
        reasons.append(f"{counts['rows_whitespace_only']} row(s) differ by whitespace only")
    if counts["rows_case_only"]:
        reasons.append(f"{counts['rows_case_only']} row(s) differ by letter case only")
    if counts["dup_keys_old"] or counts["dup_keys_new"]:
        dups = sorted(set(counts["dup_keys_old"]) | set(counts["dup_keys_new"]))
        reasons.append(
            f"{len(dups)} ambiguous/duplicate key value(s) resolved by similarity: "
            + ", ".join(dups[:5]) + ("; ..." if len(dups) > 5 else ""))
    if key_column and (key_column not in df_old.columns or key_column not in df_new.columns):
        reasons.append(f"key column '{key_column}' missing -- fell back to positional alignment")

    counts["integrity"] = "PASS" if not reasons else "FAIL"
    counts["notes"] = "; ".join(reasons) if reasons else "Structure identical"
    return counts


def _write_stacked_sheet(wb, sheet, df_old, df_new, columns, rows, group_labels,
                         used_titles, include_unchanged=True, meta=None,
                         target_ws=None, start_row=None):
    """
    Write a stacked submittal layout (generalized from WriteToExcel.py).
    Creates sections for Group 0 and Group 1 with blue headers and rich text diffs.
    """
    if target_ws is not None:
        ws = target_ws
    else:
        title = _safe_sheet_title(sheet, " (Submittal)", used_titles)
        ws = wb.create_sheet(title)

    curr_row = start_row if start_row is not None else 1

    # Table Header
    header = [STATUS_COLUMN, "Old Row", "New Row"] + [_norm(c) for c in columns]
    n_cols = len(header)
    for col_idx, name in enumerate(header, 1):
        cell = ws.cell(row=curr_row, column=col_idx, value=name)
        cell.font = FONT_HEADER
        cell.fill = FILL_HEADER
        cell.border = BORDER_ALL
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    curr_row += 1

    header_offset = (meta or {}).get("source_row_offset", 2)
    label_old = group_labels[0] if len(group_labels) > 0 else "Old"
    label_new = group_labels[1] if len(group_labels) > 1 else "New"

    # Section 1: Baseline / Old Submittal
    banner_1 = ws.cell(row=curr_row, column=1, value=f"{label_old} Submittal")
    banner_1.font = FONT_BANNER
    banner_1.fill = FILL_BANNER
    ws.merge_cells(start_row=curr_row, start_column=1, end_row=curr_row, end_column=n_cols)
    curr_row += 1

    for row in rows:
        old_idx = row["old_idx"]
        status = row.get("_status", "Unchanged")
        if old_idx is None:
            continue
        if status == "Unchanged" and not include_unchanged:
            continue

        _write_status(ws, curr_row, status, FILL_ROW_REMOVED if status == "Removed" else None)
        _write_trace(ws, curr_row, old_idx, row.get("new_idx"), header_offset)
        old_vals = row.get("_old_vals", [_cell(df_old, old_idx, c) for c in columns])
        new_vals = row.get("_new_vals", [_cell(df_new, row.get("new_idx"), c) for c in columns])

        for col_idx, (old_v, new_v) in enumerate(zip(old_vals, new_vals), 4):
            cell = ws.cell(row=curr_row, column=col_idx)
            if status == "Removed":
                cell.value = strikethrough_rich_text(old_v)
                cell.fill = FILL_ROW_REMOVED
            else:
                cell.value = diff_strings_removed_rich_text(old_v, new_v)
                if old_v != new_v:
                    cell.fill = FILL_CHANGED
            cell.border = BORDER_ALL
            cell.alignment = Alignment(vertical="top", wrap_text=True)
        curr_row += 1

    curr_row += 1

    # Section 2: Comparison / New Submittal
    banner_2 = ws.cell(row=curr_row, column=1, value=f"{label_new} Submittal")
    banner_2.font = FONT_BANNER
    banner_2.fill = FILL_BANNER
    ws.merge_cells(start_row=curr_row, start_column=1, end_row=curr_row, end_column=n_cols)
    curr_row += 1

    for row in rows:
        new_idx = row["new_idx"]
        status = row.get("_status", "Unchanged")
        if new_idx is None:
            continue
        if status == "Unchanged" and not include_unchanged:
            continue

        _write_status(ws, curr_row, status, FILL_ROW_ADDED if status == "Added" else None)
        _write_trace(ws, curr_row, row.get("old_idx"), new_idx, header_offset)
        old_vals = row.get("_old_vals", [_cell(df_old, row.get("old_idx"), c) for c in columns])
        new_vals = row.get("_new_vals", [_cell(df_new, new_idx, c) for c in columns])

        for col_idx, (old_v, new_v) in enumerate(zip(old_vals, new_vals), 4):
            cell = ws.cell(row=curr_row, column=col_idx)
            if status == "Added":
                cell.value = new_v
                cell.fill = FILL_ROW_ADDED
            else:
                cell.value = diff_strings_rich_text(old_v, new_v)
                if old_v != new_v:
                    cell.fill = FILL_CHANGED
            cell.border = BORDER_ALL
            cell.alignment = Alignment(vertical="top", wrap_text=True)
        curr_row += 1

    _finish_sheet(ws, n_cols)
    return ws


def _write_summary(ws, stats, group_labels=("Old", "New")):
    """Per-sheet change counts, % changed, and PASS/FAIL integrity check."""
    ws["A1"] = "Excel Diff & Compare -- Summary"
    ws["A1"].font = FONT_TITLE
    ws["A1"].fill = FILL_TITLE
    ws.merge_cells("A1:R1")
    ws.row_dimensions[1].height = 22

    label_old = group_labels[0] if len(group_labels) > 0 else "Old"
    label_new = group_labels[1] if len(group_labels) > 1 else "New"

    meta = [
        (f"{label_old} file (baseline)", stats["old_file"]),
        (f"{label_new} file (comparison)", stats["new_file"]),
        ("Row alignment", stats.get("alignment") or
            (f"key column '{stats['key_column']}'" if stats["key_column"] else "positional")),
        ("Match key tiers", ", ".join(stats["match_keys"])
            if isinstance(stats.get("match_keys"), list) else str(stats.get("match_keys"))),
        ("Ignored columns", ", ".join(stats.get("ignore_columns") or []) or "none"),
        ("Generated", stats["generated"]),
        ("Sheets compared", len(stats["sheets"])),
    ]
    r = 3
    for label, value in meta:
        ws.cell(row=r, column=1, value=label).font = FONT_HEADER
        ws.cell(row=r, column=2, value=value)
        r += 1

    r += 1
    headers = ["Sheet", f"Rows {label_old}", f"Rows {label_new}", "Row Δ",
               "Rows Added", "Rows Removed", "Rows Modified",
               "Rows Modified (data only)", "Rows Matched", "Rows Moved",
               f"Cols {label_old}", f"Cols {label_new}",
               "Cells Compared", "Cells Changed", "Cells Changed (data only)",
               "% Cells Changed", "Integrity", "Notes"]
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=r, column=c, value=h)
        cell.font = FONT_HEADER
        cell.fill = FILL_HEADER
        cell.border = BORDER_ALL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    header_row = r
    r += 1

    totals = {k: 0 for k in ("rows_added", "rows_removed", "rows_modified",
                             "rows_modified_shared", "cells_compared",
                             "cells_changed", "cells_changed_shared")}
    for s in stats["sheets"]:
        values = [
            s["sheet"], s["rows_old"], s["rows_new"], s["rows_new"] - s["rows_old"],
            s["rows_added"], s["rows_removed"], s["rows_modified"],
            s["rows_modified_shared"], s.get("rows_matched", 0),
            s.get("rows_reordered", 0), s["cols_old"], s["cols_new"],
            s["cells_compared"], s["cells_changed"], s["cells_changed_shared"],
            s["pct_cells_changed"] / 100.0, s["integrity"], s["notes"],
        ]
        for c, v in enumerate(values, 1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.border = BORDER_ALL
            if c == COL_PCT:
                cell.number_format = "0.00%"
            if c == COL_INTEGRITY:
                cell.fill = FILL_PASS if v == "PASS" else FILL_FAIL
                cell.font = FONT_PASS if v == "PASS" else FONT_FAIL
                cell.alignment = Alignment(horizontal="center")
            if c == COL_NOTES:
                cell.alignment = Alignment(wrap_text=True, vertical="top")
        for k in totals:
            totals[k] += s[k]
        r += 1

    if stats["sheets"]:
        ws.cell(row=r, column=1, value="TOTAL").font = FONT_HEADER
        pct = (totals["cells_changed"] / totals["cells_compared"]) if totals["cells_compared"] else 0.0
        for c, v in ((5, totals["rows_added"]), (6, totals["rows_removed"]),
                     (7, totals["rows_modified"]), (8, totals["rows_modified_shared"]),
                     (13, totals["cells_compared"]), (14, totals["cells_changed"]),
                     (15, totals["cells_changed_shared"]), (COL_PCT, pct)):
            cell = ws.cell(row=r, column=c, value=v)
            cell.font = FONT_HEADER
            if c == COL_PCT:
                cell.number_format = "0.00%"
        for c in range(1, COL_NOTES + 1):
            ws.cell(row=r, column=c).border = BORDER_ALL
        r += 2

    structural = [
        (f"Sheets only in {label_old} file", ", ".join(stats["sheets_only_in_old"]) or "none"),
        (f"Sheets only in {label_new} file", ", ".join(stats["sheets_only_in_new"]) or "none"),
        ("Sheets skipped", ", ".join(stats["skipped"]) or "none"),
    ]
    for label, value in structural:
        ws.cell(row=r, column=1, value=label).font = FONT_HEADER
        ws.cell(row=r, column=2, value=value)
        r += 1

    r += 1
    ws.cell(row=r, column=1, value="Legend").font = FONT_HEADER
    r += 1
    legend = [
        ("Red bold text", f"Characters inserted in the {label_new} file.", ADDED_COLOR, False),
        ("Green strikethrough", f"Characters deleted from the {label_old} file.", REMOVED_COLOR, True),
        ("Light orange fill", "Cell whose value changed between the two files.", None, False),
        ("Light green fill", f"Entire row exists in only one file (Added / Removed).", None, False),
        ("Rows Modified", "Includes rows changed only because a column was added or removed.", None, False),
        ("Rows Modified (data only)", "Rows with real edits inside columns common to both files.", None, False),
        ("Rows Matched", "Old rows paired to a new row by identity, regardless of row order.", None, False),
        ("Rows Moved", "Matched rows whose position differs between the two files.", None, False),
        ("Old Row / New Row", "Source row number in each workbook, so a moved row can be located.", None, False),
    ]
    for label, desc, color, strike in legend:
        cell = ws.cell(row=r, column=1, value=label)
        cell.font = Font(bold=True, color=color, strike=strike) if color else Font(bold=True)
        ws.cell(row=r, column=2, value=desc)
        r += 1

    widths = {1: 26, 2: 34, COL_NOTES: 60}
    for c in range(3, COL_NOTES):
        widths.setdefault(c, 13)
    for c, w in widths.items():
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = ws.cell(row=header_row + 1, column=1).coordinate


# ---------------------------------------------------------------------------
# Main Entry Point: generate_diff_report
# ---------------------------------------------------------------------------


def generate_diff_report(old_file_path, new_file_path, output_file_path,
                         key_column=None, sheet_name=0, include_unchanged=True,
                         match_keys="auto", ignore_columns=None, header_row=None,
                         skiprows=None, usecols=None, column_names=None,
                         fill_na="", group_labels=("Old", "New"),
                         report_layout="paired", template_path=None,
                         template_start_row=None, template_sheet=None, parse_rich_markup=False,
                         ignore_strikethrough=True, strip_whitespace=False):
    """
    Compare two Excel/CSV files and write a styled diff workbook.

    Parameters
    ----------
    old_file_path : str
        Path to baseline dataset (.xlsx, .xlsm, .csv, .tsv).
    new_file_path : str
        Path to comparison dataset (.xlsx, .xlsm, .csv, .tsv).
    output_file_path : str
        Path where output .xlsx diff report will be saved.
    key_column : str, optional
        Single join column (legacy). Overrides match_keys when present in both.
    sheet_name : str, int, list, or None, default 0
        None -> every sheet common to both files.
        list -> list of sheet names to compare.
        str/int -> single sheet name or 0-based index.
    include_unchanged : bool, default True
        Keep unchanged rows on detail tabs for context.
    match_keys : str or list of tiers, default 'auto'
        'auto' -> auto-derived identity scoring.
        'positional' -> row-by-row matching.
        list of tiers -> tiered identity matching with optional validators.
    ignore_columns : list of str, optional
        Column names to exclude from comparison and report.
    header_row : int or dict, optional
        1-based row holding headers. None -> row 1. False -> headerless.
    skiprows : int, list, or dict, optional
        Rows to skip before reading header/data.
    usecols : list or dict, optional
        Subset of column indices or names to read.
    column_names : list or dict, optional
        Explicit column names to assign.
    fill_na : str, default ""
        Value to fill empty/missing cells with (e.g. '-' or '').
    group_labels : tuple of (str, str), default ('Old', 'New')
        Labels for the two datasets (e.g. ('Prepared', 'Revised')).
    report_layout : {'paired', 'stacked', 'both'}, default 'paired'
        'paired'  -> Summary + (Added) + (Removed) tabs.
        'stacked' -> Summary + (Submittal) with stacked sections.
        'both'    -> Summary + (Added) + (Removed) + (Submittal).
    template_path : str, optional
        Path to an existing template workbook to inject comparison into.
    template_start_row : int, optional
        Starting row in template sheet to begin writing report.
    parse_rich_markup : bool, default False
        Inspect cell-level rich text and colors via openpyxl.
    ignore_strikethrough : bool, default True
        When parse_rich_markup=True, omit struck-through text deletions.
    strip_whitespace : bool, default True
        Strip leading/trailing whitespace from string cells.

    Returns
    -------
    dict
        Detailed stats dict of differences, changes, and integrity checks.
    """
    old_sheets = _read_sheets(
        old_file_path, header_row=header_row, skiprows=skiprows,
        usecols=usecols, column_names=column_names, fill_na=fill_na,
        parse_rich_markup=parse_rich_markup, ignore_strikethrough=ignore_strikethrough,
        strip_whitespace=strip_whitespace
    )
    new_sheets = _read_sheets(
        new_file_path, header_row=header_row, skiprows=skiprows,
        usecols=usecols, column_names=column_names, fill_na=fill_na,
        parse_rich_markup=parse_rich_markup, ignore_strikethrough=ignore_strikethrough,
        strip_whitespace=strip_whitespace
    )
    ignore_columns = [_norm(c).strip() for c in (ignore_columns or [])]
    common = [s for s in new_sheets if s in old_sheets]

    if sheet_name is None:
        sheets_to_process = common
    elif isinstance(sheet_name, list):
        sheets_to_process = [s for s in sheet_name]
    elif isinstance(sheet_name, int):
        names = list(new_sheets.keys())
        sheets_to_process = [names[sheet_name]] if sheet_name < len(names) else []
    else:
        sheets_to_process = [sheet_name]

    if template_path and os.path.exists(template_path):
        wb = load_workbook(template_path)
        if "Summary" in wb.sheetnames:
            summary_ws = wb["Summary"]
        else:
            summary_ws = wb.create_sheet("Summary", 0)
    else:
        wb = Workbook()
        wb.remove(wb.active)
        summary_ws = wb.create_sheet("Summary")

    used_titles = set(wb.sheetnames)
    stats = {
        "old_file": os.path.basename(old_file_path),
        "new_file": os.path.basename(new_file_path),
        "group_labels": list(group_labels),
        "key_column": key_column,
        "match_keys": ("positional" if match_keys in (None, "positional")
                       else ("auto" if match_keys == "auto"
                             else [" + ".join(_tier_spec(t)[0]) for t in match_keys])),
        "ignore_columns": list(ignore_columns),
        "alignment": "",
        "generated": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "sheets": [],
        "sheets_only_in_old": [s for s in old_sheets if s not in new_sheets],
        "sheets_only_in_new": [s for s in new_sheets if s not in old_sheets],
        "skipped": [],
    }

    for sheet in sheets_to_process:
        if sheet not in old_sheets or sheet not in new_sheets:
            stats["skipped"].append(
                f"{sheet} (missing from {'old' if sheet not in old_sheets else 'new'} file)"
            )
            continue

        df_old = _drop_ignored(old_sheets[sheet], ignore_columns)
        df_new = _drop_ignored(new_sheets[sheet], ignore_columns)
        columns = _union_columns(df_old, df_new)
        shared_cols = [c for c in columns if c in df_old.columns and c in df_new.columns]
        rows, dup_keys, meta = _align(df_old, df_new, key_column,
                                      match_keys=match_keys, shared_cols=shared_cols)
        stats["alignment"] = meta["mode"]

        # Write paired tabs (Added / Removed)
        sheet_stats = None
        if report_layout in ("paired", "both"):
            sheet_stats = _write_sheet_pair(
                wb, sheet, df_old, df_new, columns, rows, dup_keys,
                key_column, used_titles, include_unchanged, meta,
            )

        # Write stacked submittal tab
        if report_layout in ("stacked", "both"):
            target_ws = None
            if template_path:
                t_name = template_sheet if template_sheet else sheet
                if isinstance(t_name, int) and t_name < len(wb.worksheets):
                    target_ws = wb.worksheets[t_name]
                elif isinstance(t_name, str) and t_name in wb.sheetnames:
                    target_ws = wb[t_name]
            _write_stacked_sheet(
                wb, sheet, df_old, df_new, columns, rows, group_labels,
                used_titles, include_unchanged=include_unchanged, meta=meta,
                target_ws=target_ws, start_row=template_start_row
            )

        # In stacked-only mode, compute sheet_stats if not already computed
        if sheet_stats is None:
            # Run quick counts
            sheet_stats = _compute_counts_only(
                sheet, df_old, df_new, columns, rows, dup_keys, shared_cols, key_column, meta
            )

        stats["sheets"].append(sheet_stats)

    if not stats["sheets"]:
        ws = wb.create_sheet("No Matching Sheets")
        ws["A1"] = "No sheets were common to both workbooks."
        ws["A1"].font = FONT_HEADER

    _write_summary(summary_ws, stats, group_labels=group_labels)
    os.makedirs(os.path.dirname(os.path.abspath(output_file_path)), exist_ok=True)
    wb.save(output_file_path)
    return stats


def _compute_counts_only(sheet, df_old, df_new, columns, rows, dup_keys, shared_cols, key_column, meta):
    """Compute statistics when paired tabs are skipped."""
    counts = {
        "sheet": sheet, "rows_old": len(df_old), "rows_new": len(df_new),
        "rows_added": 0, "rows_removed": 0, "rows_modified": 0, "rows_unchanged": 0,
        "cols_old": len(df_old.columns), "cols_new": len(df_new.columns),
        "cols_added": [_norm(c) for c in df_new.columns if c not in df_old.columns],
        "cols_removed": [_norm(c) for c in df_old.columns if c not in df_new.columns],
        "cells_compared": 0, "cells_changed": 0, "cells_changed_shared": 0,
        "rows_modified_shared": 0, "rows_whitespace_only": 0, "rows_case_only": 0,
        "rows_reordered": 0, "dup_keys_old": dup_keys["old"], "dup_keys_new": dup_keys["new"],
        "alignment": (meta or {}).get("mode", "positional"),
        "match_tiers": (meta or {}).get("tiers", []),
        "rows_matched": (meta or {}).get("matched", 0),
    }
    for row in rows:
        old_idx, new_idx = row["old_idx"], row["new_idx"]
        old_vals = [_cell(df_old, old_idx, c) for c in columns]
        new_vals = [_cell(df_new, new_idx, c) for c in columns]

        if new_idx is not None and old_idx is None:
            status = "Added"
            counts["rows_added"] += 1
        elif old_idx is not None and new_idx is None:
            status = "Removed"
            counts["rows_removed"] += 1
        elif old_vals != new_vals:
            status = "Modified"
            counts["rows_modified"] += 1
        else:
            status = "Unchanged"
            counts["rows_unchanged"] += 1

        counts["cells_compared"] += len(columns)
        counts["cells_changed"] += sum(1 for o, n in zip(old_vals, new_vals) if o != n)
        if old_idx is not None and new_idx is not None:
            if old_idx != new_idx:
                counts["rows_reordered"] += 1
            shared_old = [_cell(df_old, old_idx, c) for c in shared_cols]
            shared_new = [_cell(df_new, new_idx, c) for c in shared_cols]
            changed_shared = [(o, n) for o, n in zip(shared_old, shared_new) if o != n]
            counts["cells_changed_shared"] += len(changed_shared)
            if changed_shared:
                counts["rows_modified_shared"] += 1
                if all(o.strip() == n.strip() for o, n in changed_shared):
                    counts["rows_whitespace_only"] += 1
                elif all(o.strip().lower() == n.strip().lower() for o, n in changed_shared):
                    counts["rows_case_only"] += 1
        row["_status"] = status
        row["_old_vals"] = old_vals
        row["_new_vals"] = new_vals

    pct = (counts["cells_changed"] / counts["cells_compared"] * 100) if counts["cells_compared"] else 0.0
    counts["pct_cells_changed"] = round(pct, 2)
    counts["pct_cells_changed_shared"] = round(
        (counts["cells_changed_shared"] / counts["cells_compared"] * 100)
        if counts["cells_compared"] else 0.0, 2)

    reasons = []
    if counts.get("rows_reordered"):
        reasons.append(f"{counts['rows_reordered']} matched row(s) changed position")
    if counts["rows_added"]:
        reasons.append(f"{counts['rows_added']} row(s) added")
    if counts["rows_removed"]:
        reasons.append(f"{counts['rows_removed']} row(s) removed")
    if counts["cols_added"]:
        reasons.append("columns added: " + ", ".join(counts["cols_added"]))
    if counts["cols_removed"]:
        reasons.append("columns removed: " + ", ".join(counts["cols_removed"]))
    if counts["rows_whitespace_only"]:
        reasons.append(f"{counts['rows_whitespace_only']} row(s) differ by whitespace only")
    if counts["rows_case_only"]:
        reasons.append(f"{counts['rows_case_only']} row(s) differ by letter case only")
    if counts["dup_keys_old"] or counts["dup_keys_new"]:
        dups = sorted(set(counts["dup_keys_old"]) | set(counts["dup_keys_new"]))
        reasons.append(f"{len(dups)} ambiguous key(s)")
    if key_column and (key_column not in df_old.columns or key_column not in df_new.columns):
        reasons.append(f"key column '{key_column}' missing")

    counts["integrity"] = "PASS" if not reasons else "FAIL"
    counts["notes"] = "; ".join(reasons) if reasons else "Structure identical"
    return counts


# ---------------------------------------------------------------------------
# Batch Directory Comparison (replacing GetFiles.py)
# ---------------------------------------------------------------------------


def find_file_pairs(dir_old, dir_new, name_map=None, file_types=None, name_patterns=None):
    """
    Find corresponding pairs of files across two directories.

    Parameters
    ----------
    dir_old : str
        Directory holding baseline files.
    dir_new : str
        Directory holding comparison files.
    name_map : dict, optional
        Dictionary mapping old stems to new stem(s), or vice versa.
        Can also be nested dict: {group0: {...}, group1: {...}}
    file_types : list of str, optional
        Allowed file extensions (default: ['.xlsx', '.xlsm', '.csv', '.tsv']).
    name_patterns : list of str or re.Pattern, optional
        Regex patterns to strip from stems to match files with revision/date tags.

    Returns
    -------
    tuple
        (pairs, unmatched_old, unmatched_new)
        pairs: list of (old_file_path, new_file_path, stem_name)
    """
    types = [t.lower() for t in (file_types or ['.xlsx', '.xlsm', '.csv', '.tsv'])]

    def _get_files(d):
        res = {}
        if not os.path.exists(d):
            return res
        for fname in os.listdir(d):
            ext = os.path.splitext(fname)[1].lower()
            if ext in types:
                stem = os.path.splitext(fname)[0]
                res[stem] = os.path.join(d, fname)
        return res

    files_old = _get_files(dir_old)
    files_new = _get_files(dir_new)

    def _normalize_stem(s):
        cur = s
        if name_patterns:
            for pat in name_patterns:
                cur = re.sub(pat, '', cur)
        return cur.strip(' _-')

    norm_to_new = {}
    for stem, path in files_new.items():
        norm = _normalize_stem(stem)
        norm_to_new.setdefault(norm, []).append((stem, path))

    matched_pairs = []
    used_old = set()
    used_new = set()

    # Process explicit name_map first
    if name_map:
        for old_k, targets in name_map.items():
            if isinstance(targets, str):
                targets = [targets]
            if old_k in files_old:
                for tgt in targets:
                    if tgt in files_new:
                        matched_pairs.append((files_old[old_k], files_new[tgt], old_k))
                        used_old.add(old_k)
                        used_new.add(tgt)

    # Match identical stems
    for stem in files_old:
        if stem in used_old:
            continue
        if stem in files_new and stem not in used_new:
            matched_pairs.append((files_old[stem], files_new[stem], stem))
            used_old.add(stem)
            used_new.add(stem)

    # Match normalized stems
    for stem in files_old:
        if stem in used_old:
            continue
        norm = _normalize_stem(stem)
        if norm in norm_to_new:
            for new_stem, new_path in norm_to_new[norm]:
                if new_stem not in used_new:
                    matched_pairs.append((files_old[stem], new_path, stem))
                    used_old.add(stem)
                    used_new.add(new_stem)
                    break

    unmatched_old = [path for stem, path in files_old.items() if stem not in used_old]
    unmatched_new = [path for stem, path in files_new.items() if stem not in used_new]

    return matched_pairs, unmatched_old, unmatched_new


def compare_directories(dir_old, dir_new, output_dir,
                        group_labels=("Old", "New"),
                        name_map=None, file_types=None, name_patterns=None,
                        **diff_kwargs):
    """
    Batch compare all corresponding files in two directories and write diff reports.

    Writes individual diff reports (Compared_<stem>.xlsx) in output_dir,
    and writes an aggregated Batch_Summary.xlsx.

    Returns batch statistics dictionary.
    """
    os.makedirs(output_dir, exist_ok=True)
    pairs, unmatched_old, unmatched_new = find_file_pairs(
        dir_old, dir_new, name_map=name_map,
        file_types=file_types, name_patterns=name_patterns
    )

    batch_stats = {
        "dir_old": dir_old,
        "dir_new": dir_new,
        "output_dir": output_dir,
        "group_labels": list(group_labels),
        "files_compared": [],
        "unmatched_old": unmatched_old,
        "unmatched_new": unmatched_new,
        "generated": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "totals": {
            "files_paired": len(pairs),
            "sheets_compared": 0,
            "cells_compared": 0,
            "cells_changed": 0,
            "cells_changed_shared": 0,
            "rows_added": 0,
            "rows_removed": 0,
            "rows_modified": 0,
        }
    }

    for old_path, new_path, label in pairs:
        out_filename = f"Compared_{label}.xlsx"
        out_path = os.path.join(output_dir, out_filename)

        file_stats = generate_diff_report(
            old_file_path=old_path,
            new_file_path=new_path,
            output_file_path=out_path,
            group_labels=group_labels,
            **diff_kwargs
        )
        batch_stats["files_compared"].append({
            "label": label,
            "old_file": os.path.basename(old_path),
            "new_file": os.path.basename(new_path),
            "report_path": out_path,
            "stats": file_stats
        })

        # Accumulate totals
        for s in file_stats.get("sheets", []):
            batch_stats["totals"]["sheets_compared"] += 1
            for k in ("cells_compared", "cells_changed", "cells_changed_shared",
                      "rows_added", "rows_removed", "rows_modified"):
                batch_stats["totals"][k] += s.get(k, 0)

    # Write aggregated Batch_Summary.xlsx
    summary_path = os.path.join(output_dir, "Batch_Summary.xlsx")
    _write_batch_summary_workbook(summary_path, batch_stats)
    batch_stats["batch_summary_path"] = summary_path
    return batch_stats


def _write_batch_summary_workbook(summary_path, batch_stats):
    """Write overall batch comparison summary workbook."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Batch Summary"

    ws["A1"] = "Excel Diff & Compare -- Batch Summary Dashboard"
    ws["A1"].font = FONT_TITLE
    ws["A1"].fill = FILL_TITLE
    ws.merge_cells("A1:K1")
    ws.row_dimensions[1].height = 24

    labels = batch_stats.get("group_labels", ["Old", "New"])
    meta = [
        (f"Directory {labels[0]} (baseline)", batch_stats["dir_old"]),
        (f"Directory {labels[1]} (comparison)", batch_stats["dir_new"]),
        ("Files compared", len(batch_stats["files_compared"])),
        ("Unmatched baseline files", len(batch_stats["unmatched_old"])),
        ("Unmatched comparison files", len(batch_stats["unmatched_new"])),
        ("Generated", batch_stats["generated"]),
    ]
    r = 3
    for label, val in meta:
        ws.cell(row=r, column=1, value=label).font = FONT_HEADER
        ws.cell(row=r, column=2, value=str(val))
        r += 1

    r += 1
    headers = ["Item", "Baseline File", "Comparison File", "Sheets",
               "Rows Added", "Rows Removed", "Rows Modified",
               "Cells Compared", "Cells Changed", "% Cells Changed", "Integrity"]
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=r, column=c, value=h)
        cell.font = FONT_HEADER
        cell.fill = FILL_HEADER
        cell.border = BORDER_ALL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    header_row = r
    r += 1

    for item in batch_stats["files_compared"]:
        st = item["stats"]
        sheets = st.get("sheets", [])
        r_add = sum(s.get("rows_added", 0) for s in sheets)
        r_rem = sum(s.get("rows_removed", 0) for s in sheets)
        r_mod = sum(s.get("rows_modified", 0) for s in sheets)
        c_comp = sum(s.get("cells_compared", 0) for s in sheets)
        c_chg = sum(s.get("cells_changed", 0) for s in sheets)
        pct = (c_chg / c_comp) if c_comp else 0.0
        integrity = "PASS" if all(s.get("integrity") == "PASS" for s in sheets) else "FAIL"

        vals = [
            item["label"], item["old_file"], item["new_file"], len(sheets),
            r_add, r_rem, r_mod, c_comp, c_chg, pct, integrity
        ]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.border = BORDER_ALL
            if c == 10:
                cell.number_format = "0.00%"
            if c == 11:
                cell.fill = FILL_PASS if v == "PASS" else FILL_FAIL
                cell.font = FONT_PASS if v == "PASS" else FONT_FAIL
                cell.alignment = Alignment(horizontal="center")
        r += 1

    # Totals row
    tot = batch_stats["totals"]
    pct_tot = (tot["cells_changed"] / tot["cells_compared"]) if tot["cells_compared"] else 0.0
    ws.cell(row=r, column=1, value="TOTAL").font = FONT_HEADER
    for c, v in ((4, tot["sheets_compared"]), (5, tot["rows_added"]), (6, tot["rows_removed"]),
                 (7, tot["rows_modified"]), (8, tot["cells_compared"]),
                 (9, tot["cells_changed"]), (10, pct_tot)):
        cell = ws.cell(row=r, column=c, value=v)
        cell.font = FONT_HEADER
        if c == 10:
            cell.number_format = "0.00%"
    for c in range(1, len(headers) + 1):
        ws.cell(row=r, column=c).border = BORDER_ALL
    r += 2

    if batch_stats["unmatched_old"] or batch_stats["unmatched_new"]:
        ws.cell(row=r, column=1, value="Unmatched Files").font = FONT_HEADER
        r += 1
        for p in batch_stats["unmatched_old"]:
            ws.cell(row=r, column=1, value=f"Only in {labels[0]}:")
            ws.cell(row=r, column=2, value=os.path.basename(p))
            r += 1
        for p in batch_stats["unmatched_new"]:
            ws.cell(row=r, column=1, value=f"Only in {labels[1]}:")
            ws.cell(row=r, column=2, value=os.path.basename(p))
            r += 1

    widths = {1: 25, 2: 30, 3: 30, 4: 10, 5: 12, 6: 12, 7: 14, 8: 15, 9: 14, 10: 16, 11: 12}
    for c, w in widths.items():
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = ws.cell(row=header_row + 1, column=1).coordinate

    wb.save(summary_path)


if __name__ == "__main__":
    print("compare_engine.py -- Unified Excel Diff & Compare Engine.")

