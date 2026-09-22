"""Round-trip verification for the unified excel-diff engine."""
import os, sys, shutil
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import pandas as pd
import openpyxl
from openpyxl.cell.text import InlineFont
from openpyxl.cell.rich_text import CellRichText, TextBlock
from compare_engine import (generate_diff_report, compare_directories,
                            find_file_pairs, diff_strings_rich_text,
                            diff_strings_removed_rich_text)
from string_checker import StringChecker

TMP = "/tmp/exdiff_test"
FAILS = []


def check(name, cond, detail=""):
    print(("  PASS  " if cond else "  FAIL  ") + name + (f" :: {detail}" if detail and not cond else ""))
    if not cond:
        FAILS.append(name)


def build():
    shutil.rmtree(TMP, ignore_errors=True); os.makedirs(TMP)
    old = pd.DataFrame({"ID": ["A1","A2","A3","A4"],
                        "Equipment": ["Breaker 52-1","Relay SEL-351","Xfmr T1","Cable 4/0"],
                        "Rating": ["1200A","5A","25 MVA","600V"],
                        "LegacyNote": ["keep","drop","old","note"]})
    new = pd.DataFrame({"ID": ["A1","A2","A3","A5"],
                        "Equipment": ["Breaker 52-2","Relay SEL-351","Xfmr T1","Switch S9"],
                        "Rating": ["2000A","5A","25 MVA","15kV"],
                        "Status": ["In Service","In Service","Spare","New"]})
    with pd.ExcelWriter(f"{TMP}/old.xlsx") as w:
        old.to_excel(w, sheet_name="Equipment", index=False)
        old.to_excel(w, sheet_name="OldOnly", index=False)
    with pd.ExcelWriter(f"{TMP}/new.xlsx") as w:
        new.to_excel(w, sheet_name="Equipment", index=False)
        new.to_excel(w, sheet_name="NewOnly", index=False)


def main():
    build()
    print("\n[1] Cell-level rich text")
    rt_a = diff_strings_rich_text("1200A", "2000A")
    check("added view returns CellRichText", isinstance(rt_a, CellRichText))
    check("added view renders NEW text", "".join(str(x) for x in rt_a).replace("", "") and \
          "".join(getattr(x,'text',str(x)) for x in rt_a) == "2000A",
          "".join(getattr(x,'text',str(x)) for x in rt_a))
    rt_r = diff_strings_removed_rich_text("1200A", "2000A")
    check("removed view renders OLD text",
          "".join(getattr(x,'text',str(x)) for x in rt_r) == "1200A",
          "".join(getattr(x,'text',str(x)) for x in rt_r))
    strikes = [getattr(b,'font',None) for b in rt_r if hasattr(b,'font')]
    check("removed view applies strikethrough", any(f.strike for f in strikes if f))
    check("removed view uses green", any(f.color and 'B050' in str(f.color.rgb or f.color) for f in strikes if f))
    check("identical strings return plain str", isinstance(diff_strings_rich_text("x","x"), str)
          and isinstance(diff_strings_removed_rich_text("x","x"), str))

    print("\n[2] Report generation (key_column=ID)")
    out = f"{TMP}/report.xlsx"
    stats = generate_diff_report(f"{TMP}/old.xlsx", f"{TMP}/new.xlsx", out,
                                 key_column="ID", sheet_name=None)
    check("report file written", os.path.exists(out))
    wb = openpyxl.load_workbook(out, rich_text=True)
    check("Summary tab first", wb.sheetnames[0] == "Summary", str(wb.sheetnames))
    check("paired Added/Removed tabs exist",
          "Equipment (Added)" in wb.sheetnames and "Equipment (Removed)" in wb.sheetnames,
          str(wb.sheetnames))

    s = stats["sheets"][0]
    check("row added detected (A5)", s["rows_added"] == 1, str(s["rows_added"]))
    check("row removed detected (A4)", s["rows_removed"] == 1, str(s["rows_removed"]))
    check("data-only modified row detected (A1)", s["rows_modified_shared"] == 1, str(s["rows_modified_shared"]))
    check("schema churn inflates raw rows_modified", s["rows_modified"] == 3, str(s["rows_modified"]))
    check("data-only cell count excludes schema churn",
          s["cells_changed_shared"] == 2, str(s["cells_changed_shared"]))
    check("column added detected (Status)", s["cols_added"] == ["Status"], str(s["cols_added"]))
    check("column removed detected (LegacyNote)", s["cols_removed"] == ["LegacyNote"], str(s["cols_removed"]))
    check("integrity flagged FAIL on structural change", s["integrity"] == "FAIL", s["integrity"])
    check("sheet only in old reported", stats["sheets_only_in_old"] == ["OldOnly"], str(stats["sheets_only_in_old"]))
    check("sheet only in new reported", stats["sheets_only_in_new"] == ["NewOnly"], str(stats["sheets_only_in_new"]))

    print("\n[3] Strikethrough survives save/reload")
    ws_rem = wb["Equipment (Removed)"]
    found_strike = False
    for row in ws_rem.iter_rows(min_row=2):
        for c in row:
            if isinstance(c.value, CellRichText):
                for b in c.value:
                    f = getattr(b, "font", None)
                    if f is not None and f.strike:
                        found_strike = True
    check("green strikethrough present in saved (Removed) tab", found_strike)

    ws_add = wb["Equipment (Added)"]
    hdr_a = [c.value for c in ws_add[1]]
    check("Status column present on Added tab", hdr_a[0] == "Diff Status", str(hdr_a[:2]))
    check("old-only column retained in header", "LegacyNote" in hdr_a, str(hdr_a))
    statuses_rem = [ws_rem.cell(row=r, column=1).value for r in range(2, ws_rem.max_row+1)]
    check("Removed tab excludes added-only row",
          "Added" not in [str(x) for x in statuses_rem], str(statuses_rem))
    statuses_add = [ws_add.cell(row=r, column=1).value for r in range(2, ws_add.max_row+1)]
    check("Added tab excludes removed-only row",
          "Removed" not in [str(x) for x in statuses_add], str(statuses_add))

    print("\n[4] Whitespace / case-only detection")
    a = pd.DataFrame({"ID":["1","2","3"],"V":["Breaker","relay","Xfmr"]})
    b = pd.DataFrame({"ID":["1","2","3"],"V":["Breaker ","RELAY","Xfmr"]})
    a.to_excel(f"{TMP}/w1.xlsx", index=False); b.to_excel(f"{TMP}/w2.xlsx", index=False)
    stw = generate_diff_report(f"{TMP}/w1.xlsx", f"{TMP}/w2.xlsx", f"{TMP}/ws.xlsx",
                               key_column="ID", sheet_name=0)["sheets"][0]
    check("whitespace-only change flagged", stw["rows_whitespace_only"] == 1, str(stw["rows_whitespace_only"]))
    check("case-only change flagged", stw["rows_case_only"] == 1, str(stw["rows_case_only"]))
    check("whitespace/case noted in integrity notes",
          "whitespace" in stw["notes"] and "case" in stw["notes"], stw["notes"])\

    print("\n[5] Edge cases")
    st2 = generate_diff_report(f"{TMP}/old.xlsx", f"{TMP}/old.xlsx", f"{TMP}/same.xlsx",
                              key_column="ID", sheet_name=None)
    eq = st2["sheets"][0]
    check("identical files -> 0 cells changed", eq["cells_changed"] == 0, str(eq["cells_changed"]))
    check("identical files -> integrity PASS", eq["integrity"] == "PASS", eq["notes"])

    st3 = generate_diff_report(f"{TMP}/old.xlsx", f"{TMP}/new.xlsx", f"{TMP}/pos.xlsx",
                              key_column=None, sheet_name="Equipment")
    check("positional mode keeps all rows",
          st3["sheets"][0]["cells_compared"] > 0, str(st3["sheets"][0]["cells_compared"]))

    st4 = generate_diff_report(f"{TMP}/old.xlsx", f"{TMP}/new.xlsx", f"{TMP}/badkey.xlsx",
                              key_column="NoSuchCol", sheet_name="Equipment")
    check("missing key column degrades gracefully",
          "missing" in st4["sheets"][0]["notes"], st4["sheets"][0]["notes"])

    long_name = "A_Very_Long_Sheet_Name_Exceeding_Limit"
    df = pd.DataFrame({"X": ["1"]})
    with pd.ExcelWriter(f"{TMP}/l1.xlsx") as w: df.to_excel(w, sheet_name=long_name[:31], index=False)
    with pd.ExcelWriter(f"{TMP}/l2.xlsx") as w: pd.DataFrame({"X":["2"]}).to_excel(w, sheet_name=long_name[:31], index=False)
    generate_diff_report(f"{TMP}/l1.xlsx", f"{TMP}/l2.xlsx", f"{TMP}/long.xlsx", sheet_name=None)
    wbl = openpyxl.load_workbook(f"{TMP}/long.xlsx")
    check("long sheet names truncated to <=31 chars",
          all(len(n) <= 31 for n in wbl.sheetnames), str([len(n) for n in wbl.sheetnames]))

    print("\n[6] Order-independent matching (shuffled rows)")
    base = pd.DataFrame({
        "Point Type": ["Analog","Analog","Status","Status","Counter","Analog"],
        "EMS DNP Ind Addr": ["1","2","3","4","5","6"],
        "RTU Variable Name": ["TR1_MW","TR1_MVAR","BKR1_52A","BKR2_52A","FDR1_KWH","BUS1_KV"],
        "Input Device Type": ["M650","M650","SEL-351","SEL-351","M650","M650"],
        "Voltage Level": ["12.5kV"]*6,
    })
    shuffled = base.iloc[[4,0,5,2,1,3]].reset_index(drop=True).copy()
    shuffled.loc[shuffled["RTU Variable Name"] == "TR1_MW", "Voltage Level"] = "13.8kV"
    base.to_excel(f"{TMP}/o1.xlsx", index=False, sheet_name="SCADA Points")
    shuffled.to_excel(f"{TMP}/o2.xlsx", index=False, sheet_name="SCADA Points")

    tiers = [{"columns": ["Point Type","EMS DNP Ind Addr"],
              "require": {"EMS DNP Ind Addr": "int",
                          "Point Type": ["Status","Analog","Counter"]}},
             {"columns": ["Input Device Type"], "require": {"Input Device Type": "nonempty"}},
             {"columns": ["RTU Variable Name"], "require": {"RTU Variable Name": "nonempty"}}]

    so = generate_diff_report(f"{TMP}/o1.xlsx", f"{TMP}/o2.xlsx", f"{TMP}/order.xlsx",
                              sheet_name="SCADA Points", match_keys=tiers)["sheets"][0]
    check("shuffled rows are not reported as added", so["rows_added"] == 0, str(so["rows_added"]))
    check("shuffled rows are not reported as removed", so["rows_removed"] == 0, str(so["rows_removed"]))
    check("all rows matched despite reordering", so["rows_matched"] == 6, str(so["rows_matched"]))
    check("row moves detected", so["rows_reordered"] > 0, str(so["rows_reordered"]))
    check("only the real edit counts as modified",
          so["rows_modified_shared"] == 1, str(so["rows_modified_shared"]))
    check("only one cell changed", so["cells_changed_shared"] == 1, str(so["cells_changed_shared"]))

    print("\n[7] Pure reorder with no edits")
    pure = base.iloc[[5,4,3,2,1,0]].reset_index(drop=True)
    pure.to_excel(f"{TMP}/o3.xlsx", index=False, sheet_name="SCADA Points")
    sp = generate_diff_report(f"{TMP}/o1.xlsx", f"{TMP}/o3.xlsx", f"{TMP}/pure.xlsx",
                              sheet_name="SCADA Points", match_keys=tiers)["sheets"][0]
    check("pure reorder -> 0 data cells changed", sp["cells_changed_shared"] == 0,
          str(sp["cells_changed_shared"]))
    check("pure reorder -> 0 added/removed",
          sp["rows_added"] == 0 and sp["rows_removed"] == 0,
          f"{sp['rows_added']}/{sp['rows_removed']}")
    check("pure reorder still reported as moves", sp["rows_reordered"] > 0, str(sp["rows_reordered"]))

    print("\n[8] Auto key derivation + duplicate-safe keys")
    sa = generate_diff_report(f"{TMP}/o1.xlsx", f"{TMP}/o2.xlsx", f"{TMP}/auto.xlsx",
                              sheet_name="SCADA Points", match_keys="auto")["sheets"][0]
    check("auto mode matches shuffled rows", sa["rows_matched"] == 6, str(sa["rows_matched"]))
    check("auto mode finds the single edit",
          sa["cells_changed_shared"] == 1, str(sa["cells_changed_shared"]))

    dup_old = pd.DataFrame({"Point Type":["Analog","Analog","Analog"],
                            "EMS DNP Ind Addr":["","",""],
                            "RTU Variable Name":["DUP","DUP","UNIQ"],
                            "Input Device Type":["M650","M650","M650"],
                            "Voltage Level":["12.5kV","69kV","12.5kV"]})
    dup_new = dup_old.iloc[[2,1,0]].reset_index(drop=True)
    dup_old.to_excel(f"{TMP}/d1.xlsx", index=False, sheet_name="SCADA Points")
    dup_new.to_excel(f"{TMP}/d2.xlsx", index=False, sheet_name="SCADA Points")
    sd = generate_diff_report(f"{TMP}/d1.xlsx", f"{TMP}/d2.xlsx", f"{TMP}/dup.xlsx",
                              sheet_name="SCADA Points", match_keys=tiers)["sheets"][0]
    check("duplicate keys never drop rows",
          sd["rows_added"] == 0 and sd["rows_removed"] == 0,
          f"{sd['rows_added']}/{sd['rows_removed']}")
    check("duplicate-key rows all matched", sd["rows_matched"] == 3, str(sd["rows_matched"]))
    check("duplicate keys produce no false edits",
          sd["cells_changed_shared"] == 0, str(sd["cells_changed_shared"]))

    print("\n[9] Genuine add/remove still detected under reordering")
    plus = base.iloc[[3,1,0]].reset_index(drop=True).copy()
    plus = pd.concat([plus, pd.DataFrame({"Point Type":["Status"],
                                          "EMS DNP Ind Addr":["99"],
                                          "RTU Variable Name":["NEW_PT"],
                                          "Input Device Type":["SEL-451"],
                                          "Voltage Level":["69kV"]})], ignore_index=True)
    plus.to_excel(f"{TMP}/o4.xlsx", index=False, sheet_name="SCADA Points")
    sg = generate_diff_report(f"{TMP}/o1.xlsx", f"{TMP}/o4.xlsx", f"{TMP}/addrem.xlsx",
                              sheet_name="SCADA Points", match_keys=tiers)["sheets"][0]
    check("new point detected as added", sg["rows_added"] == 1, str(sg["rows_added"]))
    check("dropped points detected as removed", sg["rows_removed"] == 3, str(sg["rows_removed"]))

    print("\n[10] ignore_columns excluded from report and counts")
    si = generate_diff_report(f"{TMP}/o1.xlsx", f"{TMP}/o2.xlsx", f"{TMP}/ign.xlsx",
                              sheet_name="SCADA Points", match_keys=tiers,
                              ignore_columns=["Voltage Level"])["sheets"][0]
    check("ignored column removes its edits from counts",
          si["cells_changed_shared"] == 0, str(si["cells_changed_shared"]))
    wbi = openpyxl.load_workbook(f"{TMP}/ign.xlsx")
    hdr_i = [c.value for c in wbi["SCADA Points (Added)"][1]]
    check("ignored column absent from report header",
          "Voltage Level" not in hdr_i, str(hdr_i))
    check("source-row trace columns present",
          hdr_i[1] == "Old Row" and hdr_i[2] == "New Row", str(hdr_i[:3]))

    print("\n[11] Positional mode still available")
    spos = generate_diff_report(f"{TMP}/o1.xlsx", f"{TMP}/o3.xlsx", f"{TMP}/posmode.xlsx",
                                sheet_name="SCADA Points",
                                match_keys="positional")["sheets"][0]
    check("positional mode flags reordered rows as edits",
          spos["cells_changed_shared"] > 0, str(spos["cells_changed_shared"]))

    print("\n[12] Custom group labels")
    df_la = pd.DataFrame({"ID": ["1", "2"], "Val": ["A", "B"]})
    df_lb = pd.DataFrame({"ID": ["1", "2"], "Val": ["A", "C"]})
    df_la.to_excel(f"{TMP}/lbl_a.xlsx", index=False)
    df_lb.to_excel(f"{TMP}/lbl_b.xlsx", index=False)
    lbl_out = f"{TMP}/lbl_report.xlsx"
    st_lbl = generate_diff_report(
        f"{TMP}/lbl_a.xlsx", f"{TMP}/lbl_b.xlsx", lbl_out,
        group_labels=("Prepared", "Revised"), key_column="ID"
    )
    wb_lbl = openpyxl.load_workbook(lbl_out)
    summary_text = [wb_lbl["Summary"].cell(r, 1).value for r in range(1, 15)]
    check("Summary contains custom group label 0", any("Prepared file" in str(x) for x in summary_text))
    check("Summary contains custom group label 1", any("Revised file" in str(x) for x in summary_text))

    print("\n[13] Ingestion overrides (header_row, skiprows, usecols, column_names, fill_na)")
    banner_data_old = [
        ["Banner line 1", None, None, None, None],
        ["Confidential internal info", None, None, None, None],
        [None, None, "Sub-banner", None, None],
        [None, None, None, None, None],
        ["Rev", "DevType", "DevName", "Addr", "Desc"],
        ["1", "M650", "TR1", "0", "Phase A"],
        ["1", "M650", "TR1", "1", None],
    ]
    wb_raw = openpyxl.Workbook(); ws_raw = wb_raw.active; ws_raw.title = "SCADA"
    for r_idx, row in enumerate(banner_data_old, 1):
        for c_idx, val in enumerate(row, 1): ws_raw.cell(r_idx, c_idx, val)
    wb_raw.save(f"{TMP}/raw_scada_old.xlsx")

    banner_data_new = [
        ["Banner line 1", None, None, None, None],
        ["Confidential internal info", None, None, None, None],
        [None, None, "Sub-banner", None, None],
        [None, None, None, None, None],
        ["Rev", "DevType", "DevName", "Addr", "Desc"],
        ["2", "M650", "TR1", "0", "Phase A modified"],
        ["2", "M650", "TR1", "1", "Phase B added"],
    ]
    wb_raw2 = openpyxl.Workbook(); ws_raw2 = wb_raw2.active; ws_raw2.title = "SCADA"
    for r_idx, row in enumerate(banner_data_new, 1):
        for c_idx, val in enumerate(row, 1): ws_raw2.cell(r_idx, c_idx, val)
    wb_raw2.save(f"{TMP}/raw_scada_new.xlsx")

    custom_cols = ["Device Type", "Device Name", "Point Address", "Description"]
    out_override = f"{TMP}/override_report.xlsx"
    st_over = generate_diff_report(
        f"{TMP}/raw_scada_old.xlsx", f"{TMP}/raw_scada_new.xlsx", out_override,
        sheet_name="SCADA", header_row=5, usecols=[1, 2, 3, 4],
        column_names=custom_cols, fill_na="-", key_column="Point Address"
    )
    check("Header row 5 correctly ingested", st_over["sheets"][0]["rows_old"] == 2)
    check("Custom column names applied", st_over["sheets"][0]["cols_old"] == 4)
    check("Changes detected with overrides", st_over["sheets"][0]["cells_changed"] > 0)
    wb_over = openpyxl.load_workbook(out_override)
    check("Paired tabs exist for SCADA", "SCADA (Added)" in wb_over.sheetnames)

    print("\n[14] Batch directory comparison & file pairing")
    dir_old = f"{TMP}/dir_old"; dir_new = f"{TMP}/dir_new"; out_batch = f"{TMP}/batch_out"
    os.makedirs(dir_old); os.makedirs(dir_new)
    df_la.to_excel(f"{dir_old}/Substation_Alpha.xlsx", index=False)
    df_lb.to_excel(f"{dir_new}/Substation_Alpha.xlsx", index=False)
    df_la.to_excel(f"{dir_old}/Substation_Beta_v1.xlsx", index=False)
    df_lb.to_excel(f"{dir_new}/Substation_Beta_v2_Reviewed.xlsx", index=False)
    df_la.to_excel(f"{dir_old}/Substation_Gamma_OldOnly.xlsx", index=False)

    name_map = {"Substation_Beta_v1": ["Substation_Beta_v2_Reviewed"]}
    batch_res = compare_directories(
        dir_old=dir_old, dir_new=dir_new, output_dir=out_batch,
        group_labels=("Prepared", "Revised"), name_map=name_map,
        key_column="ID"
    )
    check("Batch paired 2 files", len(batch_res["files_compared"]) == 2, str(len(batch_res["files_compared"])))
    check("Unmatched old file detected", len(batch_res["unmatched_old"]) == 1)
    check("Compared files written to output_dir", os.path.exists(f"{out_batch}/Compared_Substation_Alpha.xlsx"))
    check("Batch summary workbook generated", os.path.exists(f"{out_batch}/Batch_Summary.xlsx"))
    wb_bs = openpyxl.load_workbook(f"{out_batch}/Batch_Summary.xlsx")
    check("Batch Summary tab exists", "Batch Summary" in wb_bs.sheetnames)

    print("\n[15] Stacked submittal report layout")
    out_stacked = f"{TMP}/stacked_report.xlsx"
    st_stack = generate_diff_report(
        f"{TMP}/lbl_a.xlsx", f"{TMP}/lbl_b.xlsx", out_stacked,
        group_labels=("Prepared", "Revised"), report_layout="stacked", key_column="ID"
    )
    wb_stack = openpyxl.load_workbook(out_stacked)
    check("Stacked layout tab exists", "Sheet1 (Submittal)" in wb_stack.sheetnames)
    ws_st = wb_stack["Sheet1 (Submittal)"]
    stack_cell_vals = [ws_st.cell(r, 1).value for r in range(1, 15)]
    check("Prepared Submittal banner written", any("Prepared Submittal" in str(v) for v in stack_cell_vals))
    check("Revised Submittal banner written", any("Revised Submittal" in str(v) for v in stack_cell_vals))

    print("\n[16] Rich markup inspection / strikethrough stripping")
    wb_rich = openpyxl.Workbook(); ws_rich = wb_rich.active
    rt_strike = CellRichText()
    rt_strike.append(TextBlock(InlineFont(strike=True), "OldStruckText "))
    rt_strike.append(TextBlock(InlineFont(color="00FF0000", b=True), "NewReplacement"))
    ws_rich["A1"].value = rt_strike
    wb_rich.save(f"{TMP}/rich_test.xlsx")

    sc = StringChecker(file_path=f"{TMP}/rich_test.xlsx")
    eff_text = sc.extract_effective_text(ws_rich["A1"].value, ignore_strikethrough=True)
    check("Strikethrough text dropped in effective text", eff_text == "NewReplacement", eff_text)
    check("Manual markup detected", sc.has_markup(ws_rich["A1"]))

    print("\n" + ("ALL CHECKS PASSED (16/16)" if not FAILS else f"{len(FAILS)} FAILURE(S): {FAILS}"))
    return 1 if FAILS else 0


if __name__ == "__main__":
    sys.exit(main())

