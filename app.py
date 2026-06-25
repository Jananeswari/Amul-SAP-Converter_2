"""
Amul SAP PO Converter — Streamlit Web App
Upload e-commerce platform PO files -> Get SAP-mapped monthly box
quantity projections, ready for raising Purchase Orders.
"""

import streamlit as st
import pandas as pd
import os
from mapping_engine import (
    load_master_wide, update_or_add_mapping, bulk_update_from_unmapped_list,
    PLATFORM_COLUMNS, parse_case_pack, find_current_mapping, correct_mapping,
    suggest_fuzzy_matches, MASTER_FILE_PATH, PREVIOUS_MASTER_FILE_PATH,
)
from platform_readers import PLATFORM_READERS, read_generic_po
from etl_engine import convert_platform_orders, build_manager_projection, export_projection_to_excel
from github_sync import is_github_configured, push_file_to_github


def _sync_master_to_github(commit_message: str):
    """Pushes the current master file to GitHub if secrets are configured,
    so changes made through the website survive Streamlit Cloud restarts.
    Silently does nothing (no error shown) if not configured, since local/
    offline use without GitHub is fully supported."""
    if is_github_configured(st.secrets):
        result = push_file_to_github(MASTER_FILE_PATH, "data/Amul_Article_Master.xlsx",
                                      commit_message, st.secrets)
        if result["success"]:
            st.toast("✅ " + result["message"])
        else:
            st.warning("⚠️ " + result["message"])


def _accepted_extensions(platform_cfg: dict) -> list:
    """Streamlit's file_uploader type=[] wants extensions without the dot."""
    return list(platform_cfg["file_types"].keys())


def _read_platform_file(platform_cfg: dict, file_bytes: bytes, filename: str) -> pd.DataFrame:
    """Picks the right reader for the uploaded file's actual extension.
    Always calls whatever reader function is registered for that
    file type — verified readers take just file_bytes, unverified ones
    additionally take their configured column-name kwargs."""
    ext = filename.rsplit(".", 1)[-1].lower()
    if ext not in platform_cfg["file_types"]:
        accepted = ", ".join(platform_cfg["file_types"].keys())
        raise ValueError(f"This platform accepts {accepted} files — got .{ext}")

    type_cfg = platform_cfg["file_types"][ext]
    reader = type_cfg["reader"]
    if not type_cfg.get("verified", False):
        return reader(file_bytes, **type_cfg["default_cols"])
    return reader(file_bytes)

st.set_page_config(page_title="Amul SAP PO Converter", page_icon="📦", layout="wide")

st.markdown("""
<style>
    .stApp { background-color: #FAFAFA; }
    div[data-testid="metric-container"] {
        background: white; border: 1px solid #E5E5E5; border-radius: 8px;
        padding: 14px; box-shadow: 0 1px 3px rgba(0,0,0,0.04);
    }
    .stDownloadButton > button {
        background-color: #14223B !important; color: white !important;
        border-radius: 6px !important; font-weight: 600; width: 100%;
    }
    .flag-pill {
        background: #FFF3CD; color: #946C00; padding: 2px 9px;
        border-radius: 12px; font-size: 12.5px; font-weight: 600;
    }
    .ok-pill {
        background: #E6F4EA; color: #1E7E34; padding: 2px 9px;
        border-radius: 12px; font-size: 12.5px; font-weight: 600;
    }
    h1 { color: #14223B; }
</style>
""", unsafe_allow_html=True)

if "results_cache" not in st.session_state:
    st.session_state.results_cache = {}

# ── Sidebar Navigation ───────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📦 Amul SAP Converter")
    page = st.radio("Go to", ["Convert Orders", "Manage SKU Mapping", "Master File Backup & Restore"],
                     label_visibility="collapsed")
    st.markdown("---")
    st.caption(f"Master mapping file: **{len(load_master_wide())} products** loaded")
    if is_github_configured(st.secrets):
        st.caption("🔗 GitHub sync: connected")
    else:
        st.caption("⚪ GitHub sync: not configured (changes are local only)")

# ══════════════════════════════════════════════════════════════════════════
# PAGE 1: CONVERT ORDERS
# ══════════════════════════════════════════════════════════════════════════
if page == "Convert Orders":
    st.markdown("# Convert Platform Orders to SAP PO Quantities")
    st.caption("Upload PO files from each platform. Get back Amul SAP codes, box quantities, and order dates — ready to raise POs.")
    st.markdown("---")

    available_platforms = list(PLATFORM_READERS.keys())
    uploaded = {}

    cols = st.columns(len(available_platforms))
    for i, platform in enumerate(available_platforms):
        platform_cfg = PLATFORM_READERS[platform]
        accepted_exts = _accepted_extensions(platform_cfg)
        any_unverified = any(not t.get("verified", False) for t in platform_cfg["file_types"].values())
        with cols[i]:
            label = f"**{platform}**" + (" ⚠️" if any_unverified else "")
            st.markdown(label, help="Column mapping not yet confirmed against a real file from this platform" if any_unverified else None)
            f = st.file_uploader(f"Upload {platform}", type=accepted_exts,
                                  key=f"up_{platform}", label_visibility="collapsed")
            if f:
                uploaded[platform] = f
                st.markdown('<span class="ok-pill">✓ Ready</span>', unsafe_allow_html=True)
            st.caption(f"Accepts: {', '.join('.' + e for e in accepted_exts)}")

    st.markdown("---")
    process = st.button("⚡ Convert to SAP PO Quantities", type="primary",
                         disabled=len(uploaded) == 0, use_container_width=True)

    if process and uploaded:
        all_mapped, unmapped_by_platform, stats_list, errors = {}, {}, [], []
        progress = st.progress(0, text="Reading files...")

        for i, (platform, file) in enumerate(uploaded.items()):
            progress.progress((i + 1) / len(uploaded), text=f"Processing {platform}...")
            try:
                orders = _read_platform_file(PLATFORM_READERS[platform], file.read(), file.name)
                result = convert_platform_orders(orders, platform)
                all_mapped[platform] = result["mapped"]
                unmapped_by_platform[platform] = result["unmapped"]
                stats_list.append(result["stats"])
            except Exception as e:
                errors.append((platform, str(e)))

        progress.empty()
        if errors:
            for platform, msg in errors:
                st.error(f"**{platform}**: {msg}")
        st.session_state.results_cache = {
            "all_mapped": all_mapped, "unmapped": unmapped_by_platform, "stats": stats_list,
        }

    if st.session_state.results_cache:
        all_mapped = st.session_state.results_cache["all_mapped"]
        unmapped_by_platform = st.session_state.results_cache["unmapped"]
        stats_list = st.session_state.results_cache["stats"]

        st.success("✅ Conversion complete", icon="✅")
        st.markdown("## Summary")

        k1, k2, k3, k4 = st.columns(4)
        total_rows = sum(s["total_rows"] for s in stats_list)
        total_mapped = sum(s["mapped_rows"] for s in stats_list)
        total_boxes = sum(s["total_boxes"] for s in stats_list)
        total_rounded = sum(s["rounded_up_rows"] for s in stats_list)

        k1.metric("Order Lines Processed", f"{total_rows:,}")
        k2.metric("Successfully Mapped", f"{total_mapped:,}",
                  delta=f"{round(total_mapped/total_rows*100) if total_rows else 0}% match rate")
        k3.metric("Total SAP PO Boxes", f"{total_boxes:,.0f}")
        k4.metric("Rows Rounded Up", f"{total_rounded:,}", delta="check these before PO", delta_color="off")

        st.markdown("---")
        st.markdown("## Monthly PO Projection (Manager View)")
        st.caption("One row per SAP product. Quantity, date, and rounding flag shown separately for each platform.")

        projection = build_manager_projection(all_mapped)

        if not projection.empty:
            rounded_cols = [c for c in projection.columns if "Rounded Up" in c]
            highlight_mask = projection[rounded_cols].any(axis=1) if rounded_cols else pd.Series(False, index=projection.index)

            def highlight_rows(row):
                return ['background-color: #FFF3CD' if highlight_mask.loc[row.name] else '' for _ in row]

            st.dataframe(projection.style.apply(highlight_rows, axis=1), use_container_width=True, height=420)
            st.caption(f"🟨 {int(highlight_mask.sum())} rows had a remainder and were rounded up — review highlighted rows before finalizing the PO.")
        else:
            st.warning("No SKUs were successfully mapped. Check the Unmapped tab below.")

        st.markdown("---")
        st.markdown("## Per-Platform Detail")
        tabs = st.tabs(list(all_mapped.keys()) + ["⚠️ Unmapped SKUs"])
        for i, platform in enumerate(all_mapped.keys()):
            with tabs[i]:
                s = next(s for s in stats_list if s["platform"] == platform)
                c1, c2, c3 = st.columns(3)
                c1.metric("Rows", s["total_rows"])
                c2.metric("Mapped", s["mapped_rows"])
                c3.metric("Unmapped", s["unmapped_rows"])
                df_show = all_mapped[platform][[
                    "platform_sku", "sap_code", "sap_description",
                    "po_qty_pieces", "case_pack_size", "po_qty_cases",
                    "had_remainder", "order_date"
                ]].rename(columns={
                    "platform_sku": "Platform SKU", "sap_code": "SAP Code",
                    "sap_description": "SAP Description", "po_qty_pieces": "Ordered (Pieces)",
                    "case_pack_size": "Case Pack Size", "po_qty_cases": "PO Qty (Cases)",
                    "had_remainder": "Rounded Up?", "order_date": "Order Date",
                })
                st.dataframe(df_show, use_container_width=True, height=300)

        with tabs[-1]:
            unmapped_frames = []
            for platform, df in unmapped_by_platform.items():
                if not df.empty:
                    tmp = df[["platform_sku", "raw_product_name", "order_qty_units"]].copy()
                    tmp.insert(0, "Platform", platform)
                    unmapped_frames.append(tmp)
            if unmapped_frames:
                df_unm = pd.concat(unmapped_frames, ignore_index=True)
                df_unm.columns = ["Platform", "SKU", "Product Name (from platform)", "Qty Ordered"]

                # Dedupe by Platform+SKU for the fill-in template — manager only
                # needs to type the SAP Code once per unique SKU, not once per order line
                template_df = (
                    df_unm.groupby(["Platform", "SKU"], as_index=False)
                    .agg({"Product Name (from platform)": "first", "Qty Ordered": "sum"})
                )
                template_df["SAP Code"] = ""

                st.warning(f"{len(df_unm)} order lines ({len(template_df)} unique SKUs) could not be matched to a SAP code.")
                st.caption("Fill in the **SAP Code** column below, then go to **Manage SKU Mapping → Bulk Update** in the sidebar and upload it back.")
                st.dataframe(df_unm, use_container_width=True, height=300)

                import io as _io
                buf = _io.BytesIO()
                with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                    template_df.to_excel(writer, sheet_name="Fill SAP Code", index=False)
                st.download_button(
                    "📥 Download Fill-In Template (for Bulk Mapping Update)",
                    data=buf.getvalue(),
                    file_name="Unmapped_SKUs_Fill_SAP_Code.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

                st.markdown("---")
                st.markdown("#### 🔎 Suggest a match for one SKU")
                st.caption("Pick a SKU below to see the closest-matching SAP products by name similarity, then add it directly to the master file.")

                unique_skus = template_df["Product Name (from platform)"].tolist()
                pick = st.selectbox("Unmapped product", options=unique_skus, key="fuzzy_pick")

                if pick:
                    matches = suggest_fuzzy_matches(pick, top_n=5)
                    if matches.empty:
                        st.info("No reasonable matches found — this is likely a genuinely new product.")
                    else:
                        matches_display = matches.copy()
                        matches_display["similarity"] = matches_display["similarity"].round(1).astype(str) + "%"
                        st.dataframe(matches_display, use_container_width=True, hide_index=True)

                        chosen_sap_code = st.selectbox(
                            "If one of these is correct, select its SAP Code to add this mapping",
                            options=["-- none, this is a new product --"] + matches["SAP Code"].tolist(),
                            key="fuzzy_sap_choice",
                        )
                        matched_row = template_df[template_df["Product Name (from platform)"] == pick].iloc[0]

                        if chosen_sap_code != "-- none, this is a new product --":
                            if st.button(f"Add: {matched_row['SKU']} → {chosen_sap_code}", type="primary"):
                                result = update_or_add_mapping(
                                    chosen_sap_code, "", "",
                                    {matched_row["Platform"]: matched_row["SKU"]},
                                )
                                _sync_master_to_github(
                                    f"Map {matched_row['Platform']} SKU {matched_row['SKU']} to {chosen_sap_code} (fuzzy match)"
                                )
                                if result["action"] in ("added_new_product", "updated_existing"):
                                    st.success(f"✅ Mapped {matched_row['SKU']} to {chosen_sap_code}.")
                                else:
                                    st.warning("No change made — that platform may already have a different SKU mapped to this SAP Code.")
            else:
                st.success("🎉 All SKUs were mapped successfully!")

        st.markdown("---")
        st.markdown("## Download")
        excel_bytes = export_projection_to_excel(projection, unmapped_by_platform, stats_list)
        st.download_button(
            "📥 Download SAP PO Projection (Excel)",
            data=excel_bytes,
            file_name="Amul_SAP_PO_Projection.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        st.caption("Includes: Summary | Monthly PO Projection (rounded-up rows highlighted) | Unmapped SKUs")

# ══════════════════════════════════════════════════════════════════════════
# PAGE 2: MANAGE SKU MAPPING
# ══════════════════════════════════════════════════════════════════════════
elif page == "Manage SKU Mapping":
    st.markdown("# Manage SKU Mapping")
    st.caption("Add new products or new platform SKU codes to the master mapping table.")
    st.markdown("---")

    master_df = load_master_wide()

    tab1, tab2, tab3, tab4 = st.tabs([
        "➕ Add New Mapping", "📤 Bulk Update (Unmapped List)",
        "✏️ Correct Existing Mapping", "🔍 View / Search Mapping",
    ])

    with tab1:
        st.markdown("### Add a new product, or attach a platform SKU to an existing SAP Code")
        st.caption("If the SAP Code already exists, leave description/FG group blank — just fill in the new platform SKU(s) and they'll be attached to that existing product.")

        with st.form("add_mapping_form", clear_on_submit=True):
            c1, c2 = st.columns(2)
            with c1:
                sap_code = st.text_input("SAP Code *", placeholder="e.g. TDMCP01")
                sap_desc = st.text_input("SAP Product Description (required only for a NEW SAP Code)", placeholder="e.g. Amul Taaza Fresh Toned Milk 12x1 Ltr TP")
            with c2:
                fg_group = st.text_input("FG Group Description (optional)", placeholder="e.g. Milk - UHT")
                if sap_desc:
                    cp, conf = parse_case_pack(sap_desc)
                    badge = "✅ detected" if conf == "high" else ("⚠️ ambiguous, please verify" if conf == "low" else "ℹ️ defaulted to 1 (no pack pattern found)")
                    st.info(f"Case pack size: **{cp} units/box** — {badge}")

            st.markdown("**Platform SKU Codes** (fill in whichever platforms apply)")
            platform_inputs = {}
            pcols = st.columns(4)
            for i, platform in enumerate(PLATFORM_COLUMNS):
                with pcols[i % 4]:
                    platform_inputs[platform] = st.text_input(platform, key=f"pf_{platform}")

            submitted = st.form_submit_button("Save Mapping", type="primary", use_container_width=True)

            if submitted:
                sap_code_clean = sap_code.strip() if sap_code else ""
                is_existing = sap_code_clean.upper() in master_df["SAP Code"].astype(str).str.strip().str.upper().values

                if not sap_code_clean:
                    st.error("SAP Code is required.")
                elif not is_existing and not sap_desc:
                    st.error(f"SAP Code '{sap_code_clean}' doesn't exist yet — SAP Product Description is required to create it.")
                elif not any(v.strip() for v in platform_inputs.values()):
                    st.error("Enter at least one platform SKU code.")
                else:
                    result = update_or_add_mapping(sap_code_clean, sap_desc, fg_group, platform_inputs)

                    if result["action"] == "added_new_product":
                        _sync_master_to_github(f"Add new SAP Code {sap_code_clean}")
                        st.success(f"✅ Added new SAP Code **{sap_code_clean}** with platform SKU(s): {', '.join(result['updated_platforms'])}")
                        st.rerun()
                    elif result["action"] == "updated_existing":
                        _sync_master_to_github(f"Attach SKU(s) to existing SAP Code {sap_code_clean}: {', '.join(result['updated_platforms'])}")
                        st.success(f"✅ SAP Code **{sap_code_clean}** already existed — attached new SKU(s) for: {', '.join(result['updated_platforms'])}")
                        if result["conflicts"]:
                            for plat, existing_sku in result["conflicts"].items():
                                st.warning(f"⚠️ {plat} already had a different SKU mapped (`{existing_sku}`) — not overwritten. Use **Correct Existing Mapping** if this needs to change.")
                        st.rerun()
                    else:  # no_change
                        st.warning("No new SKUs were added — every platform you entered already had the exact same SKU mapped, or all conflicted with an existing different SKU:")
                        for plat, existing_sku in result["conflicts"].items():
                            st.warning(f"⚠️ {plat} already maps to a different SKU (`{existing_sku}`).")

    with tab2:
        st.markdown("### Bulk update from a filled-in Unmapped SKUs list")
        st.caption(
            "Go to **Convert Orders → Unmapped SKUs tab**, download the fill-in template, "
            "type the correct **SAP Code** next to each SKU, then upload it here. "
            "Only attaches SKUs to SAP Codes that already exist — empty platform slots are filled in, "
            "existing different SKUs are never overwritten."
        )

        bulk_file = st.file_uploader("Upload filled-in template", type=["xlsx", "xls"], key="bulk_upload")

        if bulk_file:
            try:
                preview_df = pd.read_excel(bulk_file)
                st.markdown(f"**Preview** ({len(preview_df)} rows)")
                st.dataframe(preview_df, use_container_width=True, height=250)

                if st.button("Apply Bulk Update", type="primary", use_container_width=True):
                    bulk_file.seek(0)
                    apply_df = pd.read_excel(bulk_file)
                    result = bulk_update_from_unmapped_list(apply_df)

                    if result["updated_count"]:
                        _sync_master_to_github(f"Bulk update: mapped {result['updated_count']} SKU(s)")
                        st.success(f"✅ {result['updated_count']} SKU(s) successfully mapped.")
                        with st.expander("See what was added"):
                            st.dataframe(pd.DataFrame(result["updated_rows"]), use_container_width=True)
                    else:
                        st.info("No new SKUs were mapped — see details below for why.")

                    if result["conflicts"]:
                        st.warning(f"⚠️ {len(result['conflicts'])} row(s) skipped — that platform already had a DIFFERENT SKU mapped to the SAP Code given. Nothing was overwritten.")
                        st.dataframe(pd.DataFrame(result["conflicts"]), use_container_width=True)

                    if result["skipped_unknown_sap_code"]:
                        st.warning(f"⚠️ {len(result['skipped_unknown_sap_code'])} row(s) skipped — the SAP Code typed in doesn't exist in the master file yet. Use **Add New Mapping** for these instead.")
                        st.dataframe(pd.DataFrame(result["skipped_unknown_sap_code"]), use_container_width=True)

                    if result["skipped_no_sap_code"]:
                        st.info(f"ℹ️ {len(result['skipped_no_sap_code'])} row(s) skipped — SAP Code column was left blank.")
                        st.dataframe(pd.DataFrame(result["skipped_no_sap_code"]), use_container_width=True)

            except ValueError as e:
                st.error(str(e))
            except Exception as e:
                st.error(f"Couldn't read this file: {e}")

    with tab3:
        st.markdown("### Correct an existing mapping")
        st.caption(
            "Use this when a mapping was found to be **wrong** — not for adding a new SKU "
            "(use Add New Mapping or Bulk Update for that, which never overwrite). "
            "This always shows the current value first and requires you to confirm before changing it."
        )

        lookup_mode = st.radio("Find the mapping by:", ["SAP Code", "Platform + SKU"], horizontal=True)

        current_row = None
        if lookup_mode == "SAP Code":
            sap_code_lookup = st.text_input("SAP Code to correct", placeholder="e.g. TDMCP01", key="correct_sap_lookup")
            if sap_code_lookup:
                result_df = find_current_mapping(sap_code=sap_code_lookup)
                if result_df.empty:
                    st.error(f"SAP Code '{sap_code_lookup}' not found.")
                else:
                    current_row = result_df.iloc[0]
        else:
            c1, c2 = st.columns(2)
            with c1:
                platform_lookup = st.selectbox("Platform", options=PLATFORM_COLUMNS, key="correct_platform_lookup")
            with c2:
                sku_lookup = st.text_input("SKU as it currently appears", key="correct_sku_lookup")
            if sku_lookup:
                result_df = find_current_mapping(platform=platform_lookup, sku=sku_lookup)
                if result_df.empty:
                    st.error(f"No SAP Code currently has '{sku_lookup}' mapped under {platform_lookup}.")
                else:
                    current_row = result_df.iloc[0]

        if current_row is not None:
            st.markdown("---")
            st.markdown(f"**Found:** `{current_row['SAP Code']}` — {current_row['Product Description as per SAP']}")

            platform_to_fix = st.selectbox(
                "Which platform's SKU is wrong on this product?",
                options=PLATFORM_COLUMNS, key="correct_platform_to_fix",
            )
            old_val = current_row.get(platform_to_fix)
            old_val_display = str(old_val).strip() if pd.notna(old_val) else "(empty)"
            st.info(f"Current value for **{platform_to_fix}**: `{old_val_display}`")

            new_val = st.text_input("Correct SKU value", key="correct_new_value")

            if new_val:
                st.warning(f"This will change **{platform_to_fix}** for `{current_row['SAP Code']}` "
                           f"from `{old_val_display}` to `{new_val.strip()}`. This cannot be undone "
                           f"except by restoring the previous-version backup.")
                confirm = st.checkbox("Yes, I've checked this and want to overwrite it", key="correct_confirm")
                if confirm and st.button("Apply Correction", type="primary"):
                    result = correct_mapping(current_row["SAP Code"], platform_to_fix, new_val)
                    if result["success"]:
                        _sync_master_to_github(
                            f"Correct {platform_to_fix} mapping for {current_row['SAP Code']}: "
                            f"'{result['old_value']}' -> '{new_val.strip()}'"
                        )
                        st.success(f"✅ {result['message']}")
                        st.rerun()
                    else:
                        st.error(result["message"])

    with tab4:
        st.markdown("### Search the master mapping table")
        search = st.text_input("Search by SAP Code, Description, or FG Group", placeholder="Type to filter...")
        display_df = master_df
        if search:
            mask = (
                master_df["SAP Code"].astype(str).str.contains(search, case=False, na=False)
                | master_df["Product Description as per SAP"].astype(str).str.contains(search, case=False, na=False)
                | master_df["FG Group Description"].astype(str).str.contains(search, case=False, na=False)
            )
            display_df = master_df[mask]
        st.caption(f"Showing {len(display_df)} of {len(master_df)} products")
        st.dataframe(display_df, use_container_width=True, height=500)


# ══════════════════════════════════════════════════════════════════════════
# PAGE 3: MASTER FILE BACKUP & RESTORE
# ══════════════════════════════════════════════════════════════════════════
elif page == "Master File Backup & Restore":
    st.markdown("# Master File Backup & Restore")
    st.caption("Download the current or previous master file, or import an updated master file to replace what's currently loaded.")
    st.markdown("---")

    st.markdown("## Download")
    st.caption(
        "Every change made through this website (Add New Mapping, Bulk Update, "
        "Correct Existing Mapping, or importing a file below) is also pushed to "
        "GitHub automatically" if is_github_configured(st.secrets) else
        "GitHub sync is not configured, so changes made through this website only "
        "last until the app restarts — download regularly as a backup."
    )

    dl_col1, dl_col2 = st.columns(2)
    with dl_col1:
        st.markdown("**Current master file**")
        st.caption("Includes every edit made so far in this session.")
        with open(MASTER_FILE_PATH, "rb") as f:
            st.download_button(
                "📥 Download Current Master File",
                data=f.read(),
                file_name="Amul_Article_Master_CURRENT.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
    with dl_col2:
        st.markdown("**Previous version**")
        st.caption("Automatically saved right before the last change — for comparison.")
        if os.path.exists(PREVIOUS_MASTER_FILE_PATH):
            with open(PREVIOUS_MASTER_FILE_PATH, "rb") as f:
                st.download_button(
                    "📥 Download Previous Version",
                    data=f.read(),
                    file_name="Amul_Article_Master_PREVIOUS.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
        else:
            st.info("No previous version yet — this is created automatically the first time a change is made.")

    st.markdown("---")
    st.markdown("## Import a Master File")
    st.warning(
        "⚠️ This **replaces** the entire master mapping table with the file you upload. "
        "Make sure it has a 'Sheet1' with the same columns as the current file "
        "(SAP Code, Product Description as per SAP, and one column per platform). "
        "The current file is automatically snapshotted as the 'previous version' before this happens."
    )

    import_file = st.file_uploader("Upload master file to import", type=["xlsx"], key="import_master")
    if import_file:
        try:
            preview = pd.read_excel(import_file, sheet_name="Sheet1")
            required_cols = {"SAP Code", "Product Description as per SAP"}
            missing = required_cols - set(preview.columns.str.strip())
            if missing:
                st.error(f"This file is missing required column(s): {', '.join(missing)}. Import cancelled.")
            else:
                st.success(f"File looks valid — {len(preview)} products found.")
                st.dataframe(preview.head(10), use_container_width=True)

                confirm_import = st.checkbox(
                    f"Yes, replace the current {len(load_master_wide())}-product master file with this {len(preview)}-product file",
                    key="confirm_import",
                )
                if confirm_import and st.button("Import and Replace Master File", type="primary"):
                    import shutil as _shutil
                    if os.path.exists(MASTER_FILE_PATH):
                        _shutil.copy(MASTER_FILE_PATH, PREVIOUS_MASTER_FILE_PATH)
                    import_file.seek(0)
                    with open(MASTER_FILE_PATH, "wb") as f:
                        f.write(import_file.read())
                    _sync_master_to_github(f"Import replacement master file ({len(preview)} products)")
                    st.success("✅ Master file imported and is now active.")
                    st.rerun()
        except Exception as e:
            st.error(f"Couldn't read this file: {e}")
