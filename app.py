"""
Streamlit web UI for the course roster builder.

Run locally:   streamlit run app.py
Deploy free:   push this repo to GitHub, then deploy at https://share.streamlit.io

Repo must contain:
    app.py
    build_rosters.py
    requirements.txt
    templates/   <- your three blank template .docx files
"""

import io
import os
import tempfile
import zipfile

import streamlit as st
import build_rosters as br

TEMPLATE_DIR = "templates"
DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

st.set_page_config(page_title="Course Roster Builder", page_icon="📋")

# --- Optional password gate -------------------------------------------------
# To require a password, add this to the app's Secrets in Streamlit Cloud:
#     password = "your-password-here"
# Leave it unset to keep the app open.
_pw = st.secrets.get("password", None) if hasattr(st, "secrets") else None
if _pw:
    if st.session_state.get("authed") is not True:
        entered = st.text_input("Password", type="password")
        if entered and entered == _pw:
            st.session_state["authed"] = True
            st.rerun()
        elif entered:
            st.error("Incorrect password.")
        st.stop()

# --- UI ---------------------------------------------------------------------
st.title("Course Roster Builder")
st.caption("Upload the Excel export and download the Word rosters.")

uploaded = st.file_uploader("Excel export (.xlsx)", type=["xlsx"])
c1, c2 = st.columns(2)
keep_order = c1.checkbox("Keep export order (don't sort A–Z)")
include_cancelled = c2.checkbox("Include cancelled courses")

if uploaded and st.button("Generate rosters", type="primary"):
    template_map = br.load_template_map(TEMPLATE_DIR)
    if not template_map:
        st.error(f"No templates found in '{TEMPLATE_DIR}/'. "
                 "Add your three blank template .docx files there and redeploy.")
        st.stop()

    log, files = [], []          # files = list of (filename, bytes)
    with tempfile.TemporaryDirectory() as tmp:
        xlsx_path = os.path.join(tmp, "export.xlsx")
        with open(xlsx_path, "wb") as f:
            f.write(uploaded.getbuffer())
        out_dir = os.path.join(tmp, "out")
        os.makedirs(out_dir, exist_ok=True)

        rows = br.read_export(xlsx_path)
        made = 0
        for rec in rows:
            status, info, ctype, n = br.build_one(
                rec, template_map, out_dir,
                sort_alpha=not keep_order,
                include_cancelled=include_cancelled)
            if status == "ok":
                made += 1
                with open(info, "rb") as fh:
                    files.append((os.path.basename(info), fh.read()))
                log.append(("ok", f"{os.path.basename(info)} — {ctype}, {n} participants"))
            else:
                log.append(("skip", f"{rec.get('customer')} — {info}"))

    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in files:
            z.writestr(name, data)
    zip_buf.seek(0)

    st.session_state["results"] = {
        "files": files,
        "zip": zip_buf.getvalue(),
        "log": log,
        "made": made,
        "skipped": sum(1 for s, _ in log if s == "skip"),
    }

# --- Results (persisted so download clicks don't clear them) ----------------
res = st.session_state.get("results")
if res:
    st.divider()
    st.subheader(f"{res['made']} roster(s) created"
                 + (f", {res['skipped']} skipped" if res["skipped"] else ""))

    if res["files"]:
        st.download_button("⬇ Download all rosters (.zip)", res["zip"],
                           file_name="rosters.zip", mime="application/zip",
                           type="primary")
        with st.expander("Download individual files"):
            for i, (name, data) in enumerate(res["files"]):
                st.download_button(name, data, file_name=name,
                                   mime=DOCX_MIME, key=f"dl_{i}")

    for status, msg in res["log"]:
        (st.success if status == "ok" else st.warning)(
            ("" if status == "ok" else "Skipped: ") + msg)
