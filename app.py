"""Test Equity - Cash Application Specialist: EOD loader (Streamlit)."""
import base64
import datetime as dt
import hmac

import requests
import streamlit as st

import cash_report as cr

st.set_page_config(page_title="Cash Application - EOD loader", page_icon=":material/payments:", layout="centered")


# ---------------------------------------------------------------- secrets / password
def secret(name, default=None):
    try:
        return st.secrets.get(name, default)
    except Exception:  # no secrets file at all
        return default


APP_PASSWORD = secret("APP_PASSWORD")
if not APP_PASSWORD:
    st.error("APP_PASSWORD is missing in the app's Secrets. Add it (Manage app > Settings > Secrets) and reload.")
    st.stop()

if not st.session_state.get("authed"):
    st.title("Cash Application - EOD loader")
    pw = st.text_input("Password", type="password")
    if pw:
        if hmac.compare_digest(pw.encode(), str(APP_PASSWORD).encode()):
            st.session_state["authed"] = True
            st.rerun()
        else:
            st.error("Wrong password.")
    st.stop()

GH = secret("GITHUB")
GH = dict(GH) if GH else None


# ---------------------------------------------------------------- GitHub helpers
def gh_headers(raw=False):
    h = {"Authorization": f"Bearer {GH['token']}", "X-GitHub-Api-Version": "2022-11-28"}
    h["Accept"] = "application/vnd.github.raw+json" if raw else "application/vnd.github+json"
    return h


def gh_url():
    return f"https://api.github.com/repos/{GH['repo']}/contents/{GH['path']}"


def gh_load():
    """Return (bytes, sha) of the report in the private repo, or stop with a clear message."""
    p = {"ref": GH.get("branch", "main")}
    meta = requests.get(gh_url(), headers=gh_headers(), params=p, timeout=60)
    if meta.status_code != 200:
        try:
            why = meta.json().get("message")
        except Exception:
            why = meta.text[:200]
        st.error(f"GitHub replied {meta.status_code}: {why}. "
                 "Check the token, repo name and file name in Secrets.")
        st.stop()
    sha = meta.json()["sha"]
    r = requests.get(gh_url(), headers=gh_headers(raw=True), params=p, timeout=180)
    if r.status_code != 200:
        st.error(f"GitHub could not send the file (status {r.status_code}).")
        st.stop()
    return r.content, sha


def gh_save(data, sha, message):
    body = {"message": message, "content": base64.b64encode(data).decode(),
            "sha": sha, "branch": GH.get("branch", "main")}
    r = requests.put(gh_url(), headers=gh_headers(), json=body, timeout=300)
    if r.status_code in (200, 201):
        return r.json()["content"]["sha"], None
    try:
        why = r.json().get("message")
    except Exception:
        why = r.text[:200]
    return None, f"GitHub replied {r.status_code}: {why}"


# ---------------------------------------------------------------- load the master report
st.title("Cash Application - EOD loader")
st.caption("Test Equity - Cash Application Specialist report")

if "master" not in st.session_state:
    if GH:
        with st.spinner("Loading the current report from GitHub..."):
            st.session_state["master"], st.session_state["sha"] = gh_load()
    else:
        st.session_state["master"], st.session_state["sha"] = None, None

if GH:
    c1, c2 = st.columns([4, 1])
    c1.write(f"Report loaded from GitHub ({len(st.session_state['master']) / 1e6:.1f} MB).")
    if c2.button("Reload"):
        for k in ("master", "sha", "result"):
            st.session_state.pop(k, None)
        st.rerun()
else:
    up = st.file_uploader("Current report (.xlsx) - local mode, nothing is saved to GitHub", type="xlsx", key="master_up")
    if up:
        st.session_state["master"] = up.getvalue()
    if not st.session_state.get("master"):
        st.info("Upload the current report to continue.")
        st.stop()

# ---------------------------------------------------------------- upload the EOD files
st.subheader("1. Select the agents' EOD files")
st.caption("Select both agents' files together (hold Ctrl and click each one).")
files = st.file_uploader("EOD files (.xlsx)", type="xlsx", accept_multiple_files=True, key="eod_files")

parsed = []
if files:
    for f in files:
        parsed.append(cr.parse_eod(f.name, f.getvalue()))
    for p in parsed:
        for w in p["warnings"]:
            st.warning(w)
    table = cr.preview(parsed)
    if table:
        st.dataframe(table, hide_index=True)

    detected = cr.detect_date(parsed)
    dates = sorted({r["date"] for p in parsed for sh in p["sheets"] for r in sh["rows"] if r["date"]})
    st.subheader("2. Check the date")
    if len(dates) > 1:
        st.warning("The files contain more than one date: " + ", ".join(d.strftime("%d-%b-%Y") for d in dates)
                   + ". Each row keeps its own date unless you override it below.")
    if detected:
        st.write(f"Date found in the files: **{detected.strftime('%d-%b-%Y')}**")
    else:
        st.error("No date found in the files. Tick the box below and choose the date.")
    other = st.checkbox("Use a different date for all rows", value=(detected is None))
    override = None
    if other:
        override = st.date_input("Date to use", value=detected or dt.date.today(), format="DD/MM/YYYY")

    st.subheader("3. Update the report")
    go = st.button("Update report", type="primary", disabled=not table)
    if go:
        with st.spinner("Updating the report..."):
            new, msgs, stats = cr.update_report(st.session_state["master"], parsed, override_date=override)
        res = dict(msgs=msgs, data=new, name=None, saved=None, save_error=None)
        if new is not None:
            d = stats["date"]
            res["name"] = f"Test Equity - Cash Application Specialist {d.strftime('%d-%b-%Y')}.xlsx"
            if GH:
                with st.spinner("Saving back to GitHub (wait for the confirmation)..."):
                    sha, err = gh_save(new, st.session_state["sha"], f"EOD update {d.isoformat()}")
                if sha:
                    st.session_state["sha"] = sha
                    st.session_state["master"] = new
                    res["saved"] = True
                else:
                    res["saved"] = False
                    res["save_error"] = err
            else:
                st.session_state["master"] = new
        st.session_state["result"] = res

# ---------------------------------------------------------------- show the result
res = st.session_state.get("result")
if res:
    st.divider()
    for level, text in res["msgs"]:
        if level == "error":
            st.error(text)
    if res["data"] is not None:
        st.success("Report updated.")
        if res["saved"] is True:
            st.success("Saved back to GitHub. You can close this page after you download the file.")
        elif res["saved"] is False:
            st.error(f"The report was updated but NOT saved to GitHub ({res['save_error']}). "
                     "Download it now, then upload it to GitHub by hand as report.xlsx. "
                     "Do not load these files again until that is done.")
        for level, text in res["msgs"]:
            if level == "ok":
                st.write(f":green[+] {text}")
        for level, text in res["msgs"]:
            if level == "warn":
                st.warning(text)
            elif level == "info":
                st.info(text)
        st.download_button("Download updated report", res["data"], file_name=res["name"],
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.caption("Excel recalculates the Daily, Weekly and Monthly sheets when you open the file. "
                   "If Excel asks to save changes on closing, that is normal.")
    else:
        for level, text in res["msgs"]:
            if level == "warn":
                st.warning(text)
