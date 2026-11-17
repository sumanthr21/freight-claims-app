import base64, requests, datetime as dt
import pandas as pd, streamlit as st
import freight_report as fr

st.set_page_config(page_title="Hisco Freight Claims Report", layout="centered")
st.title("Hisco Freight Claims Report")

if st.secrets.get("APP_PASSWORD") and st.text_input("Password", type="password") != st.secrets["APP_PASSWORD"]:
    st.stop()

GH = st.secrets.get("GITHUB", {})
if GH:
    _h = {"Authorization": f"Bearer {GH['token']}"}
    _u = f"https://api.github.com/repos/{GH['repo']}/contents/{GH['path']}"
    _p = {"ref": GH.get("branch", "main")}
    meta = requests.get(_u, headers=_h, params=_p, timeout=60)
    if meta.status_code != 200:
        st.error(f"GitHub replied {meta.status_code}: {meta.json().get('message')}. Check the token, repo name and file name in Secrets.")
        st.stop()
    sha = meta.json()["sha"]
    master = requests.get(_u, headers={**_h, "Accept": "application/vnd.github.raw+json"}, params=_p, timeout=120).content
else:
    up = st.file_uploader("Current report (.xlsx)", type="xlsx"); master, sha = (up.getvalue() if up else None), None

files = st.file_uploader("Agent EOD files (one per agent - select both together)", type="xlsx", accept_multiple_files=True)

if master and files:
    frames = [pd.read_excel(f) for f in files]
    for f, d in zip(files, frames):
        st.write(f"**{f.name}**: {len(d.dropna(how='all'))} rows")
    if st.button("Update report", type="primary"):
        try:
            new, log = fr.update_report(master, frames)
        except ValueError as e:
            st.error(str(e)); st.stop()
        for line in log: (st.warning if "WARNING" in line else st.success)(line)
        if GH:
            r = requests.put(_u, headers=_h, timeout=120, json={"message": f"EOD update {dt.date.today()}",
                "content": base64.b64encode(new).decode(), "sha": sha, "branch": GH.get("branch", "main")})
            st.info("Saved back to GitHub." if r.status_code in (200, 201) else f"Could not save to GitHub ({r.status_code}). Download the file below and tell the report owner.")
        st.download_button("Download updated report", new, f"Hisco_Freight_Claims_{dt.date.today():%d-%b-%Y}.xlsx")
