import base64, re, requests, datetime as dt
import streamlit as st
import om_report as om

st.set_page_config(page_title="Order Maintenance Report", layout="centered")
st.title("Test Equity - OS - Order Maintenance Report")

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

    @st.cache_data(max_entries=1, show_spinner="Loading the report (large file)...")
    def fetch(_sha):   # re-downloaded only when the file in GitHub changes
        return requests.get(_u, headers={**_h, "Accept": "application/vnd.github.raw+json"}, params=_p, timeout=300).content
    master = fetch(sha)
else:
    up = st.file_uploader("Current report (.xlsx)", type="xlsx"); master, sha = (up.getvalue() if up else None), None

@st.cache_data(max_entries=1)
def ctx_of(b): return om.learn(b)

def date_in_name(n):
    m = re.search(r"(\d{1,2})[-_]?(\d{2})[-_]?(20\d{2})", n)
    try: return dt.date(int(m.group(3)), int(m.group(1)), int(m.group(2)))
    except Exception: return None

files = st.file_uploader("Agent EOD files (select all together)", type="xlsx", accept_multiple_files=True)
if master and files:
    ctx = ctx_of(master); agents = ctx["agents"]
    day = st.date_input("EOD date (the EOD files have no date column)", value=date_in_name(files[0].name) or dt.date.today())
    jobs = []
    for f in files:
        g = om.guess_agent(f.name, agents)
        a = st.selectbox(f"Agent for {f.name}", agents, index=agents.index(g) if g in agents else 0, key=f.name)
        jobs.append((f, a, day))
    if st.button("Update report", type="primary"):
        try:
            with st.spinner("Updating (about a minute for a large report)..."):
                new, log = om.update_report(master, jobs)
        except ValueError as e:
            st.error(str(e)); st.stop()
        for line in log: (st.warning if "WARNING" in line else st.success)(line)
        if GH:
            r = requests.put(_u, headers=_h, timeout=600, json={"message": f"EOD update {day}", "content": base64.b64encode(new).decode(),
                                                                "sha": sha, "branch": GH.get("branch", "main")})
            st.info("Saved back to GitHub." if r.status_code in (200, 201) else f"Could not save to GitHub ({r.status_code}). Download the file below and tell the report owner.")
        st.download_button("Download updated report", new, f"Test_Equity_-_OS_-_Order_Maintenance_{day:%d-%b_%Y}.xlsx")
