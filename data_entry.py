import logging
import sys

import streamlit as st
import json
import os
from datetime import datetime
import io  # මෙය මුලට එකතු කරන්න

# --- Secrets Bridge: Streamlit Cloud secrets → env var ---
if "GEMINI_API_KEY" in st.secrets:
    os.environ.setdefault("GEMINI_API_KEY", st.secrets["GEMINI_API_KEY"])

# --- Ensure required directories exist ---
os.makedirs("outputs", exist_ok=True)
os.makedirs("saved_chats", exist_ok=True)

from main2 import generate_report
if "needs_reset" not in st.session_state:
    st.session_state.needs_reset = False

if st.session_state.needs_reset:
    # 1. භාව සටහන් (Multi-selects) හිස් කිරීම
    for i in range(1, 13):
        # Key එක Session State එකේ තියෙනවා නම් විතරක් හිස් කරන්න
        if f"l_ms_{i}" in st.session_state:
            st.session_state[f"l_ms_{i}"] = []
        if f"n_ms_{i}" in st.session_state:
            st.session_state[f"n_ms_{i}"] = []

    # 2. විශේෂ ගැටලු හිස් කිරීම
    st.session_state["special_issues_input"] = ""

    # 3. Reset කරලා ඉවරයි කියලා දන්වන්න
    st.session_state.needs_reset = False

# --- Reset Logic අවසානයි ---

# --- 2. Session State ආරම්භ කිරීම (Initialization) ---
# මේ කොටස අර උඩ කොටසට පහළින් තියෙන්න ඕනේ
for i in range(1, 13):
    if f"l_ms_{i}" not in st.session_state: st.session_state[f"l_ms_{i}"] = []
    if f"n_ms_{i}" not in st.session_state: st.session_state[f"n_ms_{i}"] = []

for handler in logging.root.handlers[:]:
    logging.root.removeHandler(handler)

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s]: %(message)s',
    handlers=[logging.StreamHandler(sys.stdout)], # Terminal එකට ලොග් එවීමට
    force=True
)
logger = logging.getLogger("AstroApp")

JSON_FILE = "birth_records.json"
st.set_page_config(page_title="Astro Stable UI", layout="wide")

# --- 3. Session State Initialization ---
if 'init_done' not in st.session_state:
    for i in range(1, 13):
        if f"l_ms_{i}" not in st.session_state: st.session_state[f"l_ms_{i}"] = []
        if f"n_ms_{i}" not in st.session_state: st.session_state[f"n_ms_{i}"] = []
    st.session_state.init_done = True

# --- Helper Functions ---
def load_existing_data():
    if os.path.exists(JSON_FILE):
        with open(JSON_FILE, "r", encoding="utf-8") as f:
            try:
                return json.load(f)
            except:
                return []
    return []


def save_data(new_record):
    data = load_existing_data()
    data.append(new_record)
    with open(JSON_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    st.success("දත්ත සාර්ථකව ගබඩා කළා! ✅")


def save_all_records(records):
    with open(JSON_FILE, "w", encoding="utf-8") as f:
        json.dump(records, f, ensure_ascii=False, indent=2)


# --- Session State කළමනාකරණය ---
# මුලින්ම සෑම භාවයකටම අදාළ Key එකක් Session State හි සාදා ගනිමු
if 'init_done' not in st.session_state:
    # --- කේතයේ මුල තිබිය යුතු කොටස ---
    for i in range(1, 13):
        if f"l_ms_{i}" not in st.session_state:
            st.session_state[f"l_ms_{i}"] = []
        if f"n_ms_{i}" not in st.session_state:
            st.session_state[f"n_ms_{i}"] = []

all_planets = ["රවි", "චන්ද්‍ර", "කුජ", "බුද", "ගුරු", "ශුක්‍ර", "ශනි", "රාහු", "කේතු", "යුරේනස්", "නැප්චූන්", "ප්ලූටෝ"]

# --- SIDEBAR ---
with st.sidebar:
    st.header("📋 විස්තර")
    gender = st.radio("ලිංගභේදය", ["පිරිමි", "ගැහැනු"])
    # නව Radio Button එක - Default 'Normal' select වී ඇත
    package_type = st.radio("පැකේජය තෝරන්න", ["Normal", "VIP (Rs. 1500)"], index=0)
    phone = st.text_input("දුරකතන අංකය")
    dob_date = st.date_input("උපන් දිනය", min_value=datetime(1900, 1, 1))
    hour = st.slider("පැය", 0, 23, 12)
    minute = st.slider("විනාඩි", 0, 59, 0)

    if st.button("දත්ත මකන්න"):
        # JSON එක මකා දැමීම
        if os.path.exists(JSON_FILE):
            os.remove(JSON_FILE)

        # Session State එකේ ඇති භාව වල දත්ත Reset කිරීම
        for i in range(1, 13):
            st.session_state[f"l_ms_{i}"] = []
            st.session_state[f"n_ms_{i}"] = []

        st.success("සියලු දත්ත මකා දැමුවා!")
        st.rerun()

    # --- JSON Records Viewer / Editor ---
    st.divider()
    st.subheader("📋 Added Records")
    records = load_existing_data()
    if not records:
        st.info("No records yet.")
    else:
        st.write(f"{len(records)} record(s)")
        for idx, record in enumerate(records):
            try:
                phone = record["කේන්ද්‍ර_සටහන"]["දුරකතන_අංකය"]
                dob   = record["කේන්ද්‍ර_සටහන"]["උපන්_දිනය"]
            except (KeyError, TypeError):
                phone, dob = "?", "?"
            edit_key = f"editing_{idx}"
            if edit_key not in st.session_state:
                st.session_state[edit_key] = False

            with st.expander(f"#{idx+1} | {phone} | {dob}"):
                if st.session_state[edit_key]:
                    edited = st.text_area(
                        "Edit JSON",
                        value=json.dumps(record, ensure_ascii=False, indent=2),
                        height=300,
                        key=f"edit_area_{idx}"
                    )
                    c1, c2 = st.columns(2)
                    if c1.button("💾 Save", key=f"save_{idx}"):
                        try:
                            updated = json.loads(edited)
                            records[idx] = updated
                            save_all_records(records)
                            st.session_state[edit_key] = False
                            st.rerun()
                        except json.JSONDecodeError as e:
                            st.error(f"Invalid JSON: {e}")
                    if c2.button("❌ Cancel", key=f"cancel_{idx}"):
                        st.session_state[edit_key] = False
                        st.rerun()
                else:
                    st.json(record)
                    c1, c2 = st.columns(2)
                    if c1.button("✏️ Edit", key=f"edit_{idx}"):
                        st.session_state[edit_key] = True
                        st.rerun()
                    if c2.button("🗑️ Delete", key=f"del_{idx}"):
                        records.pop(idx)
                        save_all_records(records)
                        st.rerun()

    # --- Generated Reports File Browser ---
    st.divider()
    st.subheader("📄 Generated Reports")
    if st.button("🔄 Refresh", key="refresh_reports"):
        st.rerun()
    outputs_dir = "outputs"
    if not os.path.exists(outputs_dir):
        st.info("No reports generated yet.")
    else:
        folders = sorted([
            f for f in os.listdir(outputs_dir)
            if os.path.isdir(os.path.join(outputs_dir, f))
        ])
        if not folders:
            st.info("No reports generated yet.")
        else:
            for folder_name in folders:
                folder_path = os.path.join(outputs_dir, folder_name)
                with st.expander(f"📁 {folder_name}"):
                    # docx first, then txt files
                    for fname in os.listdir(folder_path):
                        fpath = os.path.join(folder_path, fname)
                        if not os.path.isfile(fpath):
                            continue
                        if fname.endswith(".docx"):
                            mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        elif fname.endswith(".txt"):
                            mime = "text/plain"
                        else:
                            continue
                        col1, col2 = st.columns([3, 1])
                        col1.write(fname)
                        with open(fpath, "rb") as f:
                            col2.download_button(
                                "⬇",
                                data=f,
                                file_name=fname,
                                mime=mime,
                                key=fpath
                            )

# --- MAIN PAGE ---
rashi_list = ["මේෂ", "වෘෂභ", "මිථුන", "කටක", "සිංහ", "කන්‍යා", "තුලා", "වෘශ්චික", "ධනු", "මකර", "කුම්භ", "මීන"]

st.subheader("♈ ලග්නය සහ ☸️ නවංශකය")
lagna = st.radio("L", rashi_list, key="l_r_input", horizontal=True)
navansha = st.radio("N", rashi_list, key="n_r_input", horizontal=True)

st.divider()


# --- භාව සටහන ඇඳීමේ ශ්‍රිතය ---
def render_bhava_grid(prefix, title):
    st.subheader(title)

    # 1. දැනට සියලුම භාව වල තෝරාගෙන ඇති ග්‍රහයන් ලැයිස්තුවක් ගනිමු
    used_in_this_chart = []
    for i in range(1, 13):
        used_in_this_chart.extend(st.session_state[f"{prefix}_{i}"])

    cols = st.columns(6)
    chart_results = {}

    for i in range(1, 13):
        key = f"{prefix}_{i}"
        with cols[(i - 1) % 6]:
            current_selection = st.session_state[key]
            # ඉතිරි ග්‍රහයන් = (සියල්ල - පාවිච්චි කළ අය) + මේ කොටුවේ දැනට ඉන්න අය
            options = [p for p in all_planets if p not in used_in_this_chart or p in current_selection]

            # multiselect එක කෙලින්ම session_state key එකට සම්බන්ධ කිරීමෙන් අගය මැකී නොයයි
            st.multiselect(f"භාවය {i}", options=options, key=key)
            chart_results[str(i)] = st.session_state[key]

    return chart_results


# ලග්න කේන්ද්‍රය
l_data = render_bhava_grid("l_ms", "🪐 ලග්න කේන්ද්‍රය")

st.divider()

# නවංශක කේන්ද්‍රය
n_data = render_bhava_grid("n_ms", "☸️ නවංශක කේන්ද්‍රය")

st.divider()
# Main Page එකේ ඇති කොටස
special_issues = st.text_area("විශේෂ ගැටලු", key="special_issues_input")


# --- පහළ ඇති Button එක ---
if st.button("JSON එකට ඇතුළත් කරන්න 🚀", use_container_width=True):

    # 1. දත්ත Save කිරීමේ කොටස (මේක වෙනස් කරන්න එපා)
    issues_val = st.session_state.get("special_issues_input", "")
    new_record = {
        "package_type": package_type,
        "කේන්ද්‍ර_සටහන": {
            "ලිංගහේදය": gender,
            "උපන්_දිනය": dob_date.strftime("%Y/%m/%d"),
            "උපන්_වේලාව": f"{hour:02d}:{minute:02d}",
            "දුරකතන_අංකය": phone,
            "ලග්න_කේන්ද්‍රය": {"ලග්නය": lagna, "ග්‍රහ_පිහිටීම්": l_data},
            "නවංශක_කේන්ද්‍රය": {"නවංශකය": navansha, "ග්‍රහ_පිහිටීම්": n_data},
            "විශේශ_ගැටලු": [q.strip() for q in issues_val.split('\n') if q.strip()]
        }
    }
    save_data(new_record)

    # 2. වැදගත්ම කොටස: Reset Trigger එක On කිරීම පමණයි!
    # මෙතනදී දත්ත මකන්න යන්න එපා. Error එනවා.
    st.session_state.needs_reset = True

    # 3. Page එක Refresh කරන්න
    st.rerun()

st.divider()
st.subheader("📄 වාර්තා උත්පාදනය (Report Generation)")

if st.button("සම්පූර්ණ ජ්‍යොතිෂ වාර්තාව සාදන්න 🚀", use_container_width=True):
    with st.spinner("කෘතිම බුද්ධිය (AI) හරහා වාර්තාව සකස් කරමින් පවතිී... කරුණාකර රැඳී සිටින්න."):
        try:
            # ඔයාගේ generate_report() function එක මෙතනදී call වෙනවා
            # වැදගත්: එක් record එකකට වඩා තියෙනවා නම් ඒ හැම එකකටම මේක ක්‍රියාත්මක වේවි
            output_files = generate_report()

            st.success("වාර්තාව සාර්ථකව සකස් කළා! ✅")
            for docx_path in (output_files or []):
                if os.path.exists(docx_path):
                    with open(docx_path, "rb") as f:
                        st.download_button(
                            label=f"📥 Download: {os.path.basename(docx_path)}",
                            data=f,
                            file_name=os.path.basename(docx_path),
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )


        except Exception as e:
            st.error(f"වාර්තාව සෑදීමේදී දෝෂයක් සිදු විය: {e}")
