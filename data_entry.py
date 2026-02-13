import logging
import sys

import streamlit as st
import json
import os
from datetime import datetime
import io  # මෙය මුලට එකතු කරන්න

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
            report_data = generate_report()

            st.success("වාර්තාව සාර්ථකව සකස් කළා! ✅")


        except Exception as e:
            st.error(f"වාර්තාව සෑදීමේදී දෝෂයක් සිදු විය: {e}")
