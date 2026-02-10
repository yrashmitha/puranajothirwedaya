import streamlit as st
import json
import os
import logging
from google import genai
from docx import Document

# --- 1. LOGGING SETUP ---
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [UI-INFO] %(message)s',
    handlers=[logging.FileHandler("chat_sync.log", encoding='utf-8'), logging.StreamHandler()]
)
logger = logging.getLogger(__name__)

API_KEY = "AIzaSyCTwLEBq6-a90LVsXfUvbxagQkGeKbzNgU"
CHATS_DIR = "saved_chats"
if not os.path.exists(CHATS_DIR): os.makedirs(CHATS_DIR)

st.set_page_config(page_title="Astrology Chat UI", layout="wide")

def save_chat(filename, history):
    with open(os.path.join(CHATS_DIR, filename), 'w', encoding='utf-8') as f:
        json.dump(history, f, ensure_ascii=False, indent=4)
    logger.info(f"Chat synced from UI to file: {filename}")

# --- SIDEBAR ---
st.sidebar.title("Rooms 🏠")
existing_files = [f for f in os.listdir(CHATS_DIR) if f.endswith('.json')]
selected_file = st.sidebar.selectbox("Chat එකක් තෝරන්න", ["New Chat"] + existing_files)

if "messages" not in st.session_state:
    st.session_state.messages = []

if st.sidebar.button("Load/Refresh Chat"):
    if selected_file != "New Chat":
        with open(os.path.join(CHATS_DIR, selected_file), 'r', encoding='utf-8') as f:
            st.session_state.messages = json.load(f)
        logger.info(f"UI loaded existing session: {selected_file}")

# --- CHAT INTERFACE ---
for msg in st.session_state.messages:
    with st.chat_message(msg["role"]): st.markdown(msg["content"])

if prompt := st.chat_input("පණිවිඩය මෙහි ටයිප් කරන්න..."):
    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user"): st.markdown(prompt)

    client = genai.Client(api_key=API_KEY)
    response = client.models.generate_content(model="gemini-2.0-flash", contents=prompt)

    with st.chat_message("assistant"): st.markdown(response.text)
    st.session_state.messages.append({"role": "assistant", "content": response.text})

    fname = selected_file if selected_file != "New Chat" else "chat_room_1.json"
    save_chat(fname, st.session_state.messages)