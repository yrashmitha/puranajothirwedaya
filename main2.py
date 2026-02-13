import json
import logging
import re
import sys

from docx.oxml import OxmlElement, ns

from dasha_system import AstroSystem

logger = logging.getLogger("AstroApp")

# --- 2. CONFIGURATION & PROMPTS ---
FONT_NAME = 'Abhaya Libre'
API_KEY = "AIzaSyCTwLEBq6-a90LVsXfUvbxagQkGeKbzNgU"


def load_birth_records():
    with open("birth_records.json", "r", encoding="utf-8") as f:
        return json.load(f)


def sanitize_filename(text):
    return re.sub(r'[^0-9A-Za-z]', '', text)


def add_page_number(run):
    """පිටු අංකය ස්වයංක්‍රීයව පෙන්වීමට XML ක්ෂේත්‍රයක් එක් කරයි."""
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(ns.qn('w:fldCharType'), 'begin')

    instrText = OxmlElement('w:instrText')
    instrText.set(ns.qn('xml:space'), 'preserve')
    instrText.text = "PAGE"

    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(ns.qn('w:fldCharType'), 'end')

    run._r.append(fldChar1)
    run._r.append(instrText)
    run._r.append(fldChar2)


def read_file(file_path):
    """Text ගොනුවක අන්තර්ගතය කියවා ලබා දෙයි."""
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            return file.read().strip()
    except FileNotFoundError:
        logger.error(f"Error: {file_path} ගොනුව සොයාගත නොහැක.")
        return ""


def generate_section(chat, sec, limit_text, folder):
    task_prompt = (
        f"[{sec}] පිළිබඳව {limit_text}\n"
        "වර්තමාන කාලය යාවත්කාලීන කිරීම: සැමවිටම වාර්තාව සකස් කරන මොහොතේ පවතින සැබෑ වර්තමාන දිනය සහ වර්ෂය (උදා:2026) පදනම් කරගෙන කරන්න. පැරණි දත්ත මත පදනම්ව වර්තමානය තීරණය නොකරන්න. මේ 2026 අවුරුද්ද වේ.\n"
        "**විග්‍රහයේදී පහත ගැඹුරු ජ්‍යොතිෂ තර්ක අනිවාර්යයෙන් භාවිතා කරන්න:**\n"
        "1. ලග්න කේන්ද්‍ර සහ නවන්ෂක සැසඳීම"
    )

    log_prompt(folder, task_prompt)
    response = chat.send_message(task_prompt)
    log_prompt(folder, f"[AI RESPONSE]\n{response.text}")

    return response.text


SYSTEM_PROMPT = read_file("prompts/system_prompt.txt")
BIRTH_DATA = read_file("prompts/birth_data.json")

BASE_CONTEXT = f"""
{SYSTEM_PROMPT}

මෙම කේන්ද්‍ර දත්ත සම්පූර්ණ වාර්තාව සඳහා පදනම වේ:

{BIRTH_DATA}
"""

sections = [
    # "පෞරුෂය",
    # "අධ්‍යාපනය",
    "වෘත්තීය ජීවිතය සහ ආර්ථික ශක්තිය",
    # "ප්‍රේමය සහ විවාහ ජීවිතය",
    # "දේපළ, භූමිය, නිවාස සහ වාහන භාග්‍යය",
    # "ශාරීරික සෞඛ්‍යය, මාරක අපල, හදිසි අනතුරු",
    # "දරු පල",
    # "මෙතෙක් දැක්වූ කරුණු අනුව ජීවන ගමනේ සමස්ත සාරාංශය",
    # "වර්තමාන දශාව අනුව පලාපල"
]

# --- 3. HELPER FUNCTIONS ---
import time
import logging
import sys
from google import genai
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH


# --- 2. CONFIGURATION ---
MODEL_NAME = "gemini-2.5-flash"
# Deterministic and low-cost config
GEN_CONFIG = {
    "temperature": 0,  # deterministic output
    "top_p": 0.05,  # only top 5% probability tokens considered
    "top_k": 1,  # only the single most probable token is picked
}


# --- 4. UTILITY FUNCTIONS ---
def apply_style(run, size, bold=False):
    run.font.name = FONT_NAME
    run.font.size = Pt(size)
    run.font.bold = bold


import os
from datetime import datetime


def create_record_folder(base_name):
    """
    outputs/<base_name>/ folder create කරයි
    """
    folder_path = os.path.join("outputs", base_name)
    os.makedirs(folder_path, exist_ok=True)
    return folder_path


def save_birth_snapshot(folder, birth_json):
    """
    Birth data txt එක save කරයි
    """
    path = os.path.join(folder, "birth_details.txt")

    with open(path, "w", encoding="utf-8") as f:
        f.write("===== BIRTH DATA SNAPSHOT =====\n\n")
        f.write(json.dumps(birth_json, ensure_ascii=False, indent=2))

    logger.info("Birth snapshot saved")


def log_prompt(folder, text):
    """
    Executed prompts log file එකට append කරයි
    """
    path = os.path.join(folder, "prompts_log.txt")

    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    with open(path, "a", encoding="utf-8") as f:
        f.write(f"\n\n===== {timestamp} =====\n")
        f.write(text)
        f.write("\n")


# --- 5. MAIN EXECUTION ---
def generate_report():
    logger.info("Starting Horoscope Generation Process...")

    try:
        client = genai.Client(api_key=API_KEY)

        records = load_birth_records()

        for record in records:

            global BIRTH_DATA
            BIRTH_DATA = json.dumps(record["කේන්ද්‍ර_සටහන"], ensure_ascii=False)

            phone = record["කේන්ද්‍ර_සටහන"]["දුරකතන_අංකය"]
            dob = record["කේන්ද්‍ර_සටහන"]["උපන්_දිනය"]

            filename = sanitize_filename(phone) + "_" + sanitize_filename(dob)
            # ⭐ Create folder per record
            record_folder = create_record_folder(filename)
            # ⭐ Save birth snapshot
            doc = Document()

            # --- Footer සහ පිටු අංකය සැකසීම ---
            section = doc.sections[0]
            footer = section.footer
            footer_para = footer.paragraphs[0]
            footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

            # පාදකයේ පෙළ ඇතුළත් කිරීම
            run_footer = footer_para.add_run("පුරාණ ජෝතිර්වේදය හදහන් සේවය | පිටුව: ")
            apply_style(run_footer, 10, bold=False)

            # පිටු අංකය එක් කිරීම
            add_page_number(footer_para.add_run())

            # Justify format එක සඳහා අවශ්‍ය styles සැකසීම
            style = doc.styles['Normal']
            style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

            # ප්‍රධාන මාතෘකාව
            title_p = doc.add_paragraph()
            title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = title_p.add_run("නමෝ බුද්ධාය!")
            apply_style(run, 24, bold=True)

            # --- Create a single chat session for this record ---
            full_record = record["කේන්ද්‍ර_සටහන"].copy()

            # 1. NASA Processor එක හරහා දත්ත Inject කිරීම
            astro_tool = AstroSystem()
            dob = record["කේන්ද්‍ර_සටහන"]["උපන්_දිනය"]
            time_raw = record["කේන්ද්‍ර_සටහන"]["උපන්_වේලාව"]# "2004/08/26"
            nasa_data = astro_tool.get_nasa_analysis(dob, time_raw)

            # මුල් JSON එකට NASA තොරතුරු එක් කිරීම
            full_record = record["කේන්ද්‍ර_සටහන"].copy()
            if "error" not in nasa_data:
                full_record["නාසා_තාරකා_විද්‍යාත්මක_තහවුරු_කිරීම"] = nasa_data
                logger.info("NASA Dasha & Transit data successfully injected.")

            # 4. සාමාන්‍ය Sections වලට යැවීමට ප්‍රශ්න ඉවත් කළ JSON එකක් සෑදීම
            normal_data_for_sections = full_record.copy()
            normal_data_for_sections.pop("විශේශ_ගැටලු", None) # ප්‍රශ්න ඉවත් කළා

            birth_data_for_ai = json.dumps(normal_data_for_sections, ensure_ascii=False)
            save_birth_snapshot(record_folder, full_record)
            logger.info("Clean Birth snapshot saved with NASA data.")

            chat = client.chats.create(
                model=MODEL_NAME,
                config={
                    "system_instruction": SYSTEM_PROMPT + "\n" + birth_data_for_ai,
                    **GEN_CONFIG
                }
            )

            for index, sec in enumerate(sections):
                # පළමු අංශය හැර (Index 0) අනෙක් සෑම අංශයකටම පෙර අලුත් පිටුවක් එක් කරයි
                if index > 0:
                    doc.add_page_break()
                    logger.info(f"Page Break added before section: {sec}")

                logger.info(f"--- Processing Section: {sec} ---")

                # සාරාංශය (Summary/Sranshaya) නම් පමණක් වචන සීමා කිරීමේ නියෝගය දැඩි කරන්න
                if "සාරාංශය" in sec or "Summary" in sec:
                    limit_text = "වචන 800කට වඩා අඩු, ඉතා සංක්ෂිප්ත විග්‍රහයක් ලබා දෙන්න. පිළියම් අවශ්‍ය නැත."
                else:
                    limit_text = "වෘත්තීය මට්ටමේ විග්‍රහයක් ලබා දෙන්න."

                content = generate_section(chat, sec, limit_text, record_folder)

                # Section Header
                heading_p = doc.add_paragraph()
                run = heading_p.add_run(f"{sec}")
                apply_style(run, 18, bold=True)
                heading_p.alignment = WD_ALIGN_PARAGRAPH.LEFT  # මාතෘකා Justify කරන්නේ නැත

                # --- පිළිතුර පිරිසිදු කිරීම සහ Formatting Logic එක ---
                # 1. ඡේද එකට ඇලීම වැළැක්වීමට මුලින්ම Double Newline මගින් වෙන් කරමු
                # ඔබගේ saved instructions අනුව වඩා හොඳ ලොග් සටහන් (Better Logs) මෙහි ඇතුළත් වේ.
                content_clean = content.replace('\n\n', '[[NEW_PARA]]')
                rsponse_sections = content_clean.split('[[NEW_PARA]]')

                logger.info(f">>> [LOG]: Starting formatting for {len(rsponse_sections)} main sections/paragraphs.")

                for section in rsponse_sections:
                    section = section.strip()
                    if not section: continue

                    # 2. සෑම ඡේදයක්/කොටසක් ඇතුළත ඇති පේළි පිරික්සීම
                    # මෙහිදී 'content' වෙනුවට 'section' භාවිතා කළ යුතුයි (පරණ කේතයේ දෝෂය නිවැරදි කරන ලදී)
                    lines = section.split('\n')

                    for line in lines:
                        line = line.strip()
                        if not line: continue

                        # --- (A) ප්‍රධාන අනු-මාතෘකා (###) සැකසීම ---
                        if line.startswith('###'):
                            p = doc.add_paragraph()
                            clean_text = line.replace('###', '').replace('*', '').strip()
                            run = p.add_run(clean_text)
                            apply_style(run, 16, bold=True)
                            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                            p.space_before = Pt(12)
                            p.space_after = Pt(6)
                            continue  # මාතෘකාවක් නම් පහළ තීරණ වලට නොගොස් ඊළඟ පේළියට යයි

                        # --- (B) Bullet Point හෝ අංකිත ලැයිස්තු (1. 2. 3.) හඳුනා ගැනීම ---
                        is_bullet = line.startswith('* ') or line.startswith('- ')
                        is_numbered = line[0].isdigit() and (
                                len(line) > 1 and (line[1] == '.' or line[1] == ')')) if len(line) >= 2 else False

                        if is_bullet or is_numbered:
                            # Bullet point හෝ Numbered style එක භාවිතා කරයි
                            p = doc.add_paragraph(style='List Bullet' if is_bullet else 'List Number')
                            # Numbered ඒවා පමණක් වම් පසට (Left) පෙළගැස්වීම
                            if is_numbered:
                                p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT

                            if is_bullet:
                                line = line[2:].strip()
                            # Numbered නම් අංකය පවතින සේ තබා ගනී (නැතිනම් Gemini ගේ අංකය මැකී යයි)
                        else:
                            p = doc.add_paragraph()
                            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

                        p.paragraph_format.space_after = Pt(8)

                        # --- (C) වැදගත්ම කොටස: ඕනෑම පේළියක ඇති Bold වචන සැකසීම ---
                        # මෙහි 'elif' නැති නිසා Bullet points සහ Numbered ලැයිස්තු ඇතුළත ඇති **Bold** ද නිවැරදිව ක්‍රියා කරයි.
                        if '**' in line or '*' in line:
                            # මුලින්ම ** සලකුණු පිරික්සා පසුව * පිරික්සයි
                            marker = '**' if '**' in line else '*'
                            parts = line.split(marker)
                            for i, part in enumerate(parts):
                                # ඔත්තේ සංඛ්‍යාත කොටස් Bold වේ
                                is_bold_part = (i % 2 != 0)
                                run = p.add_run(part)
                                apply_style(run, 12, bold=is_bold_part)
                        else:
                            run = p.add_run(line)
                            apply_style(run, 12, bold=False)

                logger.info(">>> [LOG]: Line-by-line formatting and bold processing completed successfully.")

            # -------------------------------------------------------
            # -------------------------------------------------------
            # විශේෂ ප්‍රශ්න සහ විසඳුම් (Custom Solutions) කොටස
            # -------------------------------------------------------
            questions = record["කේන්ද්‍ර_සටහන"].get("විශේශ_ගැටලු", [])

            if questions:
                # Original BIRTH_DATA with special questions
                birth_data_special = json.dumps(full_record, ensure_ascii=False)
                chat_special = client.chats.create(
                    model=MODEL_NAME,
                    config={
                        "system_instruction": SYSTEM_PROMPT + "\n" + birth_data_special,
                        **GEN_CONFIG
                    }
                )
                doc.add_page_break()
                logger.info(">>> CUSTOM QUESTIONS DETECTED: Formatting with new rules...")

                # ප්‍රධාන මාතෘකාව
                q_header = doc.add_paragraph()
                q_header.alignment = WD_ALIGN_PARAGRAPH.LEFT
                run = q_header.add_run("විශේෂ උපදේශනය සහ විසඳුම් සේවාව (2026 සිට ඉදිරියට)")
                apply_style(run, 18, bold=True)

                for i, question in enumerate(questions):
                    question = question.strip()
                    if not question:
                        continue

                    start_q_time = time.time()
                    logger.info(f"QUESTION {i + 1} START: Processing -> {question}")

                    # 1. Gemini හට ලබා දෙන නව උපදෙස් මාලාව (Prompt)
                    q_prompt = (
                        f"{SYSTEM_PROMPT}\n\n"
                        f"{BIRTH_DATA}\n\n"
                        f"මෙම විශේෂ ප්‍රශ්නයට පිළිතුරු අවශ්‍යයි. හරවත් සත්‍ය පිලිතුරක් ලබා දෙන්න. : '{question}'\n\n"
                        "උපදෙස්:\n"
                        "1. වර්තමාන කාලය යාවත්කාලීන කිරීම: සැමවිටම වාර්තාව සකස් කරන මොහොතේ පවතින සැබෑ වර්තමාන දිනය සහ වර්ෂය (උදා:2026) පදනම් කරගෙන වර්තමාන ග්‍රහ ගෝචරය සහ දශා පල විග්‍රහ කරන්න. පැරණි දත්ත මත පදනම්ව වර්තමානය තීරණය නොකරන්න. මේ 2026 අවුරුද්ද වේ.\n"
                        "2. ඡේද සැකසුම: සෑම ප්‍රධාන උප-මාතෘකාවක්ම '###' සලකුණෙන් ආරම්භ කරන්න.\n"
                        "3. පිළියම් Bullet points එක එක් පේළියෙන් '-' සලකුණ යොදා ඉදිරිපත් කරන්න.\n"
                        "4. සෑම ඡේදයක් අවසානයේදීම Double Enter එකක් තබන්න.\n"
                        "5. Sinhala only, 600-800 words."
                    )

                    log_prompt(record_folder, f"[SPECIAL QUESTION]\n{question}\n\n{q_prompt}")
                    q_response = chat_special.send_message(q_prompt)
                    log_prompt(record_folder, f"[AI RESPONSE]\n{q_response.text}")

                    raw_ans = q_response.text.strip()

                    # --- ප්‍රශ්නය Word ගොනුවට ලිවීම ---
                    p_q = doc.add_paragraph()
                    p_q.space_before = Pt(20)
                    run_q = p_q.add_run(f"ගැටලුව {i + 1}: {question}")
                    apply_style(run_q, 13, bold=True)

                    # 2. පිළියම් සහ ඡේද නිවැරදිව වෙන් කරන Logic එක
                    # සැබෑ ඡේද වෙන් කිරීම
                    clean_content = raw_ans.replace('\n\n', '[[NEW_PARA]]')
                    ans_lines = clean_content.split('[[NEW_PARA]]')

                    for a_line in ans_lines:
                        a_line = a_line.strip()
                        if not a_line: continue

                        # --- (A) මාතෘකා (###) ---
                        if a_line.startswith('###'):
                            p_a = doc.add_paragraph()
                            clean_header = a_line.replace('###', '').replace('*', '').strip()
                            run_a = p_a.add_run(clean_header)
                            apply_style(run_a, 14, bold=True)
                            p_a.alignment = WD_ALIGN_PARAGRAPH.LEFT
                            p_a.space_before = Pt(12)

                        # --- (B) පිළියම් (Bullet Points) ---
                        elif a_line.startswith('-'):
                            # එක් පේළියක පිළියම් කිහිපයක් තිබේ නම් ඒවා වෙන් කරමු
                            sub_points = a_line.split('\n')
                            for point in sub_points:
                                if not point.strip(): continue
                                p_p = doc.add_paragraph()
                                p_p.style = 'List Bullet'
                                # තරු සහ ඉරි ලකුණු පිරිසිදු කිරීම
                                clean_point = point.lstrip('- ').replace('*', '').strip()
                                run_p = p_p.add_run(clean_point)
                                apply_style(run_p, 12)
                                p_p.paragraph_format.space_after = Pt(6)

                        # --- (C) සාමාන්‍ය විස්තරය ---
                        else:
                            p_a = doc.add_paragraph()
                            p_a.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                            # Bold කිරීමේ Logic එක
                            current_text = a_line.replace('**', '*')
                            if '*' in current_text:
                                parts = current_text.split('*')
                                for idx, part in enumerate(parts):
                                    run = p_a.add_run(part)
                                    apply_style(run, 12, bold=(idx % 2 != 0))
                            else:
                                run_a = p_a.add_run(current_text)
                                apply_style(run_a, 12)

                    logger.info(f"QUESTION {i + 1} COMPLETED: Formatted perfectly.")
                    doc.add_paragraph("-" * 55).alignment = WD_ALIGN_PARAGRAPH.CENTER

            # -------------------------------------------------------
            # 2. විශේෂ ශාස්ත්‍රීය සටහන (Final Special Page)
            # -------------------------------------------------------
            special_note_content = read_file("prompts/special_note.txt")

            if special_note_content:
                doc.add_page_break()
                logger.info("විශේෂ ශාස්ත්‍රීය සටහන පිටුව එක් කරමින්...")

                lines = special_note_content.split('\n')
                for line in lines:
                    line = line.strip()
                    if not line:
                        doc.add_paragraph()  # හිස් පේළි සඳහා
                        continue

                    p = doc.add_paragraph()

                    # ප්‍රධාන මාතෘකාව (විශේෂ ශාස්ත්‍රීය සටහන සහ ප්‍රකාශය)
                    if "විශේෂ ශාස්ත්‍රීය සටහන" in line:
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        run = p.add_run(line)
                        apply_style(run, 18, bold=True)
                        p.space_after = Pt(12)

                    # අනු මාතෘකා (උදා: දේශය, කාලය සහ පෞරුෂය)
                    elif line in ["දේශය, කාලය සහ පෞරුෂය", "හවුල් කර්මය සහ සහකරුගේ බලපෑම",
                                  "ග්‍රහයන්ගේ අංශකමය බලය සහ දශා කාල", "පුරුෂාර්ථය සහ වීර්යය", "අවසාන නිගමනය"]:
                        run = p.add_run(line)
                        apply_style(run, 14, bold=True)
                        p.space_before = Pt(10)

                    # ආශිර්වාදය (අවසාන පේළිය)
                    elif "තෙරුවන් සරණින්" in line:
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        p.space_before = Pt(20)
                        run = p.add_run(line)
                        apply_style(run, 13, bold=True)

                    # සාමාන්‍ය විස්තරය
                    else:
                        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                        run = p.add_run(line)
                        apply_style(run, 12)

            # -------------------------------------------------------
            # 3. අවසාන Signature කොටස
            # -------------------------------------------------------
            contact_p = doc.add_paragraph()
            contact_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            contact_p.space_before = Pt(30)

            run = contact_p.add_run("මෙයට,\nපුරාණ ජෝතිර්වේදය හදහන් සේවය")
            apply_style(run, 12, bold=True)

            # අවසාන සුරැකීම
            output_file = os.path.join(record_folder, f"{filename}.docx")

            doc.save(output_file)
            logger.info(f"SUCCESS: Report saved as {output_file}")
            time.sleep(2)

    except Exception as e:
        logger.error(f"Error: {str(e)}", exc_info=True)


if __name__ == "__main__":
    generate_report()
