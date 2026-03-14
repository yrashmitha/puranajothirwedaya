import json
import logging
import os
import re
import sys

from docx.oxml import OxmlElement, ns

from dasha_system import AstroSystem

logger = logging.getLogger("AstroApp")

# --- 2. CONFIGURATION & PROMPTS ---
FONT_NAME = 'Abhaya Libre'
API_KEY = os.environ.get("GEMINI_API_KEY", "")


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

    # --- ඒ ඒ මාතෘකාවට අදාළව අනිවාර්යයෙන්ම ඇතුළත් කළ යුතු කරුණු ---
    section_guides = {
    "පෞරුෂය": "ජන්මියාගේ සහජ හැකියාවන්, සැඟවුණු දක්ෂතා සහ දුර්වලතා විස්තර කරන්න. සමාජය ඔහුව දකින ආකාරය සහ ඔහුගේ සැබෑ ඇතුළාන්තය අතර ඇති වෙනස, කෝපය පාලනය කරගන්නා ආකාරය, සහ අන් අයව පහසුවෙන් විශ්වාස කිරීමේ පුරුද්දක් ඇත්නම් ඒ ගැන ගැඹුරින් පවසන්න.",
    
    "අධ්‍යාපනය": "ධාරණ ශක්තිය, ඉගෙනීමට ඇති උනන්දුව, තරග විභාග වලින් ජය ලැබීමේ හැකියාව ගැන සඳහන් කරන්න. අධ්‍යාපනයට බාධා ඇතිවන කාල සීමාවන්, ගැළපෙනම විෂය ධාරාවන් සහ උසස් අධ්‍යාපනයට හෝ විදේශ අධ්‍යාපනයට ඇති වාසනාව ගැන පැහැදිලි කරන්න.",
    
    "වෘත්තීය ජීවිතය සහ ආර්ථික ශක්තිය": "වඩාත්ම සාර්ථක විය හැකි වෘත්තීය ක්ෂේත්‍ර, අනුන් යටතේ කරන රැකියාවක් ද නැතිනම් ස්වයං රැකියාවක්/ව්‍යාපාරයක් ද වඩාත් සුදුසු යන්න සෘජුව කියන්න. රැකියා ස්ථානයේ ඇතිවිය හැකි රහස් සතුරු කරදර, ඊර්ෂ්‍යාවන්, විදෙස් රැකියා හෝ විදේශගත වීමේ වාසනාව, ධනය ඉපයීමේ හැකියාව, මුදල් ඉතිරි නොවන ස්වභාවයක් හෝ ණය තුරුස් වීමේ අවදානමක් ඇත්නම් ඒ ගැන පැහැදිලිව පාරිභෝගිකයාව දැනුවත් කරන්න.",
    
    "ප්‍රේමය සහ විවාහ ජීවිතය": "විවාහය ප්‍රමාද වේද නැද්ද යන්න, සහකරුගේ හෝ සහකාරියගේ ස්වභාවය සහ ගතිගුණ ගැන පවසන්න. ප්‍රේම සබඳතා බිඳවැටීමේ අවදානම්, විවාහයෙන් පසු ජීවිතය, පවුල් පසුබිම් ගැටලු සහ අදාළ ග්‍රහ දෝෂ (ඇත්නම් පමණක් බිය නොගන්වා) ගැන විස්තර කරන්න.",
    
    "දේපළ, භූමිය, නිවාස සහ වාහන භාග්‍යය": "ස්වකීය දහදිය මහන්සියෙන් ගෙවල් දොරවල් සෑදීමේ හෝ වාහන මිලදී ගැනීමේ භාග්‍යය ගැන පවසන්න. පාරම්පරික උරුමයන් ලැබේද යන්න සහ දේපළ සම්බන්ධයෙන් පවුලේ අය සමඟ නඩුහබ හෝ බාධා ඇතිවීමේ අවදානමක් ඇත්නම් ඒ ගැන පවසන්න.",
    
    "ශාරීරික සෞඛ්‍යය, මාරක අපල, හදිසි අනතුරු": "පෙළඹිය හැකි ලෙඩ රෝග (උදා: ආමාශ, ස්නායු, අස්ථි සම්බන්ධ රෝග), මානසික පීඩනයන්, සහ හදිසි අනතුරු අවදානම් ගැන කල්තියා අනතුරු අඟවන්න. බිය ගැන්වීමකින් තොරව මනුෂ්‍යයෙකු සේ ජීවිත කාලය පුරාම පරිස්සම් විය යුතු සෞඛ්‍ය පුරුදු ගැන උපදෙස් දෙන්න.",
    
    "දරු පල": "දරු පල ප්‍රමාදවීම් (ඇත්නම් පමණක්) ඉතා සංවේදීව ඉඟි කරන්න. දරුවන්ගේ අනාගත සාර්ථකත්වය, දරුවන්ගෙන් දෙමව්පියන්ට ලැබෙන සතුට සහ දරුවන්ගේ සාමාන්‍ය ස්වභාවය ගැන විස්තර කරන්න.",
    
    "මෙතෙක් දැක්වූ කරුණු අනුව ජීවන ගමනේ සමස්ත සාරාංශය": "ඉහත සියලු කරුණු කැටි කර, විදේශයක ස්ථිර පදිංචියට (PR) ඇති වාසනාව ඇතුළුව, ජීවිතයේ සාර්ථකම සහ ධනය ගලාගෙන එන ස්වර්ණමය කාල සීමාවන් මෙන්ම වඩාත්ම පරිස්සම් විය යුතු අඳුරු කාල සීමාවන් පෙන්වා දෙමින් ජීවන ගමනේ සමස්ත සාරාංශයක් ලබා දෙන්න.",
    
    "වර්තමාන දශාව අනුව පලාපල": "දැනට ගතවන මහ දශාව සහ අන්තර් දශාව අනුව, මේ මොහොතේ (වර්තමාන වර්ෂයේ) සහ ඉදිරි වසර කිහිපය තුළ අපේක්ෂා කළ හැකි සුවිශේෂී වෙනස්කම් මොනවාද යන්න සහ මේ කාලයේදී විශේෂයෙන් පරිස්සම් විය යුතු දේවල් ගැන සෘජු උපදෙස් දෙන්න.",
    
    "ජීවිතයේ අභියෝග ජයගැනීම සඳහා වූ පොදු ශාස්ත්‍රීය සහ බෞද්ධ පිළියම්": "මිථ්‍යා සහ අධික වියදම් යන ශාන්තිකර්ම බැහැර කර, මෙම කේන්ද්‍රයට ආවේණික වූ ප්‍රධානතම ග්‍රහ දෝෂ (අපල) සඳහා නිවසේදීම කළ හැකි ප්‍රායෝගික බෞද්ධ වත්පිළිවෙත්, බෝධි පූජා ක්‍රම, දානමාන සහ ජීවන රටාවේ වෙනස් කරගත යුතු පුරුදු පෙළගස්වන්න."
    }

    # අදාළ මාතෘකාවට (sec) ගැළපෙන උපදෙස තෝරාගැනීම. නැත්නම් හිස්තැනක් ගනී.
    specific_guide = section_guides.get(sec, "")

    # තෝරාගත් උපදෙස Prompt එකට දැමීමට සකස් කිරීම
    specific_prompt_text = f"**මෙම අංශය සඳහා අනිවාර්යයෙන්ම ඇතුළත් කළ යුතු කරුණු:** {specific_guide}\n" if specific_guide else ""

    task_prompt = (
        f"ඔබ දැන් විශ්ලේෂණය කළ යුත්තේ කේන්ද්‍රයේ [{sec}] යන අංශය පිළිබඳව පමනයි, මෙම විස්තර කිරීමෙදී වෙනත් කිසිදු අංශයක් ගැන විස්තර දමන්න එපා (උදා:- විවාහය ගැන කියද්දී දරු පල කියන්න් එපා ). {limit_text}\n\n"
        f"{specific_prompt_text}\n"
        "කරුණාකර පහත උපදෙස් දැඩිව පිළිපදින්න:\n"
        "1. කතාවක් මෙන් ලියන්න (Narrative Flow): 'ලග්න කේන්ද්‍රය අනුව', 'නවාංශකය අනුව', 'සුබ පල', 'අසුබ පල' ලෙස දැඩි මාතෘකා යටතේ කරුණු නොබෙදන්න. ඒ වෙනුවට එම සියලු දත්ත එකට මුසු කර, කියවීමට පහසු, ගලාගෙන යන ඡේද කිහිපයක් ලෙස ගැඹුරු විග්‍රහයක් කරන්න. කේන්ද්‍රයේ ඇති සුබ පල මෙන්ම අසුබ පල (අදාළ මාතෘකාවට අදාල ඒවා පමණක්) කිසිවක් වසන් නොකර සෘජුව සහ පැහැදිලිව සඳහන් කරන්න\n"
        "2. කිසිදු ශාන්තිකර්මයක් හෝ පිළියමක් මෙහි ඇතුළත් නොකරන්න (ඒවා වෙනම කොටසකින් ලබා දෙනු ඇත). මෙහිදී කළ යුත්තේ ශාස්ත්‍රීය විග්‍රහය පමණි.\n"
        "3. සෘජුවම කරුණට පිවිසෙන්න. හැඳින්වීම් අනවශ්‍යයි.\n"
        "4. අතිශය වැදගත්: මීට පෙර අංශ (Sections) විස්තර කිරීමේදී ඔබ භාවිතා කළ වාක්‍ය, වාක්‍ය ඛණ්ඩ හෝ අදහස් ඒ ආකාරයෙන්ම නැවත භාවිතා කිරීමෙන් සම්පූර්ණයෙන්ම වළකින්න. අදාළ මාතෘකාවට පමණක් සුවිශේෂී වූ නව කරුණු පමණක් ඉදිරිපත් කරන්න.\n"
        "5. අසුබ පල සඟවන්න එපා, නමුත් මනුෂ්‍යවාදීව පවසන්න (Honest but Empathetic): කේන්ද්‍රයේ පාප, නීච, අස්ත ග්‍රහයන් හෝ 6, 8, 12 ස්ථානවල බලපෑම් ඇත්නම්, එයින් සිදුවිය හැකි විවාහ බාධා, ලෙඩ රෝග, ධන හානි හෝ රැකියා ගැටලු වැනි අසුබ පල අනිවාර්යයෙන්ම පැහැදිලිව සඳහන් කරන්න (අදාළ මාතෘකාවට අදාල ඒවා පමණක්). ඒවා කිසිසේත් වසන් නොකරන්න. **නමුත්**, එම අසුබ පල පැවසූ වහාම, ජන්මියාගේ හිත නොකැඩෙන පරිදි කේන්ද්‍රයේ ඇති වෙනත් සුබ ග්‍රහ බලයන් හෝ ජන්මියාගේ සහජ වීර්යය පෙන්වා දී, 'මෙම අභියෝග සහ පෙර කර්ම බාධක ඔබේ නොපසුබට උත්සාහයෙන්, බුද්ධියෙන් සහ ඉවසීමෙන් සාර්ථකව මඟහරවා ගත හැකියි' යනුවෙන් සිත සනසන සහ ධෛර්යවත් කරන වචන අනිවාර්යයෙන් භාවිතා කරන්න."
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
    "temperature": 0.4,  # ශාස්ත්‍රීය කරුණු වෙනස් නොකර සාමාන්‍ය නිර්මාණශීලීත්වයක් ලබා දීමට
    "top_p": 0.8,        # ස්වාභාවික සිංහල භාෂා රටාවක් පවත්වා ගැනීමට
    "top_k": 40,         # එකම වචන නැවත නැවත ලිවීමෙන් වැළකීමට (Repetition නවතාලීමට)
}


# --- 4. UTILITY FUNCTIONS ---
def apply_style(run, size, bold=False):
    run.font.name = FONT_NAME
    run.font.size = Pt(size)
    run.font.bold = bold

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

    api_key = os.environ.get("GEMINI_API_KEY", API_KEY)
    client = genai.Client(api_key=api_key)
    records = load_birth_records()
    logger.info(f"Total records to process: {len(records)}")
    output_files = []

    for record in records:
        try:

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

            sections = [
                "පෞරුෂය",
                "අධ්‍යාපනය",
                "වෘත්තීය ජීවිතය සහ ආර්ථික ශක්තිය",
                "ප්‍රේමය සහ විවාහ ජීවිතය",
                "දේපළ, භූමිය, නිවාස සහ වාහන භාග්‍යය",
                "ශාරීරික සෞඛ්‍යය, මාරක අපල, හදිසි අනතුරු",
                "දරු පල",
                "මෙතෙක් දැක්වූ කරුණු අනුව ජීවන ගමනේ සමස්ත සාරාංශය",
                "වර්තමාන දශාව අනුව පලාපල",
                "ජීවිතයේ අභියෝග ජයගැනීම සඳහා වූ පොදු ශාස්ත්‍රීය සහ බෞද්ධ පිළියම්"
            ]

            package_type = record.get("package_type", "Normal")

            logger.info(f"package {package_type}")
            current_sections = sections.copy() # මූලික 10
            if package_type == "VIP (Rs. 1500)":
                logger.info(f">>> [LOG]: VIP customer, therefore adding mp3")
                current_sections.append("විශේෂ VIP උපදේශනය: ග්‍රහ අපල සඳහා වන සුවිශේෂී ස්තෝත්‍රය සහ වත්පිළිවෙත්")

            # 4. සාමාන්‍ය Sections වලට යැවීමට ප්‍රශ්න ඉවත් කළ JSON එකක් සෑදීම
            normal_data_for_sections = full_record.copy()
            normal_data_for_sections.pop("විශේශ_ගැටලු", None) # ප්‍රශ්න ඉවත් කළා
            normal_data_for_sections.pop("package_type", None) # ප්‍රශ්න ඉවත් කළා



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



            for index, sec in enumerate(current_sections):
                # පළමු අංශය හැර (Index 0) අනෙක් සෑම අංශයකටම පෙර අලුත් පිටුවක් එක් කරයි
                if index > 0:
                    doc.add_page_break()
                    logger.info(f"Page Break added before section: {sec}")

                logger.info(f"--- Processing Section: {sec} ---")

                # සාරාංශය (Summary/Sranshaya) නම් පමණක් වචන සීමා කිරීමේ නියෝගය දැඩි කරන්න
                if "සාරාංශය" in sec or "Summary" in sec:
                    limit_text = "වචන 400කට වඩා අඩු, ඉතා සංක්ෂිප්ත සාරාංශයක් ලබා දෙන්න. කිසිදු පිළියමක් මෙහි ඇතුළත් නොකරන්න."
                elif "දරු පල" in sec:
                    limit_text = (
                        "ලග්න කේන්ද්‍රයේ සහ නවාංශකයේ 5 වැන්න සහ 5 අධිපතිගේ 'සැබෑ බලය' සසඳා බලන්න. "
                        "අනවශ්‍ය ලෙස 'ප්‍රමාද වීම්' ගැන සඳහන් නොකර, ගුරුගේ දෘෂ්ටිය පවතී නම් එය ඉතා සුබ පලයක් ලෙස දක්වන්න. "
                        "කිසිදු ශාන්තිකර්මයක් හෝ පිළියමක් මෙහි ඇතුළත් නොකරන්න."
                    )
                elif "පොදු ශාස්ත්‍රීය සහ බෞද්ධ පිළියම්" in sec:
                    limit_text = (
                        "මෙම කේන්ද්‍රයේ ඇති ප්‍රධානතම දුර්වලතා හෝ අපල හඳුනාගෙන, මුළු ජීවිතයටම බලපාන පරිදි සවිස්තරාත්මක ශාස්ත්‍රීය සහ බෞද්ධ පිළියම් ලබා දෙන්න. \n"
                        "අනිවාර්ය ආකෘතිය (Formatting): "
                        "1. ග්‍රහයාගේ නම හෝ යෝගය අනිවාර්යයෙන්ම '###' සලකුණෙන් ආරම්භ කර ප්‍රධාන අනු-මාතෘකාවක් ලෙස දක්වන්න (උදා: ### ගුරු චණ්ඩාල යෝගය සමනය කිරීම සඳහා). "
                        "2. එම මාතෘකාවට ඉදිරියෙන් කිසිදු විටෙක Bullet points (- හෝ *) නොයොදන්න. "
                        "3. මාතෘකාව යටතේ ඇති පිළියම් (වර්ණය, මල්, දෛනික පුරුදු) පමණක් එක් පේළියකට එක බැගින් '-' සලකුණ යොදා Bullet points ලෙස ඉදිරිපත් කරන්න."
                        "ශාස්ත්‍රීය පිළියම් සහ බෞද්ධ වත්පිළිවෙත් (Remedial Measures): අසුබ පල පවතින විට හෝ පවතින ශක්තීන් වර්ධනය කර ගැනීමට පිළියම් ලබා දීමේදී, පහත සඳහන් ව්‍යුහය අනුගමනය කරමින් ඉතාමත් සවිස්තරාත්මකව කරුණු දක්වන්න:\n"
                        "ග්‍රහයාට අදාළ විශේෂිත වතාවත්: අදාළ ග්‍රහයාගේ වර්ණය සහ බලපෑම අනුව (උදා: කුජ වෙනුවෙන් රතු මල්, ශනි වෙනුවෙන් නිල් මල්) බෝධි පූජා හෝ දේව වන්දනා නිර්දේශ කරන්න.\n"
                        "ශාස්ත්‍රීය සහ භෞතික හේතු දැක්වීම: පිළියම මඟින් පුද්ගලයාගේ ජීව විද්‍යාත්මක හෝ මානසික මට්ටමට සිදුවන බලපෑම පැහැදිලි කරන්න.\n"
                        "උදාහරණ: රවි 11 වැන්නේ සිටින බැවින් සූර්ය වන්දනාවේ යෙදීමෙන් සිරුරට ලැබෙන හිරු රැස් මඟින් අලස බව දුරු වී ක්‍රියාශීලී බව වර්ධනය වේ.\n"
                        "විශේෂිත පිරිත් සහ සූත්‍ර: මානසික ඒකාග්‍රතාවය සහ ග්‍රහ අපල සමනය සඳහා බෞද්ධ දර්ශනයට අනුකූල පිරිත් (උදා: අභිසම්භිධාන පිරිත, මෝර පිරිත) නිර්දේශ කරන්න.\n"
                        "දේව ආශිර්වාදය සහ මනෝවිද්‍යාත්මක ශක්තිය: බිය හෝ මැලි බව දුරු කර ගැනීමට ගැළපෙන දේව වන්දනා (උදා: කතරගම දෙවියන්, ගණ දෙවියන්) සහ ඒවා තුළින් ලැබෙන ආත්ම විශ්වාසය විස්තර කරන්න.\n"
                        "දෛනික ප්‍රායෝගික පුරුදු: ස්නානය කරන ජලයට කහ එකතු කිරීම වැනි සරල නමුත් ශාස්ත්‍රීය පදනමක් සහිත පවිත්‍රතා ක්‍රමවේද ඇතුළත් කරන්න.\n"
                    )
                elif "විශේෂ VIP උපදේශනය" in sec:
                    limit_text = (
                        "මෙම කේන්ද්‍රයේ දැනට පවතින ප්‍රබලම ග්‍රහ අපල (උදා: සෙනසුරු ඒරාෂ්ටකය, රාහු දශාව) හඳුනාගන්න. "
                        "එම අපල දුරු කිරීම සඳහා ගායනා කළ හැකි (Rhythmic/Melodic) ස්තෝත්‍ර පද පේළි 8-16 කින් යුත් (මෙය "
                        "mp3 එකක් ලෙස ලබා දෙනු ඇත.)\n"
                        "සුවිශේෂී 'ශාන්ති ස්තෝත්‍රයක්' නිර්මාණය කරන්න.\n"
                        "පහත ආකෘතිය භාවිතා කරන්න:\n"
                        "1. ### විශේෂ ශාන්ති ස්තෝත්‍රය (පද මාලාව මෙහි දක්වන්න)\n"
                        "2. ### ස්තෝත්‍රයේ තේරුම (පේලියෙන් පේලිය අර්තය දක්වන්න)\n"
                        "4. ### භාවිතා කළ යුතු ආකාරය (දිනකට කී වතාවක්ද, වේලාව සහ වර්ණය දක්වන්න)\n"
                        "4. ස්තෝත්‍රය කියවන අතරතුර කළ යුතු 'විශේෂ මානසික ඒකායන භාවනාව' කෙටියෙන් දක්වන්න."
                    )
                else:
                    limit_text = "වෘත්තීය මට්ටමේ, ගලාගෙන යන ශාස්ත්‍රීය විග්‍රහයක් ලබා දෙන්න."

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
                birth_data_special = json.dumps(normal_data_for_sections, ensure_ascii=False)
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
                        f"{birth_data_special}\n\n"
                        f"මෙම විශේෂ ප්‍රශ්නයට සෘජු පිළිතුරක් අවශ්‍යයි: '{question}'\n\n"
                        "උපදෙස් (අනිවාර්යයෙන්ම පිළිපදින්න):\n"
                        "1. අතිශය වැදගත් (Strict Rule): පාරිභෝගිකයා අසා ඇති ගැටලුවට පමණක් සෘජුවම පිළිතුරු දෙන්න. ගැටලුවට අදාළ නැති අනෙකුත් ග්‍රහයන්, රාශි (1 සිට 12 දක්වා), පෞරුෂය, විවාහය හෝ දරු පල ආදිය කිසිසේත් විස්තර නොකරන්න.\n"
                        "2. ගැටලුවට අදාළ වන ග්‍රහ පිහිටීම් පමණක් යොදාගෙන කෙලින්ම පිළිතුර ගොඩනඟන්න (උදා: විදෙස් ගමන් ගැන ඇසුවොත් 9, 12 භාව සහ රාහු පමණක් විස්තර කිරීම).\n"
                        "3. වර්තමාන කාලය 2026 ලෙස සලකා, ඊට අදාළ දශා කාලයන් පමණක් දක්වමින් ප්‍රශ්නයට අදාළ සාර්ථකම කාලය පවසන්න.\n"
                        "4. කතා කරන භාෂාවෙන් (Conversational tone), කෙටි සහ පැහැදිලි ඡේද ලෙස ලියන්න. කිසිදු විටෙක වාක්‍ය අගට 'නේද?' යන්න නොයොදන්න.\n"
                        "5. සෑම ප්‍රධාන උප-මාතෘකාවක්ම '###' සලකුණෙන් ආරම්භ කරන්න. පිළියම් ලබා දීම අත්‍යවශ්‍ය නම් පමණක් අදාළ පේළිය ආරම්භයේ '-' සලකුණ යොදා Bullet points ලෙස ඉදිරිපත් කරන්න.\n"
                        "6. Sinhala only, 300-500 words.\n"
                        "7- ශාස්ත්‍රීය පිළියම් සහ බෞද්ධ වත්පිළිවෙත් (Remedial Measures): අසුබ පල පවතින විට හෝ පවතින ශක්තීන් වර්ධනය කර ගැනීමට පිළියම් ලබා දීමේදී, පහත සඳහන් ව්‍යුහය අනුගමනය කරමින් ඉතාමත් සවිස්තරාත්මකව කරුණු දක්වන්න\n"
                        "ග්‍රහයාට අදාළ විශේෂිත වතාවත්: අදාළ ග්‍රහයාගේ වර්ණය සහ බලපෑම අනුව (උදා: කුජ වෙනුවෙන් රතු මල්, ශනි වෙනුවෙන් නිල් මල්) බෝධි පූජා හෝ දේව වන්දනා නිර්දේශ කරන්න.\n"
                        "ශාස්ත්‍රීය සහ භෞතික හේතු දැක්වීම: පිළියම මඟින් පුද්ගලයාගේ ජීව විද්‍යාත්මක හෝ මානසික මට්ටමට සිදුවන බලපෑම පැහැදිලි කරන්න.\n"
                        "උදාහරණ: 'රවි 11 වැන්නේ සිටින බැවින් සූර්ය වන්දනාවේ යෙදීමෙන් සිරුරට ලැබෙන හිරු රැස් මඟින් අලස බව දුරු වී ක්‍රියාශීලී බව වර්ධනය වේ'.\n"
                        "විශේෂිත පිරිත් සහ සූත්‍ර: මානසික ඒකාග්‍රතාවය සහ ග්‍රහ අපල සමනය සඳහා බෞද්ධ දර්ශනයට අනුකූල පිරිත් (උදා: අභිසම්භිධාන පිරිත, මෝර පිරිත) නිර්දේශ කරන්න.\n"
                        "දේව ආශිර්වාදය සහ මනෝවිද්‍යාත්මක ශක්තිය: බිය හෝ මැලි බව දුරු කර ගැනීමට ගැළපෙන දේව වන්දනා (උදා: කතරගම දෙවියන්, ගණ දෙවියන්) සහ ඒවා තුළින් ලැබෙන ආත්ම විශ්වාසය විස්තර කරන්න.\n"
                        "දෛනික ප්‍රායෝගික පුරුදු: ස්නානය කරන ජලයට කහ එකතු කිරීම වැනි සරල නමුත් ශාස්ත්‍රීය පදනමක් සහිත පවිත්‍රතා ක්‍රමවේද ඇතුළත් කරන්න.\n"
                        "සෘජු ප්‍රවේශය:කෙලින්ම විෂය කරුණ විග්‍රහ කිරීම අරඹන්න."
                    )

                    log_prompt(record_folder, f"[SPECIAL QUESTION]\n{question}\n\n{q_prompt}")
                    q_response = None
                    for _attempt in range(1, 4):
                        q_response = chat_special.send_message(q_prompt)
                        if q_response.text is not None:
                            break
                        logger.warning(f"QUESTION {i+1}: Empty response (attempt {_attempt}/3), retrying in 10s...")
                        time.sleep(10)

                    if q_response is None or q_response.text is None:
                        logger.error(f"QUESTION {i+1}: All 3 retries returned empty response. Skipping.")
                        continue

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
            output_files.append(output_file)
            logger.info(f"SUCCESS: Report saved as {output_file}")
            time.sleep(2)

        except Exception as e:
            phone = record.get("කේන්ද්‍ර_සටහන", {}).get("දුරකතන_අංකය", "unknown")
            logger.error(f"Record [{phone}] FAILED: {str(e)}", exc_info=True)
            continue

    return output_files


if __name__ == "__main__":
    generate_report()
