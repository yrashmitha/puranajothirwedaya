import logging
import sys

from docx.oxml import OxmlElement, ns

# --- 1. DETAILED LOGGING SETUP ---
# Logs will show in the terminal and save to 'horoscope_debug.log'
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s]: %(message)s',
    handlers=[
        logging.FileHandler("horoscope_debug.log", encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

# --- 2. CONFIGURATION & PROMPTS ---
FONT_NAME = 'Abhaya Libre'
API_KEY = "AIzaSyCTwLEBq6-a90LVsXfUvbxagQkGeKbzNgU"


GEN_CONFIG = {
    "temperature": 0.4,  # ශාස්ත්‍රීය කරුණු වෙනස් නොකර සාමාන්‍ය නිර්මාණශීලීත්වයක් ලබා දීමට
    "top_p": 0.8,        # ස්වාභාවික සිංහල භාෂා රටාවක් පවත්වා ගැනීමට
    "top_k": 40,         # එකම වචන නැවත නැවත ලිවීමෙන් වැළකීමට (Repetition නවතාලීමට)
}


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


SYSTEM_PROMPT = read_file("prompts/match_making.txt")
BIRTH_DATA = read_file("prompts/couple_birth_data.txt")


section_guides = {
    "දෙදෙනාගේ ජන්ම පත්‍රයන්හි ග්‍රහ ගැළපීම පිළිබඳ මූලික විග්‍රහය": 
        "මෙය ආරම්භක හැඳින්වීමයි. දෙදෙනාගේ ලග්න, රාශි සහ මූලික ග්‍රහ පිහිටීම් අනුව පොදුවේ කොතරම් දුරට ගැළපීමක් ඇත්දැයි මිත්‍රශීලීව පවසන්න. දෙදෙනාගේ එකතුවීම පිටුපස ඇති දෛවෝපගත ස්වභාවය (Karmic connection) ගැන කෙටියෙන් අදහස් දක්වන්න.",
        
    "පෞරුෂත්වයන්ගේ සමානකම්, දුර්වලතා සහ එකිනෙකාට අනුපූරක වන චරිත ලක්ෂණ": 
        "දෙදෙනාගේ චරිත වල ඇති සමානකම් සහ වෙනස්කම් විස්තර කරන්න. එක් අයෙකුගේ දුර්වලතාවයක් (උදා: ඉක්මන් කේන්තිය) අනෙක් කෙනාගේ ශක්තියකින් (උදා: ඉවසීම) පියවෙන ආකාරය (Complementary traits) ඉතා පැහැදිලිව පෙන්වා දෙන්න.",
        
    "අදහස් හුවමාරුවේ ගුණාත්මක බව සහ බුද්ධිමය මට්ටමින් ඇති අවබෝධය": 
        "දෙදෙනා අතර කතාබහ සහ අදහස් හුවමාරු වන ආකාරය විස්තර කරන්න. ප්‍රශ්නයක් ආ විට සාකච්ඡා කර විසඳාගන්නවාද නැතිනම් තර්ක කරනවාද යන්න පවසන්න. වචන නොමැතිව වුවද එකිනෙකාගේ සිතුම් පැතුම් තේරුම් ගැනීමේ මානසික බැඳීමක් ඇත්දැයි බුධ සහ සඳු ග්‍රහයින්ට අනුව විස්තර කරන්න.",
        
    "ආදරය, සෙනෙහස සහ දෙදෙනා අතර ඇතිවන චිත්තවේගීය බැඳීමේ ස්වභාවය": 
        "සිකුරු සහ කුජ ග්‍රහයින්ගේ පිහිටීම අනුව දෙදෙනා අතර ඇති ආකර්ෂණය, ආදරය ප්‍රකාශ කරන ආකාරය සහ චිත්තවේගීය ආරක්ෂාව (Emotional security) ගැන කතා කරන්න. සහකරුවාගෙන්/සහකාරියගෙන් බලාපොරොත්තු වන ආදරයේ ස්වභාවය එකිනෙකාට ගැලපේදැයි පෙන්වා දෙන්න.",
        
    "ශාරීරික ආකර්ෂණය, ලිංගික ගැළපීම සහ රහස්‍ය බැඳීම්වල ස්වභාවය":
        "සිකුරු සහ කුජ ග්‍රහයින්ගේ සංයෝග, අටවැන්න සහ දොළොස්වැන්න (ශයන සුඛය) යන ස්ථාන පදනම් කරගනිමින් දෙදෙනා අතර පවතින ශාරීරික ආකර්ෂණය සහ ලිංගික ගැළපීම පිළිබඳව විද්‍යාත්මකව සහ ශාස්ත්‍රීයව විග්‍රහ කරන්න. දෙදෙනාගේ ලිංගික අවශ්‍යතා, ආශාවන් සහ එම බැඳීමේ තෘප්තිමත්භාවය එකිනෙකාට කොතරම් දුරට ගැළපේද යන්න මෙන්ම, යම් දුර්වලතාවක් ඇත්නම් එයද කිසිවක් වසන් නොකර ඉතා සංවේදීව සහ පැහැදිලිව විස්තර කරන්න.",
        
    "මතභේද ඇතිවීමට ඇති ප්‍රවණතාවය සහ ඒවා ආදරයෙන් සමනය කරගැනීමේ මනෝවිද්‍යාත්මක මඟපෙන්වීම": 
        "කිසිදු සම්බන්ධයක් 100% ක් පරිපූර්ණ නොවන බව පවසමින්, මොවුන් දෙදෙනා අතර ගැටලු ඇතිවිය හැක්කේ කුමන කාරණා මුල්කරගෙනද (උදා: මුදල්, ඥාතීන්, ඊර්ෂ්‍යාව, හෝ කාර්යබහුල බව) යන්න පෙන්වා දෙන්න. එම තත්ත්වයන් වළක්වා ගැනීමට දෙදෙනාටම වෙන වෙනම මනෝවිද්‍යාත්මක උපදෙස් ලබා දෙන්න.",
        
    "වර්තමාන දශා කාලයන් සහ ග්‍රහ ගෝචරය (2026) සබඳතාවයට බලපාන අයුරු": 
        "දැනට වර්තමාන කාලයේ (2026 වර්ෂය) දෙදෙනාටම ගතවන දශා සහ ගෝචර ග්‍රහයන් අනුව ඔවුන්ගේ මානසිකත්වය පවතින ආකාරය විස්තර කරන්න. රැකියාවේ හෝ ආර්ථිකයේ පීඩනයන් සබඳතාවයට බලපාන්නේ නම්, මේ කාලය තුළ එකිනෙකාට සහයෝගය දැක්විය යුතු ආකාරය ගැන සෘජුව උපදෙස් දෙන්න.",

    "දරු පල වාසනාව සහ පරම්පරාවේ පැවැත්ම පිළිබඳ සෘජු විග්‍රහය":
        "දෙදෙනාගේම ජන්ම පත්‍රවල 5 වැන්න (පුත්‍රස්ථානය), 5 අධිපති ග්‍රහයා සහ දරු පල පිළිබඳ ප්‍රධාන කාරක ග්‍රහයා වන ගුරු (බ්‍රහස්පති) සිටින ආකාරය අනුව දරු පල ලැබීමේ නියත සම්භාවිතාව පරීක්ෂා කරන්න. මෙහිදී දරු පල ප්‍රමාද වීම්, දරු පල අහිමි වීම් හෝ පුත්‍ර දෝෂ වැනි තත්ත්වයන් පවතී නම්, ඒවා කිසිවක් වසන් නොකර සෘජුව ප්‍රකාශ කරන්න. විශේෂයෙන්ම කාන්තා කේන්ද්‍රයේ ගැබ්ගෙල සහ ගර්භාෂය නියෝජනය වන ස්ථානවල ග්‍රහ පීඩිත වීම් පවතීද යන්නත්, පිරිමි කේන්ද්‍රයේ ජීව ශක්තිය පිළිබඳ පවතින ග්‍රහ බලපෑමත් අනුව දරු පල වාසනාව පිළිබඳ අවසන් තීන්දුව ශාස්ත්‍රීයව ඉදිරිපත් කරන්න.",    
    
    "සබඳතාවයේ අනාගත පැවැත්ම සහ අර්බුදකාරී අවස්ථාවකදී නොබිඳී ඉදිරියට යාමේ සම්භාවිතාව": 
        "අනාගතයේදී පැමිණිය හැකි බාධක හමුවේ මෙම සබඳතාවය බිඳවැටේද නැතිනම් දෙදෙනා එක්ව එය ජයගනීද යන්න කේන්ද්‍රයේ 7 සහ 8 භාවයන්ට අනුව පැහැදිලි කරන්න. විවාහයකට එළඹීමට තරම් ස්ථාවරත්වයක් ඇත්දැයි පවසන්න.",
        
    "ප්‍රබල ග්‍රහ දෝෂ, ඒවායේ බලපෑම සහ දෙදෙනාගේ කේන්ද්‍ර මඟින් දෝෂ භංග වීම": 
        "කුජ දෝෂය, ශනි මංගල දෝෂය, සර්ප දෝෂ වැනි ප්‍රබල විවාහ දෝෂ එක් අයෙකුගේ ජන්ම පත්‍රයේ පවතී නම්, ඒවා අනෙක් ජන්ම පත්‍රයේ ග්‍රහ පිහිටීම් මඟින් සමනය වී (භංග වී) ඇත්දැයි ගැඹුරින් විශ්ලේෂණය කරන්න. දෝෂ භංග වී ඇත්නම් එය විවාහයට සුබදායක වන අයුරුත්, භංග වී නොමැති නම් අනාගතයේදී ඇතිවිය හැකි බාධාත් ඉතා පැහැදිලිව සහ සෘජුව දක්වන්න.",
        
    "සබඳතාවයේ බාධා අවම කරගැනීම සඳහා වූ ප්‍රායෝගික සහ බෞද්ධ පිළියම්": 
        "මෙම සබඳතාවය තවත් ශක්තිමත් කරගැනීමට සහ පවතින ග්‍රහ අපල (ඇත්නම්) මඟහරවා ගැනීමට දෙදෙනාටම එක්ව කළ හැකි ප්‍රායෝගික පුරුදු, සන්නිවේදන පිළියම් සහ සරල බෞද්ධ වත්පිළිවෙත් (උදා: මෙත්තා භාවනාව, එකට පන්සල් යාම වැනි දේ) පෙළගස්වන්න.",
        
    "ජීවන ගමනේ ඉදිරි පියවර පිළිබඳ සමස්ත සාරාංශය සහ අවසන් නිගමනය": 
        "ඉහත සියලු කරුණු කැටි කර, මෙම සම්බන්ධතාවය ශාස්ත්‍රීය වශයෙන් අනුමත කළ හැකිද නැද්ද යන්න සෘජුව සහ ඉතා පැහැදිලිව ප්‍රකාශ කරන්න. අනුමත කරන්නේ නම් ඊට හේතු වන ප්‍රබලම ග්‍රහ ගැළපීම් මොනවාද යන්නත්, අනුමත කළ නොහැකි නම් හෝ ඉදිරියට යාම දුෂ්කර නම් ඊට හේතු වන ප්‍රධානතම ග්‍රහ දෝෂ සහ නොගැළපීම් මොනවාද යන්නත් තර්කානුකූලව හේතු දක්වමින් විස්තර කරන්න. යම් හෙයකින් අනුමත කළ නොහැකි මට්ටමේ නොගැළපීමක් ඇත්නම් එය ඉතා කාරුණිකව පෙන්වා දෙන්න. අවසන් තීරණය ගැනීමේ පූර්ණ අයිතිය ඔවුන් සතු බව මතක් කරමින්, දෙදෙනාගේම අනාගතය වෙනුවෙන් සුබ පැතුමක් ද එක් කරන්න."
}

# --- 3. HELPER FUNCTIONS ---
import time
import logging
import sys
from google import genai
from google.genai import errors as genai_errors
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

# --- 1. DETAILED LOGGING SETUP ---
# කොන්සෝලය සහ 'app_debug.log' යන ගොනු දෙකටම විස්තරාත්මකව logs ලබා දේ.
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    handlers=[
        logging.FileHandler("app_debug.log", encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

# --- 2. CONFIGURATION ---
MODEL_NAME = "gemini-2.5-flash"


# --- 4. UTILITY FUNCTIONS ---
def apply_style(run, size, bold=False):
    run.font.name = FONT_NAME
    run.font.size = Pt(size)
    run.font.bold = bold


def send_with_retry(chat, prompt, max_retries=3, base_delay=10):
    for attempt in range(1, max_retries + 1):
        try:
            return chat.send_message(prompt)
        except genai_errors.ServerError as e:
            if attempt == max_retries:
                logger.error(f"All {max_retries} retries exhausted. Last error: {e}")
                raise
            wait = base_delay * (2 ** (attempt - 1))
            logger.warning(f"Server error (attempt {attempt}/{max_retries}): {e}. Retrying in {wait}s...")
            time.sleep(wait)


# --- 5. MAIN EXECUTION ---
def generate_report():
    logger.info("Starting Horoscope Generation Process...")

    try:
        client = genai.Client(api_key=API_KEY)
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

        # Start Chat Session
        logger.info(f"Initializing Chat with System Instruction.")
        

        chat = client.chats.create(
                model=MODEL_NAME,
                config={
                    "system_instruction": SYSTEM_PROMPT + "\n" + BIRTH_DATA,
                    **GEN_CONFIG
                }
            )
        logger.info("Context successfully registered.")

        # ප්‍රධාන මාතෘකාව
        title_p = doc.add_paragraph()
        title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = title_p.add_run("නමෝ බුද්ධාය!")
        apply_style(run, 24, bold=True)

        for index, sec in enumerate(section_guides):
            # පළමු අංශය හැර (Index 0) අනෙක් සෑම අංශයකටම පෙර අලුත් පිටුවක් එක් කරයි
            if index > 0:
                doc.add_page_break()
                logger.info(f"Page Break added before section: {sec}")

            logger.info(f"--- Processing Section: {sec} ---")

            # සාරාංශය (Summary/Sranshaya) නම් පමණක් වචන සීමා කිරීමේ නියෝගය දැඩි කරන්න
            if "සාරාංශය" in sec or "නිගමනය" in sec:
                limit_text = "වචන 600කට වඩා අඩු, ඉතා සංක්ෂිප්ත සහ සෘජු විග්‍රහයක් ලබා දෙන්න."
            else:
                limit_text = "වෘත්තීය මට්ටමේ සවිස්තරාත්මක විග්‍රහයක් ලබා දෙන්න."

            # අප විසින් කලින් සකස් කළ section_guides මඟින් අදාළ උපදෙස ලබා ගැනීම
            sec_instruction = section_guides.get(sec, "මෙම මාතෘකාව යටතේ සවිස්තරාත්මකව කරුණු දක්වන්න.")

            # >>> යාවත්කාලීන කළ නව Task Prompt එක <<<
            task_prompt = (
                f"මාතෘකාව: [{sec}]\n"
                f"මෙම මාතෘකාව ලිවීම සඳහා විශේෂ උපදෙස්: {sec_instruction}\n\n"
                f"කරුණාකර පහත උපදෙස් දැඩිව පිළිපදින්න:\n"
                f"1. කතාවක් මෙන් ලියන්න (Narrative Flow): 'ලග්න කේන්ද්‍රය අනුව', 'නවාංශකය අනුව', 'සුබ පල', 'අසුබ පල' ලෙස දැඩි මාතෘකා යටතේ කරුණු නොබෙදන්න. ඒ වෙනුවට එම සියලු දත්ත එකට මුසු කර, කියවීමට පහසු, ගලාගෙන යන ඡේද කිහිපයක් ලෙස ගැඹුරු විග්‍රහයක් කරන්න. කේන්ද්‍රයේ ඇති සුබ පල මෙන්ම අසුබ පල (අදාළ මාතෘකාවට අදාල ඒවා පමණක්) කිසිවක් වසන් නොකර සෘජුව සහ පැහැදිලිව සඳහන් කරන්න, වචන 300-500 පමණ.\n "
                f"2. කිසිදු ශාන්තිකර්මයක් හෝ පිළියමක් මෙහි ඇතුළත් නොකරන්න (ඒවා වෙනම කොටසකින් ලබා දෙනු ඇත). මෙහිදී කළ යුත්තේ ශාස්ත්‍රීය විග්‍රහය පමණි.\n"
                f"3. සෘජුවම කරුණට පිවිසෙන්න. හැඳින්වීම් අනවශ්‍යයි.\n"
                f"4. අතිශය වැදගත්: මීට පෙර අංශ (Sections) විස්තර කිරීමේදී ඔබ භාවිතා කළ වාක්‍ය, වාක්‍ය ඛණ්ඩ හෝ අදහස් ඒ ආකාරයෙන්ම නැවත භාවිතා කිරීමෙන් සම්පූර්ණයෙන්ම වලකනව.අදහසටපමණකටසංකලපලකරනව.\n"
                f"5. අසුබ පල සඟවන්න එපා, නමුත් මනුෂ්‍යවාදීව පවසන්න (Honest but Empathetic): කේන්ද්‍රයේ පාප, නීච, අස්ත ග්‍රහයන් හෝ 6, 8, 12 ස්ථානවල බලපෑම් ඇත්නම්, එයින් සිදුවිය හැකි විවාහ බාධා, ලෙඩ රෝග, ධන හානි හෝ රැකියා ගැටලු වැනි අසුබ පල අනිවාර්යයෙන්ම පැහැදිලිව සඳහන් කරන්න (අදාළ මාතෘකාවට අදාල ඒවා පමණක්). ඒවා කිසිසේත් වසන් නොකරන්න. **නමුත්**, එම අසුබ පල පැවසූ වහාම, ජන්මියාගේ හිත නොකැඩෙන පරිදි කේන්ද්‍රයේ ඇති වෙනත් සුබ ග්‍රහ බලයන් හෝ ජන්මියාගේ සහජ වීර්යය පෙන්වා දී, 'මෙම අභියෝග සහ පෙර කර්ම බාධක ඔබේ නොපසුබට උත්සාහයෙන්, බුද්ධියෙන් සහ ඉවසීමෙන් සාර්ථකව මඟහරවා ගත හැකියි' යනුවෙන් සිත සනසන සහ ධෛර්යවත් කරන වචන අනිවාර්යයෙන් භාවිතා කරන්න."
            )
            

            response = send_with_retry(chat, task_prompt)
            content = response.text

            # Section Header
            heading_p = doc.add_paragraph()
            run = heading_p.add_run(f"{sec}")
            apply_style(run, 18, bold=True)
            heading_p.alignment = WD_ALIGN_PARAGRAPH.LEFT  # මාතෘකා Justify කරන්නේ නැත

            # --- පිළිතුර පිරිසිදු කිරීම සහ Formatting Logic එක ---
            content_clean = content.replace('\n\n', '[[NEW_PARA]]')
            rsponse_sections = content_clean.split('[[NEW_PARA]]')

            logger.info(f">>> [LOG]: Starting formatting for {len(rsponse_sections)} main sections/paragraphs.")

            for section_text in rsponse_sections:
                section_text = section_text.strip()
                if not section_text: continue

                lines = section_text.split('\n')

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
                        continue 

                    # --- (B) Bullet Point හෝ අංකිත ලැයිස්තු (1. 2. 3.) හඳුනා ගැනීම ---
                    is_bullet = line.startswith('* ') or line.startswith('- ')
                    is_numbered = line[0].isdigit() and (len(line) > 1 and (line[1] == '.' or line[1] == ')')) if len(line) >= 2 else False

                    if is_bullet or is_numbered:
                        p = doc.add_paragraph(style='List Bullet' if is_bullet else 'List Number')
                        if is_numbered:
                            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT

                        if is_bullet:
                            line = line[2:].strip()
                    else:
                        p = doc.add_paragraph()
                        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

                    p.paragraph_format.space_after = Pt(8)

                    # --- (C) වැදගත්ම කොටස: ඕනෑම පේළියක ඇති Bold වචන සැකසීම ---
                    if '**' in line or '*' in line:
                        marker = '**' if '**' in line else '*'
                        parts = line.split(marker)
                        for i, part in enumerate(parts):
                            is_bold_part = (i % 2 != 0)
                            run = p.add_run(part)
                            apply_style(run, 12, bold=is_bold_part)
                    else:
                        run = p.add_run(line)
                        apply_style(run, 12, bold=False)

            logger.info(">>> [LOG]: Line-by-line formatting and bold processing completed successfully.")
            logger.info(f"Completed {sec}. Sleeping for 5 seconds...")
            time.sleep(5)


        # -------------------------------------------------------
        # විශේෂ ප්‍රශ්න සහ විසඳුම් (Custom Solutions) කොටස
        # -------------------------------------------------------
        questions_content = read_file("prompts/questions.txt")

        if questions_content:
            doc.add_page_break()
            logger.info(">>> CUSTOM QUESTIONS DETECTED: Formatting with new rules...")

            # ප්‍රධාන මාතෘකාව
            q_header = doc.add_paragraph()
            q_header.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = q_header.add_run("විශේෂ උපදේශනය සහ විසඳුම් සේවාව (2026 සිට ඉදිරියට)")
            apply_style(run, 18, bold=True)

            questions = [q.strip() for q in questions_content.split('\n') if q.strip()]

            for i, question in enumerate(questions):
                start_q_time = time.time()
                logger.info(f"QUESTION {i+1} START: Processing -> {question}")

                # >>> යාවත්කාලීන කළ නව Special Question Prompt එක <<<
                q_prompt = (
                    f"මෙම විශේෂ ප්‍රශ්නයට සෘජු සහ සවිස්තරාත්මක පිළිතුරක් ලබා දෙන්න: '{question}'\n\n"
                    "අනිවාර්ය නීති:\n"
                    "1. අතිශය වැදගත්: පාරිභෝගිකයා අසා ඇති ගැටලුවට පමණක් සෘජුවම පිළිතුරු දෙන්න. ගැටලුවට අදාළ නැති අනෙකුත් ග්‍රහයන් (1 සිට 12 භාවයන්), පෞරුෂය හෝ අනවශ්‍ය දෑ කිසිසේත් විස්තර නොකරන්න.\n"
                    "2. වර්තමාන කාලය 2026 ලෙස සලකා, ඊට අදාළ දශා කාලයන් පමණක් දක්වමින් ප්‍රශ්නයට අදාළ පිළිතුර පවසන්න.\n"
                    "3. ඡේද සැකසුම: සෑම ප්‍රධාන උප-මාතෘකාවක්ම '###' සලකුණෙන් ආරම්භ කරන්න. (උදා: ### රැකියාවේ අර්බුද සඳහා පිළියම්)\n"
                    "4. පිළියම් ඉදිරිපත් කිරීම: පිළියම් හෝ වත්පිළිවෙත් ඉදිරිපත් කිරීමේදී ඒවා එකම ඡේදයක ලියනවා වෙනුවට, එක් එක් පිළියම අලුත් පේළියකින් ආරම්භ කර ඉදිරියෙන් '-' සලකුණ යොදන්න. බෞද්ධානුකූල පිළිවෙත් පමණක් ලබා දෙන්න.\n"
                    "5. කතා කරන මිත්‍රශීලී භාෂාවෙන් පිළිතුරු ලියන්න. කිසිදු විටෙක වාක්‍ය අගට 'නේද?' යන්න නොයොදන්න.\n"
                    "6. සෑම ඡේදයක් අවසානයේදීම අනිවාර්යයෙන්ම හිස් පේළි දෙකක් (Double Enter) තබන්න.\n"
                    "7. වචන 300-600 අතර පිළිතුරක් සිංහලෙන් පමණක් ලබා දෙන්න."
                )

                q_response = send_with_retry(chat, q_prompt)
                raw_ans = q_response.text.strip()

                # --- ප්‍රශ්නය Word ගොනුවට ලිවීම ---
                p_q = doc.add_paragraph()
                p_q.space_before = Pt(20)
                run_q = p_q.add_run(f"ගැටලුව {i+1}: {question}")
                apply_style(run_q, 13, bold=True)

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
                        sub_points = a_line.split('\n')
                        for point in sub_points:
                            if not point.strip(): continue
                            p_p = doc.add_paragraph()
                            p_p.style = 'List Bullet'
                            clean_point = point.lstrip('- ').replace('*', '').strip()
                            run_p = p_p.add_run(clean_point)
                            apply_style(run_p, 12)
                            p_p.paragraph_format.space_after = Pt(6)

                    # --- (C) සාමාන්‍ය විස්තරය ---
                    else:
                        p_a = doc.add_paragraph()
                        p_a.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                        current_text = a_line.replace('**', '*')
                        if '*' in current_text:
                            parts = current_text.split('*')
                            for idx, part in enumerate(parts):
                                run = p_a.add_run(part)
                                apply_style(run, 12, bold=(idx % 2 != 0))
                        else:
                            run_a = p_a.add_run(current_text)
                            apply_style(run_a, 12)

                logger.info(f"QUESTION {i+1} COMPLETED: Formatted perfectly.")
                doc.add_paragraph("-" * 55).alignment = WD_ALIGN_PARAGRAPH.CENTER
                time.sleep(2)

        # -------------------------------------------------------
        # විශේෂ ශාස්ත්‍රීය සටහන (Final Special Page)
        # -------------------------------------------------------
        special_note_content = read_file("prompts/special_note.txt")

        if special_note_content:
            doc.add_page_break()
            logger.info("විශේෂ ශාස්ත්‍රීය සටහන පිටුව එක් කරමින්...")

            lines = special_note_content.split('\n')
            for line in lines:
                line = line.strip()
                if not line:
                    doc.add_paragraph() # හිස් පේළි සඳහා
                    continue

                p = doc.add_paragraph()

                if "විශේෂ ශාස්ත්‍රීය සටහන" in line:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run = p.add_run(line)
                    apply_style(run, 18, bold=True)
                    p.space_after = Pt(12)

                elif line in ["දේශය, කාලය සහ පෞරුෂය", "හවුල් කර්මය සහ සහකරුගේ බලපෑම",
                              "ග්‍රහයන්ගේ අංශකමය බලය සහ දශා කාල", "පුරුෂාර්ථය සහ වීර්යය", "අවසාන නිගමනය"]:
                    run = p.add_run(line)
                    apply_style(run, 14, bold=True)
                    p.space_before = Pt(10)

                elif "තෙරුවන් සරණින්" in line:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p.space_before = Pt(20)
                    run = p.add_run(line)
                    apply_style(run, 13, bold=True)

                else:
                    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                    run = p.add_run(line)
                    apply_style(run, 12)

        # -------------------------------------------------------
        # අවසාන Signature කොටස
        # -------------------------------------------------------
        contact_p = doc.add_paragraph()
        contact_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        contact_p.space_before = Pt(30)

        run = contact_p.add_run("මෙයට,\nපුරාණ ජෝතිර්වේදය හදහන් සේවය")
        apply_style(run, 12, bold=True)

        # අවසාන සුරැකීම
        output_file = "outputs/0711292927.docx"
        doc.save(output_file)
        logger.info(f"SUCCESS: Report saved as {output_file}")


    except Exception as e:
        logger.error(f"Error: {str(e)}", exc_info=True)

if __name__ == "__main__":
    generate_report()
