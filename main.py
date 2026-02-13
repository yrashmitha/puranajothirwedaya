import time
import logging
import sys
from google import genai
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, ns
from docx.enum.text import WD_ALIGN_PARAGRAPH
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


SYSTEM_PROMPT = read_file("prompts/system_prompt.txt")
BIRTH_DATA = read_file("prompts/birth_data.json")

sections = [
    "පෞරුෂය",
    "අධ්‍යාපනය",
    "වෘත්තීය ජීවිතය සහ ආර්ථික ශක්තිය",
    "ප්‍රේමය සහ විවාහ ජීවිතය",
    "දේපළ, භූමිය, නිවාස සහ වාහන භාග්‍යය",
    "ශාරීරික සෞඛ්‍යය, මාරක අපල, හදිසි අනතුරු",
    "දරු පල",
    "මෙතෙක් දැක්වූ කරුණු අනුව ජීවන ගමනේ සමස්ත සාරාංශය",
    ]

# --- 3. HELPER FUNCTIONS ---
import time
import logging
import sys
from google import genai
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
            config={'system_instruction': SYSTEM_PROMPT}
        )

        chat.send_message(BIRTH_DATA)
        logger.info("Context successfully registered.")

        # ප්‍රධාන මාතෘකාව
        title_p = doc.add_paragraph()
        title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = title_p.add_run("නමෝ බුද්ධාය!")
        apply_style(run, 24, bold=True)

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

            task_prompt = (f"දැන් [{sec}] පිළිබඳව {limit_text} "
                           f"1. සම්පූර්ණ විස්තරය වචන 800-1000 අතර සීමා කරන්න."
                           f"2. කිසිදු ඉංග්‍රීසි වචනයක් භාවිතා නොකරන්න. "
                           f"3. ඡේද Justify කිරීමට හැකි වන ලෙස අඛණ්ඩව ලියන්න."
                           f"4. කිසිදු <h3>, <b> වැනි HTML ටැග් භාවිතා නොකරන්න."
                           f"5. වැදගත් කරුණු පෙන්වීමට අනිවාර්යයෙන්ම '**' (Double Asterisks) දෙපසටම යොදන්න. (උදා: **විශේෂ ධන යෝගයක්**)"
                           )

            response = chat.send_message(task_prompt)
            content = response.text

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
                        continue # මාතෘකාවක් නම් පහළ තීරණ වලට නොගොස් ඊළඟ පේළියට යයි

                    # --- (B) Bullet Point හෝ අංකිත ලැයිස්තු (1. 2. 3.) හඳුනා ගැනීම ---
                    is_bullet = line.startswith('* ') or line.startswith('- ')
                    is_numbered = line[0].isdigit() and (len(line) > 1 and (line[1] == '.' or line[1] == ')')) if len(line) >= 2 else False

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

            logger.info(f"Completed {sec}. Sleeping for 5 seconds...")
            time.sleep(5)



        # -------------------------------------------------------
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

                # 1. Gemini හට ලබා දෙන නව උපදෙස් මාලාව (Prompt)
                q_prompt = (
                    f"පාරිභෝගිකයාගේ මෙම විශේෂ ප්‍රශ්නයට ඔහුගේ කේන්ද්‍රය සහ වර්තමාන ග්‍රහ ගෝචරය අනුව සවිස්තරාත්මක පිළිතුරු දෙන්න: '{question}'\n\n"
                    "මෙම පිළිතුර සඳහා පමණක් අදාළ වන විශේෂ තාක්ෂණික නීති:\n"
                    "1. කාල වකවානු නිර්ණය කිරීමේදී 2026 වසරේ සිට ඉදිරියට ඇති තත්ත්වය පමණක් සලකා බලන්න.\n"
                    "2. ඡේද සැකසුම: සෑම ප්‍රධාන උප-මාතෘකාවක්ම '###' සලකුණෙන් ආරම්භ කරන්න. (උදා: ### ශනිගේ අපල සඳහා පිළියම්)\n"
                    "3. පිළියම් ඉදිරිපත් කිරීමේ විශේෂ ක්‍රමය: පිළියම් හෝ වත්පිළිවෙත් ඉදිරිපත් කිරීමේදී ඒවා එකම ඡේදයක ලියනවා වෙනුවට, එක් එක් පිළියම අලුත් පේළියකින් ආරම්භ කර ඉදිරියෙන් '-' සලකුණ යොදන්න. බෞද්ධානුකූල පිලිවෙත් ලබා දෙන්න. පැලදිය හැකි සුදුසු මනික් වර්ග යෝජනා කරන්න (නමුත් මනික් පැලදීමෙදී නිසි පරික්ශාවක් අත්‍යවශ්‍ය බව පවසන්න)\n"
                    "4. සෑම ඡේදයක් අවසානයේදීම අනිවාර්යයෙන්ම හිස් පේළි දෙකක් (Double Enter) තබන්න.\n"
                    "5. වචන 600-800 අතර සවිස්තරාත්මක පිළිතුරක් සිංහලෙන් පමණක් ලබා දෙන්න."
                )

                q_response = chat.send_message(q_prompt)
                raw_ans = q_response.text.strip()

                # --- ප්‍රශ්නය Word ගොනුවට ලිවීම ---
                p_q = doc.add_paragraph()
                p_q.space_before = Pt(20)
                run_q = p_q.add_run(f"ගැටලුව {i+1}: {question}")
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

                logger.info(f"QUESTION {i+1} COMPLETED: Formatted perfectly.")
                doc.add_paragraph("-" * 55).alignment = WD_ALIGN_PARAGRAPH.CENTER
                time.sleep(2)

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
                    doc.add_paragraph() # හිස් පේළි සඳහා
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
        output_file = "outputs/dilini.docx"
        doc.save(output_file)
        logger.info(f"SUCCESS: Report saved as {output_file}")


    except Exception as e:
        logger.error(f"Error: {str(e)}", exc_info=True)


if __name__ == "__main__":
    generate_report()
