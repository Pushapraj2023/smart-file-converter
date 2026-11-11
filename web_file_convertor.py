import os
import re
import io
import json
from pathlib import Path
from PIL import Image

# Word imports
import docx
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import parse_xml, OxmlElement

# Excel imports
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Openpyxl Rich Text imports & ECMA-376 OpenXML Patch
try:
    from openpyxl.cell.rich_text import CellRichText, TextBlock
    from openpyxl.cell.text import InlineFont
    from openpyxl.xml.functions import Element, SubElement
    HAS_RICHTEXT = True

    def _bulletproof_textblock_to_tree(self):
        el = Element("r")
        rPr = SubElement(el, "rPr")
        if hasattr(self, 'font') and self.font:
            if getattr(self.font, 'rFont', None):
                f = SubElement(rPr, "rFont")
                f.set("val", str(self.font.rFont))
            if getattr(self.font, 'b', None):
                SubElement(rPr, "b")
            if getattr(self.font, 'sz', None):
                sz = SubElement(rPr, "sz")
                sz.set("val", str(int(round(float(self.font.sz)))))
        t = SubElement(el, "t")
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        t.text = self.text if self.text is not None else ""
        return el

    TextBlock.to_tree = _bulletproof_textblock_to_tree
except Exception:
    HAS_RICHTEXT = False

# Google GenAI SDK
try:
    from google import genai
    from google.genai import types
    import logging
    logging.getLogger("google_genai.models").setLevel(logging.ERROR)
    logging.getLogger("google.genai").setLevel(logging.ERROR)
    HAS_GENAI = True
except ImportError:
    HAS_GENAI = False

# PDF to images
try:
    import pypdfium2 as pdfium
    HAS_PDFIUM = True
except ImportError:
    HAS_PDFIUM = False

try:
    from pdf2image import convert_from_path, pdfinfo_from_path
    HAS_PDF2IMAGE = True
except ImportError:
    HAS_PDF2IMAGE = False


# ==============================================================================
# ACCURATE RULE-BASED UNICODE TO DEVLYS 010 ENGINE
# ==============================================================================

def unicode_to_devlys010(text: str) -> str:
    """
    High-Precision Unicode Devanagari to DevLys 010 (Remington) Converter.
    Accurately handles Madhya Pradesh, Viruddh, Kasht, Roop, and Visarg.
    """
    if not text:
        return ""

    out = text

    # Standardize punctuation & dashes
    out = out.replace('–', '-').replace('—', '-')

    # 1. High Priority Exact Departmental Matches & Special Conjuncts
    exact_fixes = [
        ("मध्यप्रदेश", "e/;izns'k"),
        ("मध्य प्रदेश", "e/; izns'k"),
        ("मध्य", "e/;"),
        ("विरुद्ध", "fo:)"),
        ("सूचीबद्ध", "lwphc))"),
        ("कष्ट", "d\"V"),
        ("स्पष्ट", "Li\"V"),
        ("दृष्टि", "n`f\"V"),
        ("रूप", ":i"),
        ("स्क्रीनशॉट", "LØhu'kkWV"),
        # Naye identified words
        ("प्लाट", "IykWV"),
        ("प्लॉट", "IykWV"),
        ("एक", ",d"),
        ("स्रोत", "lzksr"),
        ("एवं", ",oa"),
        ("क्षेत्र", "{ks="),
        ("मॉनिटरिंग", "ekWfuVfjax"),
        ("मॉनिटर", "ekWfuVj"),
        ("ॉ", "kW"),
        ("ॅ", "W"),
    ]
    for k, v in exact_fixes:
        out = out.replace(k, v)

    # 2. Chhoti 'i' (\u093F) handling
    out = re.sub(
        r'((?:[\u0915-\u0939]\u094d)*[\u0915-\u0939])\u093f',
        r'f\1',
        out
    )

    # 3. Reph ('र्' half-r above following letter)
    out = re.sub(
        r'\u0930\u094d((?:[\u0915-\u0939]\u094d)*[\u0915-\u0939](?:[\u093e-\u094c\u0902\u0901]?))',
        r'\1Z',
        out
    )

    # 4. Standard Multi-Character Conjuncts
    conjunct_rules = [
        ("क्ष", "{k"), ("त्र", "«"), ("ज्ञ", "K"), ("श्र", "J"),
        ("द्ध", ":)"), ("द्घ", "n~?"), ("द्ब", "n~c"), ("द्म", "n~e"),
        ("द्य", "n~;"), ("द्व", "n~o"), ("ट्ट", "V~V"), ("ट्ठ", "V~B"),
        ("ड्ड", "M~M"), ("ड्ढ", "M~<"), ("क्त", "Dr"), ("क्र", "dz"),
        ("प्र", "iz"), ("ग्र", "xz"), ("ब्र", "cz"), ("भ्र", "Hkz"),
        ("म्र", "ez"), ("व्रे", "oz"), ("ह्र", "gz"), ("ष्ट", "\"V"),
        ("ष्ठ", "\"B"), ("स्थ", "LFk"), ("श्च", "'p"), ("श्न", "'u"),
        ("ट्र", "Vª"), ("ड्र", "Mª"), ("ढ्र", "<ª"), ("ण्ड", ".M"),
        ("ण्ठ", ".B"), ("रु", "#"), ("रू", ":"), ("त्त", "Ùk"),
        ("न्न", "Uu"), ("क्क", "Dd"), ("ल्ल", "Yy"), # 'त्र' ko '=' se map karein (DevLys 010 standard):
        ("त्र", "="),
        # 'ॉ' (matra) ko 'kW' se map karein:
        ("ॉ", "kW"),
    ]
    for uni, dev in conjunct_rules:
        out = out.replace(uni, dev)

    # 5. Half Consonants (Halant forms)
    half_consonants = [
        ("क्", "D"), ("ख्", "[k"), ("ग्", "X"), ("घ्", "?k"),
        ("च्", "P"), ("छ्", "N"), ("ज्", "T"), ("झ्", ">k"),
        ("ञ्", "¥"), ("त्", "R"), ("थ्", "Fk"), ("द्", "n"),
        ("ध्", "/"),  ("न्", "U"), ("प्", "I"), ("फ्", "Q"),
        ("ब्", "C"), ("भ्", "Hk"), ("म्", "E"), ("य्", ";"),
        ("ल्", "Y"), ("व्", "O"), ("श्", "'k"), ("ष्", "\"k"),
        ("स्", "L"), ("ह्", "g"),
    ]
    for uni, dev in half_consonants:
        out = out.replace(uni, dev)

    # 6. Independent Vowels (Swar)
    vowels = [
        ("अ", "v"), ("आ", "vk"), ("इ", "b"), ("ई", "bZ"),
        ("उ", "m"), ("ऊ", "Å"), ("ऋ", "_"), ("ए", ",s"),
        ("ऐ", ",S"), ("ओ", "vks"), ("औ", "vkS"),
    ]
    for uni, dev in vowels:
        out = out.replace(uni, dev)

    # 7. Full Consonants (Vyanjan)
    full_consonants = [
        ("क", "d"), ("ख", "[k"), ("ग", "x"), ("घ", "?k"), ("ङ", "³"),
        ("च", "p"), ("छ", "N"), ("ज", "t"), ("झ", ">k"), ("ञ", "¥"),
        ("ट", "V"), ("ठ", "B"), ("ड", "M"), ("ढ", "<"), ("ण", ".k"),
        ("त", "r"), ("थ", "Fk"), ("द", "n"), ("ध", "/k"), ("न", "u"),
        ("प", "i"), ("फ", "Q"), ("ब", "c"), ("भ", "Hk"), ("म", "e"),
        ("य", ";"), ("र", "j"), ("ल", "y"), ("व", "o"), ("श", "'k"),
        ("ष", "\"k"), ("स", "l"), ("ह", "g"), ("ड़", "M+"), ("ढ़", "<+"),
        ("फ़", "t+"), ("ज़", "t+"),
    ]
    for uni, dev in full_consonants:
        out = out.replace(uni, dev)

    # 8. Matras & Diacritics
    matras = [
        ("ा", "k"), ("ी", "h"), ("ु", "q"), ("ू", "w"),
        ("ृ", "`"), ("े", "s"), ("ै", "S"), ("ो", "ks"),
        ("ौ", "kS"), ("ं", "a"), ("ँ", "¡"), ("ः", "%"),
        ("्", "~"), ("़", "+"), ("।", "A"), ("॥", "AA"),
    ]
    for uni, dev in matras:
        out = out.replace(uni, dev)

    return out


# ==============================================================================
# TOKENIZER (DevLys 15pt vs Arial 13pt + Space Preservation)
# ==============================================================================

def tokenize_content(text: str):
    """
    STRICT UNIVERSAL TOKENIZER:
    1. Hindi Devanagari text -> DevLys 010 (15 pt)
    2. English words and Digits -> Arial (13 pt)
    3. Symbols, Math signs (+, -, *, /, =, %, ., ,, :) -> Arial (13 pt)
    4. Spaces -> Preserved Space
    """
    if not text:
        return []
    pattern = re.compile(r'([\u0900-\u097F]+)|([A-Za-z0-9]+)|([^\s\w\u0900-\u097F]+)|(\s+)')
    tokens = []
    for match in pattern.finditer(text):
        hindi, eng_num, symbol, space = match.groups()
        if hindi:
            tokens.append(("HINDI", hindi))
        elif eng_num:
            tokens.append(("ARIAL", eng_num))
        elif symbol:
            tokens.append(("SYMBOL", symbol))
        elif space:
            tokens.append(("SPACE", space))
    return tokens


def _set_run_font_properties(run, font_name: str, size_pt: float):
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    rFonts.set(qn('w:ascii'), font_name)
    rFonts.set(qn('w:hAnsi'), font_name)
    rFonts.set(qn('w:cs'), font_name)


def _preserve_run_space(run):
    r = run._r
    for t in r.findall(qn('w:t')):
        t.set(qn('xml:space'), 'preserve')


# ==============================================================================
# CONVERTER CLASS
# ==============================================================================

class DocumentConverter:
    def __init__(self):
        appdata_dir = os.environ.get("APPDATA")
        if appdata_dir:
            self.suite_dir = Path(appdata_dir) / "SmartOfficeSuite"
        else:
            self.suite_dir = Path.home() / ".SmartOfficeSuite"
        self.suite_dir.mkdir(parents=True, exist_ok=True)
        self.config_path = self.suite_dir / "config.json"

        self._active_model = "gemini-3.5-flash"
        self._ensure_config()

    def _ensure_config(self):
        if not self.config_path.exists():
            default_conf = {"gemini_api_key": ""}
            try:
                self.config_path.write_text(json.dumps(default_conf, indent=4), encoding="utf-8")
            except Exception:
                pass

    def get_api_key(self) -> str:
        if self.config_path.exists():
            try:
                data = json.loads(self.config_path.read_text(encoding="utf-8"))
                key = data.get("gemini_api_key", "").strip()
                if key:
                    return key
            except Exception:
                pass
        return os.environ.get("GEMINI_API_KEY", "").strip()

    def set_api_key(self, key: str):
        key = key.strip()
        data = {}
        if self.config_path.exists():
            try:
                data = json.loads(self.config_path.read_text(encoding="utf-8"))
            except Exception:
                pass
        data["gemini_api_key"] = key
        self.config_path.write_text(json.dumps(data, indent=4), encoding="utf-8")

    def _get_gemini_client(self):
        api_key = self.get_api_key()
        if not api_key:
            raise ValueError(
                "Gemini API Key missing! Kripya 'PDF / Image Converter' tab me jakar "
                "apni Gemini API Key daalein aur 'Save Key' par click karein."
            )
        if not HAS_GENAI:
            raise ImportError("google-genai library install nahi hai. Run: pip install google-genai")
        return genai.Client(api_key=api_key)

    def parse_page_range(self, range_str: str, total_pages: int) -> list:
        """
        Parses page selection strings like "1", "1,3,5", "1-3", "all".
        Returns a sorted list of 1-based page numbers.
        """
        if not range_str or range_str.strip().lower() in ["all", "sab", ""]:
            return list(range(1, total_pages + 1))
        
        selected = set()
        parts = range_str.split(',')
        for part in parts:
            part = part.strip()
            if '-' in part:
                try:
                    start, end = map(int, part.split('-'))
                    selected.update(range(max(1, start), min(total_pages, end) + 1))
                except ValueError:
                    pass
            elif part.isdigit():
                val = int(part)
                if 1 <= val <= total_pages:
                    selected.add(val)
        return sorted(list(selected)) if selected else list(range(1, total_pages + 1))

    def extract_images_from_source(self, source_path: Path, page_selection: str = "all"):
        ext = source_path.suffix.lower()
        if ext in [".png", ".jpg", ".jpeg", ".bmp", ".tiff", ".webp"]:
            img = Image.open(source_path)
            return [img.convert("RGB")]

        if ext == ".pdf":
            if HAS_PDFIUM:
                pdf = pdfium.PdfDocument(str(source_path))
                total_pages = len(pdf)
                selected_pages = self.parse_page_range(page_selection, total_pages)
                
                images = []
                for p_num in selected_pages:
                    page = pdf[p_num - 1]
                    pix = page.render(scale=2.0)
                    images.append(pix.to_pil().convert("RGB"))
                return images

            if HAS_PDF2IMAGE:
                try:
                    info = pdfinfo_from_path(str(source_path))
                    total_pages = info["Pages"]
                    selected_pages = self.parse_page_range(page_selection, total_pages)
                    
                    images = []
                    for p_num in selected_pages:
                        imgs = convert_from_path(str(source_path), dpi=150, first_page=p_num, last_page=p_num)
                        if imgs:
                            images.append(imgs[0].convert("RGB"))
                    return images
                except Exception as e:
                    raise RuntimeError(f"PDF extract error: {e}")

            raise RuntimeError("PDF reading library (pypdfium2 ya pdf2image) uplabdh nahi hai.")

        raise ValueError(f"Unsupported file format: {ext}")

    def _query_gemini_vision(self, client, pil_image: Image.Image, prompt: str) -> str:
        w, h = pil_image.size
        max_dim = 1400
        if max(w, h) > max_dim:
            scale = max_dim / max(w, h)
            pil_image = pil_image.resize((int(w * scale), int(h * scale)), Image.Resampling.BILINEAR)

        buf = io.BytesIO()
        if pil_image.mode != "RGB":
            pil_image = pil_image.convert("RGB")
        pil_image.save(buf, format="JPEG", quality=85, optimize=True)
        img_bytes = buf.getvalue()

        image_part = types.Part.from_bytes(data=img_bytes, mime_type="image/jpeg")

        call_config = types.GenerateContentConfig(
            automatic_function_calling=types.AutomaticFunctionCallingConfig(disable=True),
            temperature=0.0
        )

        models_to_try = [self._active_model, "gemini-3.5-flash", "gemini-flash-latest"]
        last_err = None
        for model_name in models_to_try:
            try:
                response = client.models.generate_content(
                    model=model_name,
                    contents=[image_part, prompt],
                    config=call_config
                )
                if response and response.text:
                    self._active_model = model_name
                    return response.text.strip()
            except Exception as e:
                last_err = e
                continue

        raise RuntimeError(f"Gemini API call failed: {last_err}")

    def _render_paragraph_runs(self, paragraph, line_text: str):
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(2)
        paragraph.paragraph_format.line_spacing = 1.15

        tokens = tokenize_content(line_text)
        for token_type, val in tokens:
            if not val:
                continue
            run = paragraph.add_run()

            if token_type == "HINDI":
                run.text = unicode_to_devlys010(val)
                _set_run_font_properties(run, "DevLys 010", 15)
                _preserve_run_space(run)
            elif token_type in ("ARIAL", "SYMBOL"):
                run.text = val
                _set_run_font_properties(run, "Arial", 13)
                _preserve_run_space(run)
            elif token_type == "SPACE":
                run.text = val
                _set_run_font_properties(run, "Arial", 13)
                _preserve_run_space(run)

    def convert_to_word(self, source_path: str, dest_path: str, page_selection: str = "all", progress_cb=None):
        src = Path(source_path)
        dest = Path(dest_path)
        client = self._get_gemini_client()

        if progress_cb: progress_cb(0, 100, "Reading document pages...")
        images = self.extract_images_from_source(src, page_selection=page_selection)
        total_pages = len(images)

        doc = docx.Document()
        for section in doc.sections:
            section.top_margin = Inches(0.8)
            section.bottom_margin = Inches(0.8)
            section.left_margin = Inches(0.8)
            section.right_margin = Inches(0.8)

        prompt_word = (
            "Transcribe this official document page exactly as printed with 100% precision. "
            "CRITICAL: Keep all exact spaces between words. Do NOT merge, collapse, or remove spaces. "
            "Preserve all exact words, numbers, punctuation, and layout. "
            "Output titles and text lines normally. "
            "If there is a table, output it strictly as a Markdown table (| Col1 | Col2 |). "
            "Do NOT use HTML tags like <br>. Do NOT add bold asterisks (**). Output ONLY document text."
        )
        for i, img in enumerate(images, 1):
            if progress_cb:
                progress_cb(i, total_pages, f"Reading page {i} of {total_pages}...")

            page_text = self._query_gemini_vision(client, img, prompt_word)
            lines = page_text.splitlines()

            in_table = False
            table_buffer = []

            for line in lines:
                stripped = line.strip().replace("**", "").replace("*", "")

                if stripped.startswith("|") and (stripped.endswith("|") or "|" in stripped[1:]):
                    if re.match(r'^\|(?:\s*:?-+:?\s*\|)+$', stripped):
                        in_table = True
                        continue
                    raw_cells = [c.strip() for c in stripped.strip("|").split("|")]
                    clean_cells = [re.sub(r'<\s*br\s*/?>', ' ', c, flags=re.IGNORECASE).strip() for c in raw_cells]
                    table_buffer.append(clean_cells)
                    in_table = True
                    continue
                else:
                    if in_table and table_buffer:
                        self._add_native_word_table(doc, table_buffer)
                        table_buffer = []
                        in_table = False

                    if not stripped:
                        continue

                    p = doc.add_paragraph()
                    self._render_paragraph_runs(p, stripped)

            if in_table and table_buffer:
                self._add_native_word_table(doc, table_buffer)

            if i < total_pages:
                doc.add_page_break()

        dest.parent.mkdir(parents=True, exist_ok=True)
        doc.save(str(dest))
        if progress_cb: progress_cb(total_pages, total_pages, "✔ Word Document Created Successfully!")

    def _add_native_word_table(self, doc, table_data):
        if not table_data:
            return

        num_rows = len(table_data)
        num_cols = max(len(row) for row in table_data)

        table = doc.add_table(rows=num_rows, cols=num_cols)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.style = 'Table Grid'

        for row_idx, row in enumerate(table_data):
            trPr = table.rows[row_idx]._tr.get_or_add_trPr()
            trPr.append(parse_xml(r'<w:cantSplit xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'))

            is_header_row = (row_idx == 0 or (row_idx == 1 and all(c.strip().isdigit() for c in row if c.strip())))

            for col_idx in range(num_cols):
                cell_text = row[col_idx] if col_idx < len(row) else ""
                cell_text = re.sub(r'<\s*br\s*/?>', '\n', cell_text, flags=re.IGNORECASE).strip()

                cell = table.cell(row_idx, col_idx)
                p = cell.paragraphs[0]
                p.paragraph_format.space_before = Pt(1)
                p.paragraph_format.space_after = Pt(1)
                p.paragraph_format.line_spacing = 1.05

                if is_header_row:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

                parts = cell_text.split('\n')
                for p_idx, part in enumerate(parts):
                    if p_idx > 0:
                        p = cell.add_paragraph()
                        p.paragraph_format.space_before = Pt(1)
                        p.paragraph_format.space_after = Pt(1)
                        if is_header_row:
                            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    self._render_paragraph_runs(p, part.strip())

        doc.add_paragraph()

    def _render_excel_cell(self, cell, val_str: str, is_title: bool = False, is_header: bool = False, is_sub_header: bool = False):
        """
        ECMA-376 COMPLIANT EXCEL CELL RENDERER:
        - Hindi text -> DevLys 010 (15 pt)
        - English words, numbers, symbols (+, -, /, %, =, etc.) -> Arial (13 pt)
        """
        val_str = str(val_str).strip()
        val_str = re.sub(r'<\s*br\s*/?>', ' ', val_str, flags=re.IGNORECASE)
        val_str = val_str.replace("**", "").replace("*", "").strip()

        is_bold = bool(is_title or is_header or is_sub_header)
        tokens = tokenize_content(val_str)

        has_hindi = any(t[0] == "HINDI" for t in tokens)
        has_other = any(t[0] in ("ARIAL", "SYMBOL") for t in tokens)

        if has_hindi and has_other and HAS_RICHTEXT:
            rt = CellRichText()
            for token_type, token_val in tokens:
                if not token_val:
                    continue
                if token_type == "HINDI":
                    converted = unicode_to_devlys010(token_val)
                    f_devlys = InlineFont(rFont="DevLys 010", sz=15, b=is_bold)
                    rt.append(TextBlock(f_devlys, converted))
                elif token_type == "SPACE":
                    f_space = InlineFont(rFont="Arial", sz=13, b=is_bold)
                    rt.append(TextBlock(f_space, token_val))
                else:
                    f_arial = InlineFont(rFont="Arial", sz=13, b=is_bold)
                    rt.append(TextBlock(f_arial, token_val))
            cell.value = rt
        elif has_hindi:
            converted = unicode_to_devlys010(val_str)
            cell.value = converted
            cell.font = Font(name="DevLys 010", size=15, bold=is_bold)
        else:
            if re.match(r'^-?\d+(\.\d+)?$', val_str):
                try:
                    num_val = float(val_str) if '.' in val_str else int(val_str)
                    cell.value = num_val
                    cell.font = Font(name="Arial", size=13, bold=is_bold)
                    cell.number_format = '#,##0.00' if isinstance(num_val, float) else '#,##0'
                    return
                except ValueError:
                    pass
            cell.value = val_str
            cell.font = Font(name="Arial", size=13, bold=is_bold)

    def convert_to_excel(self, source_path: str, dest_path: str, page_selection: str = "all", progress_cb=None):
        src = Path(source_path)
        dest = Path(dest_path)
        client = self._get_gemini_client()

        if progress_cb: progress_cb(0, 100, "Reading document for Excel conversion...")
        images = self.extract_images_from_source(src, page_selection=page_selection)
        total_pages = len(images)

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Extracted Sheet"
        current_excel_row = 1

        prompt_excel = (
            "Extract ALL contents of this official document with 100% precision: "
            "1. TITLES: Include title and reference lines as the first rows. "
            "2. TABLES: Extract all headers, subheaders, and data rows. Keep dates, numbers, references, and amounts exact. "
            "Do NOT use '<br>' tags or bold asterisks. Output strictly as a JSON array of arrays of strings: "
            "[ [\"शीर्षक पंक्ति 1\"], [\"स्तंभ 1\", \"स्तंभ 2\", \"स्तंभ 3\"] ] "
            "Return ONLY raw JSON string."
        )

        for i, img in enumerate(images, 1):
            if progress_cb:
                progress_cb(i, total_pages, f"Extracting page {i} of {total_pages}...")

            raw_resp = self._query_gemini_vision(client, img, prompt_excel)
            clean_json = re.sub(r'^```json\s*', '', raw_resp.strip(), flags=re.IGNORECASE)
            clean_json = re.sub(r'\s*```$', '', clean_json.strip())

            table_data = []
            try:
                table_data = json.loads(clean_json)
            except Exception:
                lines = raw_resp.strip().splitlines()
                for line in lines:
                    cells = [c.strip() for c in re.split(r'\||\t', line) if c.strip()]
                    if cells:
                        table_data.append(cells)

            thin_border = Border(
                left=Side(style='thin', color='C0C0C0'),
                right=Side(style='thin', color='C0C0C0'),
                top=Side(style='thin', color='C0C0C0'),
                bottom=Side(style='thin', color='C0C0C0')
            )

            for row_values in table_data:
                if not isinstance(row_values, list) or not row_values:
                    continue

                is_title_row = (len(row_values) == 1 or (len(row_values) < 3 and current_excel_row <= 3))

                if is_title_row:
                    title_text = " ".join(str(v).strip() for v in row_values if str(v).strip())
                    cell = ws.cell(row=current_excel_row, column=1)
                    self._render_excel_cell(cell, title_text, is_title=True)
                    cell.alignment = Alignment(horizontal="left", vertical="center")
                    ws.row_dimensions[current_excel_row].height = 36
                    current_excel_row += 1
                    continue

                row_str_lower = "".join(str(v).lower() for v in row_values)
                is_main_header = any(kw in row_str_lower for kw in ["gstn", "gstin", "propritor", "proprietor", "period", "recovered", "standing", "amount", "name", "विवरण", "क्रमांक"])
                is_sub_header = all(str(v).strip().isdigit() for v in row_values if str(v).strip())

                for col_idx, val in enumerate(row_values, 1):
                    val_str = str(val).strip()
                    cell = ws.cell(row=current_excel_row, column=col_idx)
                    self._render_excel_cell(cell, val_str, is_header=is_main_header, is_sub_header=is_sub_header)

                    if is_main_header or is_sub_header:
                        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                    elif isinstance(cell.value, (int, float)):
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                    else:
                        align_h = "center" if (col_idx == 1 or len(val_str) <= 4 or val_str.isdigit()) else "left"
                        cell.alignment = Alignment(horizontal=align_h, vertical="center", wrap_text=True)

                    cell.border = thin_border

                if is_main_header:
                    ws.row_dimensions[current_excel_row].height = 36
                elif is_sub_header:
                    ws.row_dimensions[current_excel_row].height = 28
                else:
                    ws.row_dimensions[current_excel_row].height = 30

                current_excel_row += 1

        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                if cell.row <= 3 and len(str(cell.value or "")) > 30:
                    continue
                lines = str(cell.value or "").split('\n')
                for l in lines:
                    if len(l) > max_len:
                        max_len = len(l)
            ws.column_dimensions[col_letter].width = max(max_len + 4, 16)

        dest.parent.mkdir(parents=True, exist_ok=True)
        wb.save(str(dest))
        if progress_cb: progress_cb(total_pages, total_pages, "✔ Excel Spreadsheet Created Successfully!")
