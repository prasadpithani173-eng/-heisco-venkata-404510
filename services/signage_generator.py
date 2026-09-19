import os
import io
import glob
import base64
from typing import List, Tuple, Dict, Any
from PIL import Image, ImageDraw, ImageFont

# Translation dictionary for common HSE signage phrases
SIGNAGE_TRANSLATIONS: Dict[str, Dict[str, str]] = {
    "SLIP, TRIP, FALL HAZARD": {
        "English": "SLIP ,TRIP, FALL\nHAZARD",
        "Arabic": "خطر الانزلاق والتعثر والسقوط",
        "Hindi": "फिसलने, ठोकर लगने और गिरने\nका खतरा",
        "Urdu": "پھسلنے ، ٹھوکر لگنے اور گرنے کا\nخطرہ"
    },
    "SLIP ,TRIP, FALL HAZARD": {
        "English": "SLIP ,TRIP, FALL\nHAZARD",
        "Arabic": "خطر الانزلاق والتعثر والسقوط",
        "Hindi": "फिसलने, ठोकर लगने और गिरने\nका खतरा",
        "Urdu": "پھسلنے ، ٹھوکر لگنے اور گرنے کا\nخطرہ"
    },
    "SLIP, TRIP, FALL": {
        "English": "SLIP ,TRIP, FALL\nHAZARD",
        "Arabic": "خطر الانزلاق والتعثر والسقوط",
        "Hindi": "फिसलने, ठोकर लगने और गिरने\nका खतरा",
        "Urdu": "پھسلنے ، ٹھوکر لگنے اور گرنے کا\nخطرہ"
    },
    "PPE FREE ZONE": {
        "English": "PPE FREE ZONE",
        "Arabic": "منطقة خالية من معدات الحماية الشخصية",
        "Hindi": "पीपीई मुक्त क्षेत्र",
        "Urdu": "پی پی ای فری زون"
    },
    "MANDATORY PPE ZONE": {
        "English": "MANDATORY PPE ZONE",
        "Arabic": "منطقة ارتداء معدات الحماية إلزامية",
        "Hindi": "अनिवार्य पीपीई क्षेत्र",
        "Urdu": "لازمی پی پی ای زون"
    },
    "HARD HAT REQUIRED": {
        "English": "HARD HAT REQUIRED",
        "Arabic": "يجب ارتداء خوذة السلامة",
        "Hindi": "सुरक्षा हेलमेट अनिवार्य",
        "Urdu": "حفاظتی ہیلمٹ لازمی ہے"
    },
    "SAFETY GLASSES REQUIRED": {
        "English": "SAFETY GLASSES REQUIRED",
        "Arabic": "يجب ارتداء نظارات السلامة",
        "Hindi": "सुरक्षा चश्मा अनिवार्य",
        "Urdu": "حفاظتی چشمہ لازمی ہے"
    },
    "SAFETY SHOES REQUIRED": {
        "English": "SAFETY SHOES REQUIRED",
        "Arabic": "يجب ارتداء أحذية السلامة",
        "Hindi": "सुरक्षा जूते अनिवार्य",
        "Urdu": "حفاظتی جوتے لازمی ہیں"
    },
    "NO SMOKING": {
        "English": "NO SMOKING",
        "Arabic": "ممنوع التدخين",
        "Hindi": "धूम्रपान निषेध",
        "Urdu": "تمباکو نوشی منع ہے"
    },
    "RESTRICTED AREA": {
        "English": "RESTRICTED AREA",
        "Arabic": "منطقة محظورة - للمصرح لهم فقط",
        "Hindi": "प्रतिबंधित क्षेत्र",
        "Urdu": "ممنوعہ علاقہ"
    }
}

def get_font_paths():
    """Find the best available TrueType fonts for each script."""
    bold_candidates = [
        "arialbd.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
    ]
    bold_font = "arialbd.ttf"
    for cand in bold_candidates:
        if os.path.exists(cand):
            bold_font = cand
            break

    regular_candidates = [
        "arial.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
    ]
    regular_font = "arial.ttf"
    for cand in regular_candidates:
        if os.path.exists(cand):
            regular_font = cand
            break

    arabic_font = regular_font
    arabic_matches = (
        glob.glob('/usr/share/fonts/**/Noto*Arabic*.ttf', recursive=True) or
        glob.glob('/usr/share/fonts/**/Kacst*.ttf', recursive=True)
    )
    if arabic_matches:
        arabic_font = arabic_matches[0]

    hindi_font = regular_font
    hindi_matches = (
        glob.glob('/usr/share/fonts/**/Lohit-Devanagari.ttf', recursive=True) or
        glob.glob('/usr/share/fonts/**/*deva*.ttf', recursive=True) or
        glob.glob('/usr/share/fonts/**/Gargi.ttf', recursive=True)
    )
    if hindi_matches:
        hindi_font = hindi_matches[0]

    urdu_font = arabic_font
    urdu_matches = glob.glob('/usr/share/fonts/**/Noto*Urdu*.ttf', recursive=True)
    if urdu_matches:
        urdu_font = urdu_matches[0]

    return {
        "bold": bold_font,
        "regular": regular_font,
        "arabic": arabic_font,
        "hindi": hindi_font,
        "urdu": urdu_font
    }

def ensure_base_template(template_path: str = "templates/signage_template.png", width: int = 1400, height: int = 1000) -> Image.Image:
    """Load existing base template or create an industrial project template with logos and borders."""
    if os.path.exists(template_path):
        try:
            return Image.open(template_path).convert("RGBA")
        except Exception:
            pass

    base = Image.new("RGBA", (width, height), (255, 255, 255, 255))
    draw = ImageDraw.Draw(base)

    # Double border
    draw.rectangle([(10, 10), (width - 10, height - 10)], outline="black", width=6)
    draw.rectangle([(18, 18), (width - 18, height - 18)], outline="black", width=2)

    # Top common header (header_final.png)
    header_img_path = "public/assets/header_final.png"
    if not os.path.exists(header_img_path):
        header_img_path = "static/assets/header_final.png"

    if os.path.exists(header_img_path):
        try:
            h_img = Image.open(header_img_path).convert("RGBA")
            h_target_w = width - 40
            h_target_h = int(h_img.height * h_target_w / h_img.width)
            if h_target_h > 95:
                h_target_h = 95
                h_target_w = int(h_img.width * h_target_h / h_img.height)
            h_img_res = h_img.resize((h_target_w, h_target_h), Image.Resampling.LANCZOS)
            h_x = (width - h_target_w) // 2
            base.paste(h_img_res, (h_x, 22), h_img_res)
        except Exception:
            pass
    else:
        # Fallback if header_final.png is missing
        if os.path.exists("public/assets/enppi_logo.png"):
            try:
                enppi = Image.open("public/assets/enppi_logo.png").convert("RGBA")
                enppi_w = int(enppi.width * 65 / enppi.height)
                enppi = enppi.resize((enppi_w, 65))
                base.paste(enppi, (35, 35), enppi)
            except Exception:
                pass

        if os.path.exists("public/assets/aramco_overseas_logo.png"):
            try:
                aramco = Image.open("public/assets/aramco_overseas_logo.png").convert("RGBA")
                aramco_w = int(aramco.width * 65 / aramco.height)
                aramco = aramco.resize((aramco_w, 65))
                base.paste(aramco, (width - aramco_w - 35, 35), aramco)
            except Exception:
                pass

        try:
            f_proj = ImageFont.truetype("arialbd.ttf", 24)
            f_bi = ImageFont.truetype("arialbd.ttf", 20)
            t1 = "DEBOTTLENECK PRODUCTION FACILITIES AT ABQAIQ"
            t2 = "BI NO. 10-10303"
            b1 = draw.textbbox((0, 0), t1, font=f_proj)
            draw.text(((width - (b1[2] - b1[0])) / 2, 42), t1, font=f_proj, fill="black")
            b2 = draw.textbbox((0, 0), t2, font=f_bi)
            draw.text(((width - (b2[2] - b2[0])) / 2, 75), t2, font=f_bi, fill="black")
        except Exception:
            pass

    os.makedirs(os.path.dirname(template_path) or ".", exist_ok=True)
    base.save(template_path, format="PNG")
    return base

def get_translations_for(signage_name: str) -> Dict[str, str]:
    """Retrieve or generate translations for the given signage phrase."""
    key = signage_name.strip().upper()
    if key in SIGNAGE_TRANSLATIONS:
        return SIGNAGE_TRANSLATIONS[key]
    
    # Partial matching
    for k, trans in SIGNAGE_TRANSLATIONS.items():
        if k in key or key in k:
            return trans

    # Default fallback
    return {
        "English": signage_name.upper(),
        "Arabic": "منطقة خالية من معدات الحماية الشخصية",
        "Hindi": "पीपीई मुक्त क्षेत्र",
        "Urdu": "پی پی ای فری زون"
    }

def generate_composite_signage(
    signage_name: str = "SLIP ,TRIP, FALL HAZARD",
    pictogram_path: str = "icons/slip_trip_fall.png",
    header_text: str = "CAUTION",
    header_color: str = "#FFEE00",
    template_path: str = "templates/signage_template.png"
) -> Dict[str, Any]:
    """
    Generate the complete multilingual industrial safety signage matching the uploaded Aramco / Enppi format:
    - Top header: Enppi, Project BI 10-10303, Saudi Aramco
    - Hazard Banner: Yellow CAUTION banner with thick black outline
    - Left column: Pictogram inside a framed black box (400x400)
    - Right column: Main English title + Urdu, Hindi, and Arabic translations
    """
    fonts = get_font_paths()
    base = ensure_base_template(template_path)
    draw = ImageDraw.Draw(base)

    # 1. Header Banner
    banner_y1 = 125
    banner_y2 = 270
    draw.rectangle([(25, banner_y1), (base.width - 25, banner_y2)], fill=header_color, outline="black", width=5)

    try:
        font_caution = ImageFont.truetype(fonts["bold"], 125)
    except Exception:
        font_caution = ImageFont.load_default()

    h_text = header_text.upper()
    try:
        cb = draw.textbbox((0, 0), h_text, font=font_caution)
        c_w = cb[2] - cb[0]
        c_x = (base.width - c_w) / 2
    except Exception:
        c_x = base.width / 2 - 250
    draw.text((c_x, banner_y1 + 8), h_text, font=font_caution, fill="black")

    # 2. Left Column: Pictogram in Frame
    pic_box_x = 55
    pic_box_y = 300
    pic_box_w = 460
    pic_box_h = 460

    # Draw frame around pictogram area
    draw.rectangle(
        [(pic_box_x, pic_box_y), (pic_box_x + pic_box_w, pic_box_y + pic_box_h)],
        outline="black",
        width=5
    )

    # Paste pictogram
    resolved_pic_path = pictogram_path
    if not os.path.exists(resolved_pic_path):
        candidates = ["icons/slip_trip_fall.png", "icons/ppe_icon.png"]
        for cand in candidates:
            if os.path.exists(cand):
                resolved_pic_path = cand
                break

    if os.path.exists(resolved_pic_path):
        try:
            pictogram = Image.open(resolved_pic_path).convert("RGBA")
            pictogram = pictogram.resize((pic_box_w - 20, pic_box_h - 20))
            base.paste(pictogram, (pic_box_x + 10, pic_box_y + 10), pictogram)
        except Exception:
            pass

    # 3. Right Column: Main English Title & Multilingual Translations
    trans = get_translations_for(signage_name)
    right_x = 560

    # Main English title
    raw_main = trans.get("English", signage_name.upper())
    # Support multiple lines if \n is present
    main_lines = raw_main.split("\n")
    font_main_size = 90 if len(main_lines) > 1 else (95 if len(raw_main) <= 16 else 75)
    try:
        font_main = ImageFont.truetype(fonts["bold"], font_main_size)
    except Exception:
        font_main = font_caution

    curr_y = 295
    for line in main_lines:
        draw.text((right_x, curr_y), line.strip(), font=font_main, fill="black")
        curr_y += int(font_main_size * 1.15)

    curr_y += 20  # gap before translations

    # Font sizes for translations
    try:
        font_urdu = ImageFont.truetype(fonts["urdu"], 48)
        font_hindi = ImageFont.truetype(fonts["hindi"], 48)
        font_arabic = ImageFont.truetype(fonts["arabic"], 48)
    except Exception:
        font_urdu = font_hindi = font_arabic = ImageFont.load_default()

    # Urdu
    urdu_text = trans.get("Urdu", "پی پی ای فری زون")
    for u_line in urdu_text.split("\n"):
        try:
            draw.text((right_x, curr_y), u_line.strip(), font=font_urdu, fill="black", direction="rtl")
        except Exception:
            draw.text((right_x, curr_y), u_line.strip(), font=font_urdu, fill="black")
        curr_y += 68

    curr_y += 10

    # Hindi
    hindi_text = trans.get("Hindi", "पीपीई मुक्त क्षेत्र")
    for h_line in hindi_text.split("\n"):
        draw.text((right_x, curr_y), h_line.strip(), font=font_hindi, fill="black")
        curr_y += 68

    curr_y += 10

    # Arabic
    arabic_text = trans.get("Arabic", "منطقة خالية من معدات الحماية الشخصية")
    # If Arabic matches Urdu or is different, render cleanly
    if arabic_text and arabic_text != urdu_text:
        for a_line in arabic_text.split("\n"):
            try:
                draw.text((right_x, curr_y), a_line.strip(), font=font_arabic, fill="black", direction="rtl")
            except Exception:
                draw.text((right_x, curr_y), a_line.strip(), font=font_arabic, fill="black")
            curr_y += 68

    # Save output to BytesIO
    output = io.BytesIO()
    base.save(output, format="PNG")
    output.seek(0)

    b64_str = base64.b64encode(output.getvalue()).decode("ascii")
    data_uri = f"data:image/png;base64,{b64_str}"

    return {
        "stream": output,
        "data_uri": data_uri,
        "signage_name": signage_name,
        "header_text": header_text,
        "header_color": header_color,
        "pictogram_path": resolved_pic_path,
        "translations": trans
    }
