import io
import os
import re
import unicodedata
from docx import Document
import pypdf
import streamlit as st

# ==========================================
# PAGE CONFIGURATION & CUSTOM STYLING
# ==========================================
st.set_page_config(
    page_title="Ultra PDF to Clean Hindi Text & Word Converter",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    
    """,
    unsafe_allow_html=True,
)


# =========================================================
# ADVANCED DEVANAGARI REPAIR & CLEANING ENGINE
# =========================================================
def clean_and_repair_hindi(text, custom_replacements=None):
    if not text:
        return ""

    # 1. Unicode Normalization (NFC) & Garbage Symbols Removal
    text = unicodedata.normalize("NFC", str(text))
    text = (
        text.replace("\ufffd", "")
        .replace("\u0000", "")
        .replace("\ufeff", "")
        .replace("\r", "")
    )

    # 2. Fix Visual-to-Logical Order for Devanagari 'ि' (Short I Vowel Sign)
    text = re.sub(
        r"(\u093f)([\u0915-\u0939](?:\u094d[\u0915-\u0939])*)", r"\g<2>\g<1>", text
    )

    # 3. Spaced out Hindi characters fix (e.g., "भ ा र त" -> "भारत")
    text = re.sub(r"([\u0900-\u097F])\s+([\u0902-\u094D])", r"\1\2", text)

    # 4. Option Label Formatting (e.g., "aजला" -> "a) जिला")
    text = re.sub(
        r"^([a-dA-D1-9])\s*([\u0900-\u097F])", r"\1) \2", text, flags=re.MULTILINE
    )

    # 5. Core Devanagari Correction Dictionary
    rep_dict = {
        "संवधान": "संविधान",
        "पुलस": "पुलिस",
        "सावर्जनक": "सार्वजनिक",
        "रा ": "राज्य ",
        " व ": " वित्त, ",
        "संप,": "संपत्ति,",
        "वषय": "विषय",
        " ह।": " हैं।",
        "सातवी": "सातवीं",
        "बारवी": "बारहवीं",
        "पचवी": "पांचवीं",
        "नौव": "नौवीं",
        "सातव": "सातवीं",
        "आठव": "आठवीं",
        "दसव": "दसवीं",
        "7व": "7वीं",
        "8व": "8वीं",
        "9व": "9वीं",
        "10व": "10वीं",
        "तीसर": "तीसरी",
        "दूसर": "दूसरी",
        "पचव": "पांचवीं",
        "चौथी": "चौथी",
        "संबंधित": "संबंधित",
        "संबंधत": "संबंधित",
        "संधत": "संबंधित",
        "परवतर्न": "परिवर्तन",
        "अयोता": "अयोग्यता",
        "ावधान": "प्रावधान",
        "दया": "दिया",
        "मौलक": "मौलिक",
        "ानीय": "स्थानीय",
        "वणर्त": "वर्णित",
        "समवत": "समवर्ती",
        "एकमा": "एकमात्र",
        "नलखत": "निम्नलिखित",
        "क शासत": "केंद्र शासित",
        "देश": "प्रदेश",
        "वभ": "विभिन्न",
        "संवैधानक": "संवैधानिक",
        "पदाधकारय": "पदाधिकारियों",
        "द्विस्तारा": "द्वारा",
        "ाप": "प्रारूप",
        "अनुसूचय": "अनुसूचियां",
        "अंेजी": "अंग्रेजी",
        "संधी": "सिंधी",
        "मराठ": "मराठी",
        "संत": "संस्कृत",
        "ककणी": "कोंकणी",
        "मणपुर": "मणिपुरी",
        "उड़या": "उड़िया",
        "नेपाली ाा:": "नेपाली",
        "कार": "प्रकार",
        "क े": "के",
        "म ": "में ",
    }

    # Add optional custom replacements provided by user
    if custom_replacements:
        rep_dict.update(custom_replacements)

    for bad, good in rep_dict.items():
        text = text.replace(bad, good)

    # 6. Clean Multiple Blank Spaces and Line-breaks
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


# =========================================================
# PDF PROCESSING ENGINE
# =========================================================
def extract_pdf_content(file_bytes, noise_keywords=None):
    reader = pypdf.PdfReader(io.BytesIO(file_bytes))
    raw_text_pages = []
    cleaned_text_pages = []

    if not noise_keywords:
        noise_keywords = [
            "To Download",
            "For More Informa",
            "Crazy Gk",
            "Contact :",
            "SPACE FOR ROUGH WORK",
            "202084-OK-ZMZS-SA-E",
            "PART - I",
            "Set A",
            "1000 Question Series",
            "Topic Wise MCQS",
            "RAILWAY TARGET BATCH",
        ]

    for idx, page in enumerate(reader.pages):
        raw_page = page.extract_text() or ""
        raw_text_pages.append(f"--- PAGE {idx+1} ---\n" + raw_page)

        cleaned_lines = []
        for line in raw_page.split("\n"):
            line_cleaned = clean_and_repair_hindi(line)
            if line_cleaned and not any(
                k.lower() in line_cleaned.lower() for k in noise_keywords
            ):
                cleaned_lines.append(line_cleaned)

        page_content = "\n".join(cleaned_lines)
        if page_content.strip():
            cleaned_text_pages.append(f"--- PAGE {idx+1} ---\n" + page_content)

    return "\n\n".join(raw_text_pages), "\n\n".join(cleaned_text_pages)


# =========================================================
# WORD (.DOCX) GENERATOR
# =========================================================
def create_word_docx(text, doc_title="Cleaned Test Series Questions"):
    doc = Document()
    doc.add_heading(doc_title, level=1)

    for paragraph in text.split("\n\n"):
        if paragraph.strip():
            doc.add_paragraph(paragraph.strip())

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output


# =========================================================
# STREAMLIT UI INTERFACE
# =========================================================
st.title("📄 Ultra PDF to Clean Hindi Text & Word (.docx) Converter")
st.caption(
    "PDF se kharab font aur symbols ko repair karke shuddh Devanagari/Hindi text aur Word document banayein."
)

# Sidebar Options
st.sidebar.title("⚙️ PDF Cleaning Controls")
doc_heading = st.sidebar.text_input(
    "Word Heading", "Cleaned Test Series Questions"
)
remove_headers = st.sidebar.checkbox("Automatic Header/Footer Watermark Hatayein", value=True)

# File Upload
uploaded_files = st.file_uploader(
    "📂 PDF Files Upload Karein", type=["pdf"], accept_multiple_files=True
)

if uploaded_files:
    for uploaded_file in uploaded_files:
        st.markdown(f"### 📑 Processing File: `{uploaded_file.name}`")

        with st.spinner("PDF se text nikala aur font repair kiya ja raha hai..."):
            file_bytes = uploaded_file.read()
            raw_text, cleaned_text = extract_pdf_content(file_bytes)

            if cleaned_text:
                # 2 Tabs for Side-by-Side Comparison
                tab_clean, tab_raw = st.tabs(
                    ["✨ Repaired & Cleaned Text", "🔍 Raw Extracted Text"]
                )

                with tab_clean:
                    st.text_area(
                        label="Clean Hindi Text (Direct Copy Karein):",
                        value=cleaned_text,
                        height=350,
                        key=f"clean_txt_{uploaded_file.name}",
                    )

                    col1, col2 = st.columns(2)

                    # Option 1: Download Word Document
                    docx_file = create_word_docx(cleaned_text, doc_heading)
                    file_base_name = os.path.splitext(uploaded_file.name)[0]

                    with col1:
                        st.download_button(
                            label="📝 Download Word (.docx)",
                            data=docx_file,
                            file_name=f"{file_base_name}_Cleaned.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            key=f"word_{uploaded_file.name}",
                        )

                    # Option 2: Download Clean Text File
                    with col2:
                        st.download_button(
                            label="📄 Download Clean Text (.txt)",
                            data=cleaned_text,
                            file_name=f"{file_base_name}_Cleaned.txt",
                            mime="text/plain",
                            key=f"txt_{uploaded_file.name}",
                        )

                with tab_raw:
                    st.info(
                        "Yeh bina repair kiya gaya original PDF text hai (compare karne ke liye):"
                    )
                    st.text_area(
                        label="Raw PDF Text:",
                        value=raw_text,
                        height=350,
                        key=f"raw_txt_{uploaded_file.name}",
                    )
            else:
                st.error("Is PDF se text extract nahi ho saka. (File scanned image ho sakti hai)")
