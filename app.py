"""
Streamlit App — Remplissage automatique du canevas MINEDUB
Colle le texte structuré → génère le document Word rempli
"""

import streamlit as st
import os
import re
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

# ─────────────────────────────────────────────
st.set_page_config(
    page_title="MINEDUB — Remplissage Canevas",
    page_icon="📄",
    layout="centered",
)

TEMPLATE_PATH = Path(__file__).parent / "template.docx"
NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
w = f"{{{NS}}}"

# ─────────────────────────────────────────────
# MAPPING : field_index → (row_index_in_data_rows, label)
# ─────────────────────────────────────────────
FIELD_MAP = [
    (0,  "N° (numéro d'ordre dans le fichier)"),
    (1,  "MATRICULE"),
    (2,  "NOMS ET PRÉNOMS"),
    (3,  "DATE DE NAISSANCE (JJ/MM/AAAA)"),
    (4,  "LIEU DE NAISSANCE"),
    (5,  "SEXE"),
    (6,  "STATUT FAMILIAL"),
    (7,  "NOMBRE D'ENFANTS"),
    (8,  "RÉGION D'ORIGINE"),
    (9,  "DÉPARTEMENT D'ORIGINE"),
    (10, "ARRONDISSEMENT D'ORIGINE"),
    (11, "ETHNIE"),
    (12, "DIPLÔME ACADÉMIQUE"),
    (13, "DIPLÔME PROFESSIONNEL"),
    (14, "TÉLÉPHONE"),
    (15, "E-MAIL"),
    (16, "DATE D'ENTRÉE DANS LA FONCTION PUBLIQUE"),
    (17, "RÉFÉRENCE ACTE DE RECRUTEMENT"),
    (18, "STATUT"),
    (19, "CORPS"),
    (20, "CADRE"),
    (21, "GRADE"),
    (22, "CATÉGORIE"),
    (23, "CLASSE"),
    (24, "ÉCHELON"),
    (25, "INDICE"),
    (26, "DREB (Délégation Régionale)"),
    (27, "DDEB / SOUS-DIRECTION / SERVICE"),
    (28, "IAEB / SERVICE / BUREAU"),
    (29, "VILLE / VILLAGE D'AFFECTATION"),
    (30, "STRUCTURE D'AFFECTATION"),
    (31, "POSTE OCCUPÉ"),
    (32, "RANG DU POSTE"),
    (33, "DATE D'AFFECTATION / NOMINATION"),
    (34, "RÉFÉRENCE ACTE D'AFFECTATION / NOMINATION"),
    (35, "POSITION DE GESTION"),
    (36, "HANDICAP"),
    (37, "OBSERVATIONS"),
]

# ─────────────────────────────────────────────
# PARSING
# ─────────────────────────────────────────────
def parse_pasted_text(text: str) -> list:
    raw_lines = text.strip().splitlines()
    values = []
    skip_patterns = [
        r"^INFORMATIONS PERSONNELLES",
        r"^\s*$",
        r"^[-─═*\s]+$",
    ]
    for line in raw_lines:
        stripped = line.strip().lstrip("*").strip()
        if not stripped:
            continue
        skip = False
        for pat in skip_patterns:
            if re.match(pat, stripped, re.IGNORECASE):
                skip = True
                break
        if skip:
            continue
        values.append(stripped)
    return values


# ─────────────────────────────────────────────
# CONVERSION DOCX → PDF (Windows / Microsoft Word)
# ─────────────────────────────────────────────
def convert_docx_to_pdf(docx_path: str, pdf_path: str) -> bool:
    try:
        import win32com.client
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(os.path.abspath(docx_path))
        doc.SaveAs(os.path.abspath(pdf_path), FileFormat=17)  # wdFormatPDF
        doc.Close()
        word.Quit()
        return True
    except Exception:
        return False


# ─────────────────────────────────────────────
# REMPLISSAGE DU DOCUMENT
# ─────────────────────────────────────────────
def fill_document(values: list, output_path: str):
    # Register namespaces to preserve all prefixes
    namespaces_to_register = {
        "wpc": "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
        "cx": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        "cx1": "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex",
        "cx2": "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex",
        "cx3": "http://schemas.microsoft.com/office/drawing/2016/5/9/chartex",
        "cx4": "http://schemas.microsoft.com/office/drawing/2016/5/10/chartex",
        "cx5": "http://schemas.microsoft.com/office/drawing/2016/5/11/chartex",
        "cx6": "http://schemas.microsoft.com/office/drawing/2016/5/12/chartex",
        "cx7": "http://schemas.microsoft.com/office/drawing/2016/5/13/chartex",
        "cx8": "http://schemas.microsoft.com/office/drawing/2016/5/14/chartex",
        "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
        "aink": "http://schemas.microsoft.com/office/drawing/2016/ink",
        "am3d": "http://schemas.microsoft.com/office/drawing/2017/model3d",
        "o": "urn:schemas-microsoft-com:office:office",
        "oel": "http://schemas.microsoft.com/office/2019/extlst",
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        "v": "urn:schemas-microsoft-com:vml",
        "wp14": "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
        "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
        "w10": "urn:schemas-microsoft-com:office:word",
        "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
        "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
        "w16cex": "http://schemas.microsoft.com/office/word/2018/wordml/cex",
        "w16cid": "http://schemas.microsoft.com/office/word/2016/wordml/cid",
        "w16": "http://schemas.microsoft.com/office/word/2018/wordml",
        "w16du": "http://schemas.microsoft.com/office/word/2023/wordml/word16du",
        "w16sdtdh": "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash",
        "w16sdtfl": "http://schemas.microsoft.com/office/word/2024/wordml/sdtformatlock",
        "w16se": "http://schemas.microsoft.com/office/word/2015/wordml/symex",
        "wpg": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
        "wpi": "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
        "wne": "http://schemas.microsoft.com/office/word/2006/wordml",
        "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    }
    for prefix, uri in namespaces_to_register.items():
        ET.register_namespace(prefix, uri)

    with zipfile.ZipFile(str(TEMPLATE_PATH), 'r') as zin:
        doc_xml_bytes = zin.read("word/document.xml")

    root = ET.fromstring(doc_xml_bytes)
    rows = root.findall(f".//{w}tr")
    data_rows = rows[1:]  # skip header

    for field_idx, (row_idx, label) in enumerate(FIELD_MAP):
        if field_idx >= len(values):
            break
        value = values[field_idx]
        # Skip empty / ❌ values — leave cell as-is
        if value in ("❌", ""):
            continue
        if row_idx >= len(data_rows):
            continue

        row = data_rows[row_idx]
        cells = row.findall(f"{w}tc")
        if not cells:
            continue
        left_cell = cells[0]
        runs = left_cell.findall(f".//{w}r")
        if not runs:
            continue

        # ── Special case: Row 0 — "N° : 1" (value embedded in colon run)
        if row_idx == 0:
            # run[1] contains " : 1" → replace with " : <value>"
            if len(runs) >= 2:
                t = runs[1].find(f"{w}t")
                if t is not None:
                    t.text = f" : {value}"
            continue

        # ── Find the colon run
        colon_idx = None
        for j, r in enumerate(runs):
            t = r.find(f"{w}t")
            if t is not None and t.text and ':' in t.text:
                colon_idx = j
                break

        if colon_idx is not None:
            # Value runs are those after the colon run
            value_runs = runs[colon_idx + 1:]
            if value_runs:
                # Put the full value in the first value run
                first_t = value_runs[0].find(f"{w}t")
                if first_t is not None:
                    first_t.text = value
                    if ' ' in value:
                        first_t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                # Clear any additional value runs (multi-run values like GRÂCE split over 3)
                for extra_r in value_runs[1:]:
                    extra_t = extra_r.find(f"{w}t")
                    if extra_t is not None:
                        extra_t.text = ""
            else:
                # No run after colon — value is embedded in colon run text, append
                colon_run = runs[colon_idx]
                t = colon_run.find(f"{w}t")
                if t is not None and t.text:
                    # Strip old value after ":" and set new
                    base = t.text.split(":")[0]
                    t.text = base + ":" + value
        else:
            # No colon found — row has no existing value (CADRE, CLASSE, etc.)
            # Append a new run to the paragraph with the value
            paras = left_cell.findall(f"{w}p")
            if paras:
                para = paras[-1]
                # Copy rPr from last existing run if possible
                last_run = runs[-1]
                rpr = last_run.find(f"{w}rPr")

                new_run = ET.SubElement(para, f"{w}r")
                if rpr is not None:
                    import copy
                    new_run.insert(0, copy.deepcopy(rpr))
                new_t = ET.SubElement(new_run, f"{w}t")
                new_t.text = " : " + value
                if ' ' in value:
                    new_t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

    # Serialize
    new_xml_str = ET.tostring(root, encoding="unicode", xml_declaration=False)
    new_xml_bytes = ('<?xml version="1.0" encoding="UTF-8"?>' + new_xml_str).encode("utf-8")

    # Write output docx
    with zipfile.ZipFile(str(TEMPLATE_PATH), 'r') as zin:
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == "word/document.xml":
                    zout.writestr(item, new_xml_bytes)
                else:
                    zout.writestr(item, zin.read(item.filename))


# ─────────────────────────────────────────────
# INTERFACE STREAMLIT
# ─────────────────────────────────────────────
st.markdown("""
<style>
    .main-title {
        font-size: 1.6rem;
        font-weight: 700;
        color: #1a4a8a;
        margin-bottom: 0.2rem;
    }
    .subtitle { font-size: 0.95rem; color: #555; margin-bottom: 1.5rem; }
    .field-card {
        background: #f0f4ff;
        border-left: 4px solid #1a4a8a;
        padding: 0.35rem 0.7rem;
        border-radius: 4px;
        margin: 0.15rem 0;
        font-size: 0.84rem;
        line-height: 1.4;
    }
    .field-card b { color: #1a4a8a; display: block; }
    .field-val { color: #111; }
    .field-empty { color: #aaa; font-style: italic; }
    .success-box {
        background: #e8f5e9; border: 1px solid #81c784;
        border-radius: 8px; padding: 1rem; margin-top: 1rem;
    }
    .warn-box {
        background: #fff8e1; border: 1px solid #ffd54f;
        border-radius: 8px; padding: 1rem; margin-top: 0.5rem;
    }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">📄 Remplissage automatique — Canevas MINEDUB</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="subtitle">Collez le texte structuré ci-dessous pour générer instantanément '
    'le document Word rempli.</div>',
    unsafe_allow_html=True
)

# ── Exemple format ──
with st.expander("ℹ️ Format attendu (38 valeurs, une par ligne précédée de *)", expanded=False):
    st.code(
        "*INFORMATIONS PERSONNELLES PAR ORDRE DE FICHIER*:\n"
        "* N°1\n* 1088581T\n* NYANGWE NJANGUI CHRISTELLE GRÂCE\n"
        "* 24/09/1985\n* YAOUNDÉ\n* FÉMININ\n* CÉLIBATAIRE\n* 02\n"
        "* LITTORAL\n* NKAM\n* WOURI\n* BASSA'A\n"
        "* BACCALAURÉAT A4 ESPAGNOL\n* CAPIEMP\n* 655597311\n"
        "* njanguichristelle@gmail.com\n* 31/08/2018\n"
        "* N°0073/CR/MINEDUB/CAB\n* CONTRACTUEL\n* CORPS DE CONTRACTUELS\n"
        "* ❌\n* CONTRACTUEL\n* CATÉGORIE 8\n* ❌\n* ÉCHELON 4\n* ❌\n"
        "* DREB DU LITTORAL\n* DDEB DU WOURI\n* IAEB DE DOUALA 5 ÈME\n"
        "* DOUALA\n* E.P CAMP MBOPPI G1B\n* CHARGÉE DE CLASSE\n"
        "* ❌\n* ❌\n* ❌\n* EN ACTIVITÉ\n* ❌\n* ❌",
        language="text",
    )
    st.caption("Les ❌ = champ non renseigné (laissé vide dans le document).")

# ── Saisie ──
st.markdown("### 📋 Texte à coller")
pasted = st.text_area(
    label="zone",
    height=380,
    placeholder="*INFORMATIONS PERSONNELLES PAR ORDRE DE FICHIER*:\n* N°1\n* 1088581T\n...",
    label_visibility="collapsed",
)

col_btn, col_info = st.columns([1, 2])
with col_btn:
    generate = st.button("⚙️ Générer le document Word", type="primary", use_container_width=True)

if generate:
    if not pasted.strip():
        st.warning("⚠️ Veuillez coller le texte avant de cliquer sur Générer.")
    else:
        values = parse_pasted_text(pasted)

        # Aperçu des valeurs
        with st.expander(f"🔍 Aperçu — {len(values)} valeurs détectées", expanded=True):
            cols = st.columns(2)
            for i, (_, label) in enumerate(FIELD_MAP):
                val = values[i] if i < len(values) else None
                col = cols[i % 2]
                with col:
                    if val and val != "❌":
                        st.markdown(
                            f'<div class="field-card"><b>{i+1:02d}. {label}</b>'
                            f'<span class="field-val">{val}</span></div>',
                            unsafe_allow_html=True
                        )
                    elif val == "❌":
                        st.markdown(
                            f'<div class="field-card"><b>{i+1:02d}. {label}</b>'
                            f'<span class="field-empty">— (vide)</span></div>',
                            unsafe_allow_html=True
                        )
                    else:
                        st.markdown(
                            f'<div class="field-card"><b>{i+1:02d}. {label}</b>'
                            f'<span class="field-empty">⚠️ MANQUANT</span></div>',
                            unsafe_allow_html=True
                        )

        if len(values) < 38:
            st.markdown(
                f'<div class="warn-box">⚠️ <b>{len(values)} valeurs</b> détectées au lieu de '
                f'38 attendues. Vérifiez que toutes les lignes sont présentes (y compris les ❌).'
                f'</div>',
                unsafe_allow_html=True
            )

        if len(values) >= 5:
            with st.spinner("⏳ Génération en cours..."):
                try:
                    tmp_docx = tempfile.mktemp(suffix=".docx")
                    tmp_pdf = tempfile.mktemp(suffix=".pdf")
                    fill_document(values, tmp_docx)

                    with open(tmp_docx, "rb") as f:
                        docx_bytes = f.read()

                    pdf_ok = convert_docx_to_pdf(tmp_docx, tmp_pdf)
                    if pdf_ok:
                        with open(tmp_pdf, "rb") as f:
                            pdf_bytes = f.read()

                    os.unlink(tmp_docx)
                    if pdf_ok:
                        os.unlink(tmp_pdf)

                    nom = values[2] if len(values) > 2 else "agent"
                    matricule = values[1] if len(values) > 1 else "XXX"
                    safe_nom = re.sub(r'[^A-Za-z0-9]', '_', nom.upper())[:30]
                    docx_filename = f"Canevas_{matricule}_{safe_nom}.docx"
                    pdf_filename = f"Canevas_{matricule}_{safe_nom}.pdf"

                    st.markdown(
                        '<div class="success-box">✅ <b>Document généré avec succès !</b> '
                        'Téléchargez-le ci-dessous.</div>',
                        unsafe_allow_html=True
                    )
                    col1, col2 = st.columns(2)
                    with col1:
                        st.download_button(
                            label="⬇️ Word (.docx)",
                            data=docx_bytes,
                            file_name=docx_filename,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            type="primary",
                            use_container_width=True,
                        )
                    with col2:
                        if pdf_ok:
                            st.download_button(
                                label="📄 PDF (.pdf)",
                                data=pdf_bytes,
                                file_name=pdf_filename,
                                mime="application/pdf",
                                type="primary",
                                use_container_width=True,
                            )
                        else:
                            st.info("ℹ️ PDF indisponible (installer pywin32)", icon="⚠️")
                except Exception as e:
                    st.error(f"❌ Erreur : {e}")
                    st.exception(e)

# ── Légende des 38 champs ──
st.divider()
with st.expander("📑 Liste complète des 38 champs (ordre attendu)", expanded=False):
    for i, (_, label) in enumerate(FIELD_MAP, 1):
        st.markdown(f"**{i:02d}.** {label}")