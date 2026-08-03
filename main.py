import io
import re
import json
import uuid
from pathlib import Path
from datetime import date, datetime

import streamlit as st
import anthropic
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload, MediaIoBaseDownload

st.set_page_config(
    page_title="Research & Insights Automation Engine",
    page_icon="🔍",
    layout="wide"
)

st.title("🔍 Research & Insights Specialist Engine")
st.caption("Automated Research & Business Intelligence Generator — aligned to the Research & Insights Specialist Handbook")

# ----------------------------
# Configuration
# ----------------------------
try:
    api_key = st.secrets["ANTHROPIC_API_KEY"]
except Exception:
    st.error("⚠️ API Key not found in Streamlit Secrets. Please check your app settings.")
    st.stop()

MODEL_NAME = "claude-sonnet-5"

KB_CATEGORIES = ["Market", "Consumer", "Competitors", "Content", "Trends", "Technology", "Statistics"]

FOCUS_TO_CATEGORY = {
    "Consumer Behavior & Pain Points": "Consumer",
    "Competitor Positioning & Branding": "Competitors",
    "Packaging & Visual Identity Trends": "Trends",
    "Social Media & Platform Updates": "Content",
    "AI Tools & MarTech Innovations": "Technology",
    "Market Gaps & Business Opportunities": "Market",
}

DELIVERABLE_TYPES = [
    "General Research Summary",
    "Competitor Research Report",
    "Industry Update Report",
    "Marketing & AI Update",
    "Competitor Highlights",
    "Opportunity Report",
    "Trend Collection & Opportunity Audit",
    "Client Onboarding Research File",
    "Knowledge Base Review",
]

# ----------------------------
# Google Drive / Docs Integration
# ----------------------------
SCOPES = [
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/documents'
]

@st.cache_resource
def get_drive_services():
    """Authenticates using Service Account credentials from Streamlit Secrets."""
    if "GDRIVE_SERVICE_ACCOUNT" not in st.secrets:
        return None, None, None
    try:
        raw_info = st.secrets["GDRIVE_SERVICE_ACCOUNT"]
        # Handle both string JSON and TOML dict formats seamlessly
        if isinstance(raw_info, str):
            service_account_info = json.loads(raw_info)
        else:
            service_account_info = dict(raw_info)
            
        creds = service_account.Credentials.from_service_account_info(
            service_account_info, scopes=SCOPES
        )
        drive_service = build('drive', 'v3', credentials=creds)
        docs_service = build('docs', 'v1', credentials=creds)
        folder_id = st.secrets.get("GDRIVE_FOLDER_ID", "")
        return drive_service, docs_service, folder_id
    except Exception as e:
        st.error(f"Failed to authenticate Google Drive API: {e}")
        return None, None, None


def archive_to_google_drive(entry_meta, docx_bytes):
    """Uploads Word document to Drive and converts it into a native, editable Google Doc."""
    drive_service, _, folder_id = get_drive_services()
    if not drive_service:
        st.warning("⚠️ Google Drive integration not configured. Report won't be saved to Drive.")
        return None

    file_metadata = {
        'name': f"{entry_meta['project_name']} — {entry_meta['industry']} ({entry_meta['created_at'][:10]})",
        'mimeType': 'application/vnd.google-apps.document',  # Native Google Doc
        'parents': [folder_id] if folder_id else [],
        'appProperties': {
            'industry': entry_meta['industry'],
            'deliverable_type': entry_meta['deliverable_type'],
            'target_market': entry_meta['target_market'],
            'kb_categories': json.dumps(entry_meta['kb_categories']),
            'objective': entry_meta['objective'][:500],
            'created_at': entry_meta['created_at']
        }
    }

    media = MediaIoBaseUpload(
        io.BytesIO(docx_bytes),
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        resumable=True
    )

    gdoc = drive_service.files().create(
        body=file_metadata,
        media_body=media,
        fields='id, webViewLink'
    ).execute()

    return gdoc.get('webViewLink')


def load_drive_kb_index():
    """Lists files directly from the shared Google Drive folder."""
    drive_service, _, folder_id = get_drive_services()
    if not drive_service or not folder_id:
        return []

    query = f"'{folder_id}' in parents and mimeType = 'application/vnd.google-apps.document' and trashed = false"
    results = drive_service.files().list(
        q=query,
        pageSize=100,
        fields="files(id, name, webViewLink, createdTime, appProperties)",
        orderBy="createdTime desc"
    ).execute()

    files = results.get('files', [])
    index = []
    for f in files:
        props = f.get('appProperties', {})
        kb_cats = []
        if 'kb_categories' in props:
            try:
                kb_cats = json.loads(props['kb_categories'])
            except Exception:
                kb_cats = []

        index.append({
            "id": f['id'],
            "name": f['name'],
            "web_link": f['webViewLink'],
            "created_at": props.get('created_at', f.get('createdTime')),
            "industry": props.get('industry', 'Uncategorized'),
            "deliverable_type": props.get('deliverable_type', 'Report'),
            "target_market": props.get('target_market', 'General'),
            "kb_categories": kb_cats,
            "objective": props.get('objective', ''),
        })
    return index


# ----------------------------
# Sidebar Settings
# ----------------------------
with st.sidebar:
    st.header("⚙️ Report Settings")
    max_output_tokens = st.slider(
        "Max output tokens", min_value=4000, max_value=16000, value=12000, step=1000
    )
    max_searches = st.slider(
        "Max web searches allowed", min_value=5, max_value=30, value=20, step=5
    )
    st.caption("📌 Reports are natively saved as editable Google Docs in your shared Google Drive folder.")

# ----------------------------
# Tabs: Generate vs. Knowledge Base
# ----------------------------
tab_generate, tab_kb = st.tabs(["📝 Generate Report", "📚 Knowledge Base"])

# ============================================================
# TAB 1 — GENERATE REPORT
# ============================================================
with tab_generate:
    with st.form("comprehensive_research_form"):
        col1, col2 = st.columns(2)

        with col1:
            project_name = st.text_input("1. Project / Client Name", placeholder="e.g., Specialty Coffee Retail Audit 2026")
            industry = st.selectbox("2. Core Industry Focus", ["Fashion", "Interior & Architecture", "Food & Beverage (F&B)", "Medical", "Cross-Industry / General"])
            deliverable_type = st.selectbox("3. Deliverable Type", DELIVERABLE_TYPES)

        with col2:
            language = st.selectbox(
                "4. Output Language & Tone",
                [
                    "English (Formal Business / Research)",
                    "Modern Standard Arabic (فصحى)",
                    "Egyptian Natural Language Arabic (عامية مصرية احترافية)"
                ]
            )
            priority_focus = st.multiselect(
                "5. Primary Focus Areas (Select up to 3)",
                list(FOCUS_TO_CATEGORY.keys()),
                default=["Consumer Behavior & Pain Points", "Market Gaps & Business Opportunities"]
            )
            target_market = st.text_input("6. Target Market / Geographic Region", placeholder="e.g., Egypt, GCC Region, Global")

        objective = st.text_area("7. Research Objective & Business Question (Crucial)", height=100)
        specific_questions = st.text_area("8. Key Research Questions (Optional)", height=80)
        kb_categories = st.multiselect(
            "9. Knowledge Base Categories",
            KB_CATEGORIES,
            default=sorted({FOCUS_TO_CATEGORY[f] for f in priority_focus if f in FOCUS_TO_CATEGORY}) or ["Market"]
        )

        submit_button = st.form_submit_button("Generate Research Report")

    COMPREHENSIVE_SYSTEM_PROMPT = f"""
You are an automated Research & Insights Specialist operating under this company's Research & Insights Specialist Handbook. Your mission is to reduce uncertainty before decisions are made — not to make the decisions yourself. You provide reliable research, verified facts, meaningful insights, and valuable opportunities so the team can move with confidence instead of assumptions.

Today's date is {date.today().strftime('%B %d, %Y')}. Use this to judge what counts as current, and to phrase search queries correctly.

CORE MISSION (from the Handbook):
- Find reliable information. Verify its accuracy. Understand what it means. Identify opportunities. Organize knowledge.
- You are NOT expected to make strategic decisions. You ARE expected to make those decisions easier by explaining WHAT is happening, WHY it is happening, WHY it matters, and WHERE the opportunity exists.
- Do NOT design creative assets, write final social media copy, build full brand strategies, or plan content campaigns — that is out of scope. Leave execution to the creative/strategy teams. If your draft starts reading like a content plan instead of a research report, rewrite it as analysis, not as production instructions.

WORKFLOW YOU MUST FOLLOW INTERNALLY BEFORE WRITING (Handbook Steps 1-9):
1. Understand the objective given below — why this research, what decision it supports, who will use it.
2. Break it into specific research questions (don't research the topic broadly).
3. Collect information from MULTIPLE independent sources — never rely on a single site for any claim that matters.
4. Verify each piece of information: Is it accurate? Is it still relevant/current? Confirmed by another source? Does it apply to the target market specified? Explicitly check the publication date of every source you cite, and flag anything more than ~18 months old as potentially dated. Where possible, trace a claim back to its original/primary source rather than resting on a secondary summary of it.
5. Organize findings into the Handbook's topical categories: Market, Consumer, Competitors, Content, Trends, Technology, Statistics. A single finding can belong to more than one.
6. Extract insights — after every important finding, ask "what does this mean for our business?"
7. Identify opportunities — what gap, pain point, or advantage does this create?
8. Build recommendations that are a natural conclusion of the findings above, never an assumption stated before the evidence supporting it.
9. (Handled outside your output — the application archives your finished report automatically.)

RESEARCH DEPTH & RIGOR:
- This is a paid, professional deliverable. Run as many distinct web searches as needed (rarely fewer than 8-10 for a full report) — do not stop after 1-2. Search each sub-question separately rather than combining topics into one broad query.
- Prefer concrete numbers, percentages, dates, and named entities over vague generalizations ("growing rapidly," "consumers increasingly prefer"). Every FACT must be traceable to a specific, named source.
- Where sources disagree, state the discrepancy rather than silently picking one.
- Never quote more than a short phrase (under ~15 words) from any single source, and never more than one such phrase per source — paraphrase everything else in your own words.

SOURCE PRIORITY (Handbook standard — always prefer the highest tier available):
- Priority 1: Official platforms, official documentation, peer-reviewed research papers, government publications.
- Priority 2: Industry reports and trusted research organizations.
- Priority 3: Trusted marketing and business publications.
- Priority 4: Industry experts.
- Priority 5: Communities and discussions (consumer-sentiment signal only — never the sole basis for a Fact).

CLASSIFICATION FRAMEWORK (STRICT SEPARATION — every substantive statement must be tagged as exactly one of these):
- FACT: A verified piece of information supported by reliable, cited sources.
- OBSERVATION: A noticeable pattern or repeated behavior identified through the research.
- INSIGHT: An explanation of WHY something matters and how it affects the business — not just what happened.
- OPPORTUNITY: A possible area where the company or client can create value.
- RECOMMENDATION: A suggested, evidence-backed action — never personal opinion, never stated before the findings that justify it.

PRIORITIZATION (do not present every finding as equally important):
- For every Insight, Opportunity, and Recommendation, assign a priority of [High], [Medium], or [Low] based on business impact and urgency, and lead each subsection with the [High] items first.
- In the Executive Summary, explicitly call out the single highest-priority finding and the single highest-priority recommendation — the reader should know in 10 seconds what matters most.

REQUIRED DELIVERABLE STRUCTURE — this is the Handbook's Research Standard. Use exactly these section headers, in exactly this order, nothing added or removed:
1. Research Objective
2. Research Questions
3. Executive Summary
4. Key Findings — organized under topical subheadings from this exact set: Market, Consumer, Competitors, Content, Trends, Technology, Statistics (omit any category with no findings). Within each subheading, tag every statement inline with its classification in bold brackets, e.g. "**[Fact — Priority 2]** ...".
5. Supporting Data — a Markdown table of the key quantitative data points referenced above (columns: Data Point, Value, Source, Source Tier, Publication Date).
6. Key Insights — the strategic "why it matters" analysis, each tagged with a priority.
7. Opportunities — each tagged with a priority.
8. Recommendations — each tagged with a priority, and each one traceable to a specific finding or insight above it.
9. References — every source actually used, grouped under "Priority 1" through "Priority 5" headers, with source name, URL, and the publication date you found for it (write "date not stated" if unavailable).

FORMATTING:
- Output ONLY the finished report in clean Markdown — no preamble like "Here is your report."
- End with this exact note: "This report covers WHAT is happening, WHY it matters, and WHERE the opportunity is. Creative execution (design, copywriting, campaign planning) is out of scope and left to the strategy/creative team."

LANGUAGE & TONE:
- IF Language = English (Formal Business / Research): formal, sharp, authoritative agency-research English.
- IF Language = Modern Standard Arabic (فصحى): precise, standard corporate Arabic, clear analytical terminology.
- IF Language = Egyptian Natural Language Arabic (عامية مصرية احترافية): professional, clean Egyptian Arabic suitable for local agency teams (عامية مصرية راقية ومفهومة لفرق العمل). Keep technical marketing/research terms intact in English (Insights, Conversion, Benchmarks, Target Audience, Positioning, CAGR, SKU, Fact, Observation, Opportunity, Recommendation) while making the surrounding prose read naturally, not like a translation. Section headers and body text should all be in Egyptian Arabic.
"""

    def generate_report(client, user_prompt, system_prompt, max_tokens, max_searches_n, status):
        messages = [{"role": "user", "content": user_prompt}]
        full_text = ""
        searches_seen = []

        for turn in range(4):
            status.update(label=f"Researching and drafting (pass {turn + 1})...")
            response = client.messages.create(
                model=MODEL_NAME,
                max_tokens=max_tokens,
                system=system_prompt,
                messages=messages,
                tools=[{
                    "type": "web_search_20250305",
                    "name": "web_search",
                    "max_uses": max_searches_n
                }],
            )

            for block in response.content:
                btype = getattr(block, "type", None)
                if btype == "server_tool_use" and getattr(block, "name", "") == "web_search":
                    query = getattr(block, "input", {}).get("query")
                    if query:
                        searches_seen.append(query)

            turn_text = "".join(
                block.text for block in response.content if getattr(block, "type", None) == "text"
            )
            full_text += turn_text

            if response.stop_reason != "max_tokens":
                break

            messages.append({"role": "assistant", "content": response.content})
            messages.append({
                "role": "user",
                "content": "Continue the report exactly where you left off. Do not repeat any content already written, do not restart headers already completed, and do not add any preamble."
            })

        return full_text, searches_seen

    # Word conversion helpers
    def _set_paragraph_rtl(paragraph):
        pPr = paragraph._p.get_or_add_pPr()
        bidi = OxmlElement('w:bidi')
        bidi.set(qn('w:val'), '1')
        pPr.append(bidi)
        paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    def _add_run_with_bold(paragraph, text, rtl=False):
        parts = re.split(r'(\*\*.*?\*\*)', text)
        for part in parts:
            if not part:
                continue
            is_bold = part.startswith('**') and part.endswith('**')
            content = part[2:-2] if is_bold else part
            run = paragraph.add_run(content)
            run.bold = is_bold
            if rtl:
                rPr = run._r.get_or_add_rPr()
                rtl_el = OxmlElement('w:rtl')
                rtl_el.set(qn('w:val'), '1')
                rPr.append(rtl_el)

    def _style_table_header_cell(cell):
        shading = OxmlElement('w:shd')
        shading.set(qn('w:val'), 'clear')
        shading.set(qn('w:fill'), 'D9D9D9')
        cell._tc.get_or_add_tcPr().append(shading)

    def markdown_to_docx(markdown_text: str, rtl: bool = False, base_font: str = None) -> bytes:
        doc = Document()
        normal_style = doc.styles['Normal']
        if base_font:
            normal_style.font.name = base_font
            rPr = normal_style.element.get_or_add_rPr()
            rFonts = rPr.find(qn('w:rFonts'))
            if rFonts is None:
                rFonts = OxmlElement('w:rFonts')
                rPr.append(rFonts)
            rFonts.set(qn('w:cs'), base_font)
        normal_style.font.size = Pt(11)

        if rtl:
            sectPr = doc.sections[0]._sectPr
            bidi = OxmlElement('w:bidi')
            sectPr.append(bidi)

        lines = markdown_text.splitlines()
        i = 0
        n = len(lines)

        while i < n:
            line = lines[i].rstrip()
            if not line.strip():
                i += 1
                continue
            if re.match(r'^-{3,}$', line.strip()):
                i += 1
                continue

            header_match = re.match(r'^(#{1,4})\s+(.*)', line)
            if header_match:
                level = len(header_match.group(1))
                text = header_match.group(2).strip()
                p = doc.add_heading(level=min(level, 4))
                _add_run_with_bold(p, text, rtl=rtl)
                if rtl:
                    _set_paragraph_rtl(p)
                i += 1
                continue

            if line.strip().startswith('|'):
                table_lines = []
                while i < n and lines[i].strip().startswith('|'):
                    table_lines.append(lines[i].strip())
                    i += 1
                rows = [
                    [c.strip() for c in tl.strip('|').split('|')]
                    for tl in table_lines
                    if not re.match(r'^\|?\s*:?-{2,}:?\s*(\|\s*:?-{2,}:?\s*)*\|?$', tl)
                ]
                if rows:
                    ncols = len(rows[0])
                    table = doc.add_table(rows=0, cols=ncols)
                    table.alignment = WD_TABLE_ALIGNMENT.CENTER
                    table.style = 'Light Grid Accent 1'
                    for r_idx, row_cells in enumerate(rows):
                        row = table.add_row()
                        for c_idx in range(ncols):
                            cell_text = row_cells[c_idx] if c_idx < len(row_cells) else ""
                            cell = row.cells[c_idx]
                            cell.paragraphs[0].text = ""
                            p = cell.paragraphs[0]
                            _add_run_with_bold(p, cell_text, rtl=rtl)
                            if rtl:
                                _set_paragraph_rtl(p)
                            if r_idx == 0:
                                for run in p.runs:
                                    run.bold = True
                                _style_table_header_cell(cell)
                continue

            bullet_match = re.match(r'^[\-\*]\s+(.*)', line)
            if bullet_match:
                p = doc.add_paragraph(style='List Bullet')
                _add_run_with_bold(p, bullet_match.group(1), rtl=rtl)
                if rtl:
                    _set_paragraph_rtl(p)
                i += 1
                continue

            numbered_match = re.match(r'^\d+\.\s+(.*)', line)
            if numbered_match:
                p = doc.add_paragraph(style='List Number')
                _add_run_with_bold(p, numbered_match.group(1), rtl=rtl)
                if rtl:
                    _set_paragraph_rtl(p)
                i += 1
                continue

            p = doc.add_paragraph()
            _add_run_with_bold(p, line.strip(), rtl=rtl)
            if rtl:
                _set_paragraph_rtl(p)
            i += 1

        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        return buffer.getvalue()

    if submit_button:
        if not objective:
            st.warning("⚠️ Please provide a Research Objective before generating.")
        else:
            client = anthropic.Anthropic(api_key=api_key)

            user_prompt = f"""
Execute research report generation based on the following standardized input:

- Project / Client Name: {project_name if project_name else "N/A"}
- Core Industry: {industry}
- Deliverable Type: {deliverable_type}
- Target Market: {target_market if target_market else "General / Unspecified"}
- Priority Focus Areas: {', '.join(priority_focus) if priority_focus else "None specified"}
- Preferred Language: {language}

BUSINESS OBJECTIVE:
{objective}

SPECIFIC RESEARCH QUESTIONS TO ANSWER:
{specific_questions if specific_questions else "Derive the key research questions yourself from the business objective above before researching."}
"""

            with st.status("Generating report...", expanded=True) as status:
                try:
                    report_text, searches_used = generate_report(
                        client, user_prompt, COMPREHENSIVE_SYSTEM_PROMPT,
                        max_output_tokens, max_searches, status
                    )
                    status.update(label="Report generated.", state="complete")
                except Exception as e:
                    status.update(label="Generation failed.", state="error")
                    st.error(f"Execution Error: {str(e)}")
                    st.stop()

            if searches_used:
                with st.expander(f"🔎 {len(searches_used)} web searches performed"):
                    for q in searches_used:
                        st.write(f"- {q}")

            st.success("Report Generated Successfully!")
            st.markdown("---")
            st.markdown(report_text)

            is_arabic = "Arabic" in language
            docx_bytes = None
            try:
                docx_bytes = markdown_to_docx(
                    report_text,
                    rtl=is_arabic,
                    base_font="Arial" if is_arabic else "Calibri"
                )
            except Exception as e:
                st.warning(f"Word export failed ({e})")

            # Google Drive Archival
            entry_meta = {
                "created_at": datetime.now().isoformat(timespec="seconds"),
                "project_name": project_name or "Untitled",
                "industry": industry,
                "deliverable_type": deliverable_type,
                "target_market": target_market or "General / Unspecified",
                "language": language,
                "kb_categories": kb_categories,
                "objective": objective,
            }
            
            gdoc_url = None
            if docx_bytes:
                gdoc_url = archive_to_google_drive(entry_meta, docx_bytes)

            if gdoc_url:
                st.success(f"📄 **Archived to Google Drive!** [Open Editable Google Doc]({gdoc_url})")
            else:
                st.info("📚 Generated report ready for local download below.")

            st.markdown("---")
            safe_name = (project_name or "Research").strip().replace(" ", "_")

            col_dl1, col_dl2, col_dl3 = st.columns(3)
            with col_dl1:
                st.download_button("📄 Download (.txt)", data=report_text, file_name=f"{safe_name}_Report.txt", mime="text/plain")
            with col_dl2:
                st.download_button("📝 Download (.md)", data=report_text, file_name=f"{safe_name}_Report.md", mime="text/markdown")
            with col_dl3:
                if docx_bytes is not None:
                    st.download_button("📘 Download (.docx)", data=docx_bytes, file_name=f"{safe_name}_Report.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

# ============================================================
# TAB 2 — KNOWLEDGE BASE (Google Drive Backend)
# ============================================================
with tab_kb:
    st.subheader("📚 Company Knowledge Base (Google Drive)")
    st.caption("All reports are live, editable Google Docs saved directly inside your team's Google Drive folder.")

    index = load_drive_kb_index()

    if not index:
        st.info("No archived research found in Google Drive folder, or Google Drive API credentials are not set up.")
    else:
        col_f1, col_f2, col_f3 = st.columns(3)
        with col_f1:
            filter_industry = st.multiselect("Filter by Industry", sorted({e["industry"] for e in index if e.get("industry")}))
        with col_f2:
            filter_category = st.multiselect("Filter by KB Category", KB_CATEGORIES)
        with col_f3:
            search_term = st.text_input("Search report title / objective", placeholder="e.g., cold brew")

        filtered = index
        if filter_industry:
            filtered = [e for e in filtered if e["industry"] in filter_industry]
        if filter_category:
            filtered = [e for e in filtered if any(c in e.get("kb_categories", []) for c in filter_category)]
        if search_term:
            term = search_term.lower()
            filtered = [
                e for e in filtered
                if term in e.get("name", "").lower() or term in e.get("objective", "").lower()
            ]

        st.write(f"**{len(filtered)}** of {len(index)} Google Docs match your filters.")
        st.markdown("---")

        for entry in filtered:
            with st.expander(f"📄 {entry['name']}"):
                st.write(f"**Industry:** {entry['industry']}")
                st.write(f"**Deliverable Type:** {entry['deliverable_type']}")
                st.write(f"**Target Market:** {entry['target_market']}")
                st.write(f"**KB Categories:** {', '.join(entry.get('kb_categories', [])) or '—'}")
                if entry.get("objective"):
                    st.write(f"**Objective:** {entry['objective']}")
                
                st.markdown(f"🔗 **[Open & Edit in Google Docs]({entry['web_link']})**")
