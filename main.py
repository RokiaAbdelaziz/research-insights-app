import streamlit as st
import anthropic
from datetime import date

st.set_page_config(
    page_title="Research & Insights Automation Engine",
    page_icon="🔍",
    layout="wide"
)

st.title("🔍 Research & Insights Specialist Engine")
st.caption("Automated Research & Business Intelligence Generator")

# ----------------------------
# Configuration
# ----------------------------
try:
    api_key = st.secrets["ANTHROPIC_API_KEY"]
except Exception:
    st.error("⚠️ API Key not found in Streamlit Secrets. Please check your app settings.")
    st.stop()

MODEL_NAME = "claude-sonnet-5"

# Output length / research depth controls (exposed so you can tune per-report
# instead of hardcoding a value that silently truncates long reports)
with st.sidebar:
    st.header("⚙️ Report Settings")
    max_output_tokens = st.slider(
        "Max output tokens", min_value=4000, max_value=16000, value=12000, step=1000,
        help="A full SOP-style report (Facts/Observations/Insights/Opportunities/"
             "Recommendations + source list) typically needs 6,000-10,000 tokens. "
             "Raise this if reports keep getting cut off."
    )
    max_searches = st.slider(
        "Max web searches allowed", min_value=5, max_value=30, value=20, step=5,
        help="Comprehensive market/competitor reports need 10-20+ searches. "
             "Too low a cap forces shallow, under-researched output."
    )

# ----------------------------
# Master Form
# ----------------------------
with st.form("comprehensive_research_form"):
    col1, col2 = st.columns(2)

    with col1:
        project_name = st.text_input(
            "1. Project / Client Name",
            placeholder="e.g., Specialty Coffee Retail Audit 2026"
        )

        industry = st.selectbox(
            "2. Core Industry Focus",
            ["Fashion", "Interior & Architecture", "Food & Beverage (F&B)", "Medical", "Cross-Industry / General"]
        )

        deliverable_type = st.selectbox(
            "3. Deliverable Type",
            [
                "General Research Summary",
                "Competitor Research Report",
                "Industry Update Report",
                "Trend Collection & Opportunity Audit",
                "Client Onboarding Research File"
            ]
        )

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
            [
                "Consumer Behavior & Pain Points",
                "Competitor Positioning & Branding",
                "Packaging & Visual Identity Trends",
                "Social Media & Platform Updates",
                "AI Tools & MarTech Innovations",
                "Market Gaps & Business Opportunities"
            ],
            default=["Consumer Behavior & Pain Points", "Market Gaps & Business Opportunities"]
        )

        target_market = st.text_input(
            "6. Target Market / Geographic Region",
            placeholder="e.g., Egypt, GCC Region, Global"
        )

    objective = st.text_area(
        "7. Research Objective & Business Question (Crucial)",
        placeholder="Why are we conducting this research? What specific business question or decision will this output support?",
        height=100
    )

    specific_questions = st.text_area(
        "8. Key Research Questions (Optional)",
        placeholder="List specific questions to answer (e.g., What are top competitors charging? What packaging material is trending?)",
        height=80
    )

    submit_button = st.form_submit_button("Generate Research Report")

# ----------------------------
# System Prompt
# ----------------------------
COMPREHENSIVE_SYSTEM_PROMPT = f"""
You are an advanced automated Research & Insights Specialist operating under strict company Standard Operating Procedures (SOPs). Your primary mission is to reduce uncertainty before strategic business decisions are made. You provide reliable research, accurate source citations, actionable insights, and market opportunities. Your goal is not merely to collect information, but to bring strategic clarity.

Today's date is {date.today().strftime('%B %d, %Y')}. Use this to judge what counts as current, and to phrase search queries correctly (do not search for last year when this year's data is what's needed).

CORE CONSTRAINTS & BOUNDARIES:
1. ACCURACY OVER SPEED: Never present unverified or uncertain statements as facts. If data is missing or a claim could not be verified, explicitly flag it as such rather than omitting the gap silently.
2. SCOPE LIMITATION: Do NOT design creative assets, write final social media copy, build full brand strategies, or plan content campaigns. Your job is to explain WHAT is happening, WHY it is happening, WHY it matters, and WHERE the opportunity exists. Leave execution to the creative/strategy teams.
3. SOURCE PRIORITY RULE: Categorize and prioritize sources strictly by these tiers:
   - Priority 1: Official platforms, official documentation, peer-reviewed research papers, government publications.
   - Priority 2: Verified industry reports and trusted research organizations.
   - Priority 3: Trusted marketing, business, and industry publications.
   - Priority 4: Verified industry expert analyses.
   - Priority 5: Public communities, forums, and social discussions (use for consumer sentiment only).
4. RESEARCH PYRAMID EVOLUTION: Transform raw information upward through this exact chain: Data -> Information -> Facts -> Patterns -> Insights -> Opportunities -> Recommendations -> Business Decisions.
5. RESEARCH DEPTH: This is a paid, professional deliverable, not a quick answer. Perform as many distinct web searches as you need (do not stop after 1-2) to independently verify market sizing, competitor specifics, regulatory details, and local context. Search each distinct sub-topic separately rather than combining them into one broad query. Prefer the most recent data available. Where sources disagree, note the discrepancy rather than picking one silently.
6. SPECIFICITY: Prefer concrete numbers, percentages, dates, and named entities over vague generalizations like "growing rapidly" or "consumers increasingly prefer." Every FACT must be traceable to a specific source.
7. COPYRIGHT: Never quote more than a short phrase (under ~15 words) from any single source, and never more than one such phrase per source. Always paraphrase in your own words otherwise.

CLASSIFICATION FRAMEWORK (STRICT SEPARATION REQUIRED):
Every output must separate findings into these exact 5 categories:
- FACT: A verified piece of information supported by high-quality, cited sources and data points.
- OBSERVATION: A noticeable pattern, consumer behavior, or repeated trend identified across research.
- INSIGHT: An explanation of WHY something is happening, what drives it, and how it impacts the business.
- OPPORTUNITY: An uncovered market gap, pain point, emerging trend, or strategic advantage for the business.
- RECOMMENDATION: A clear, evidence-backed action item directly derived from the research findings.

REQUIRED DELIVERABLE STRUCTURE (use these exact section headers, in this order):
1. Research Objective & Business Question
2. Research Questions (breakdown of specific sub-questions answered)
3. Executive Summary
4. Industry & Market Context
5. Competitor & Trend Analysis
6. SOP Classified Findings Matrix, with these five clearly labeled subsections:
   - Facts & Supporting Data (each fact tagged with its Source Tier, e.g. "Priority 2")
   - Observations & Patterns
   - Strategic Insights
   - Market Opportunities
   - Evidence-Based Recommendations
7. Source References & Methodology — list every source actually used, grouped under headers "Priority 1", "Priority 2", "Priority 3", "Priority 4", "Priority 5", with source name and URL. Also include a short paragraph describing the research approach taken.

FORMATTING:
- Output ONLY the finished report in clean Markdown (headers, tables where useful, no preamble like "Here is your report").
- Use Markdown tables for the Facts subsection (columns: #, Fact, Source Tier).
- End with the note: "This report covers WHAT, WHY, and WHERE the opportunity is. Creative execution (design, copywriting, campaign planning) is out of scope and left to the creative/strategy team."

LANGUAGE & TONE INSTRUCTIONS:
- IF Language = English (Formal Business / Research): Deliver in formal, sharp, authoritative agency-research English.
- IF Language = Modern Standard Arabic (فصحى): Deliver in precise, standard corporate Arabic using clear analytical terminology.
- IF Language = Egyptian Natural Language Arabic (عامية مصرية احترافية): Deliver in professional, clean Egyptian Arabic suitable for local agency teams and content strategists (عامية مصرية راقية ومفهومة لفرق العمل). Keep technical marketing/research terms intact in English (e.g., Insights, Conversion, Benchmarks, Target Audience, Positioning, CAGR, SKU) while making the surrounding prose sound natural to an Egyptian working environment. Write ALL section headers and body text in Egyptian Arabic (not just a translated skeleton) — this must read like a native deliverable, not a translation.
"""

# ----------------------------
# Helper: run one generation, continuing automatically if truncated
# ----------------------------
def generate_report(client, user_prompt, system_prompt, max_tokens, max_searches, status):
    messages = [{"role": "user", "content": user_prompt}]
    full_text = ""
    searches_seen = []

    # Loop to handle max_tokens truncation by asking the model to continue.
    # Also surfaces which web searches were run, for transparency.
    for turn in range(4):  # hard safety cap on continuation turns
        status.update(label=f"Researching and drafting (pass {turn + 1})...")
        response = client.messages.create(
            model=MODEL_NAME,
            max_tokens=max_tokens,
            system=system_prompt,
            messages=messages,
            tools=[{
                "type": "web_search_20250305",
                "name": "web_search",
                "max_uses": max_searches
            }],
        )

        # Collect any web_search queries the model issued, for transparency.
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

        # Truncated: append what we have and ask the model to continue exactly
        # where it left off, so the final report isn't silently cut short.
        messages.append({"role": "assistant", "content": response.content})
        messages.append({
            "role": "user",
            "content": "Continue the report exactly where you left off. Do not repeat "
                       "any content already written, do not restart headers already "
                       "completed, and do not add any preamble."
        })

    return full_text, searches_seen


# ----------------------------
# Execution
# ----------------------------
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

        # Action Buttons
        st.markdown("---")
        col_dl1, col_dl2 = st.columns(2)
        safe_name = (project_name or "Research").strip().replace(" ", "_")
        with col_dl1:
            st.download_button(
                label="📄 Download Report (.txt)",
                data=report_text,
                file_name=f"{safe_name}_Report.txt",
                mime="text/plain"
            )
        with col_dl2:
            st.download_button(
                label="📝 Download Report (.md)",
                data=report_text,
                file_name=f"{safe_name}_Report.md",
                mime="text/markdown"
            )
