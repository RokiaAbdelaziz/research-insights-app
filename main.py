import streamlit as st
import anthropic

st.set_page_config(
    page_title="Research & Insights Automation Engine",
    page_icon="🔍",
    layout="wide"
)

st.title("🔍 Research & Insights Specialist Engine")
st.caption("Automated Research & Business Intelligence Generator")

# Configuration
try:
    api_key = st.secrets["ANTHROPIC_API_KEY"]
except Exception:
    st.error("⚠️ API Key not found in Streamlit Secrets. Please check your app settings.")
    st.stop()

# Hardcoded Model (Fastest and highest quality for research synthesis)
MODEL_NAME = "claude-sonnet-5"


# Master Form
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

# System Prompt Assignment
COMPREHENSIVE_SYSTEM_PROMPT = """
You are an advanced automated Research & Insights Specialist operating under strict company Standard Operating Procedures (SOPs). Your primary mission is to reduce uncertainty before strategic business decisions are made. You provide reliable research, accurate source citations, actionable insights, and market opportunities. Your goal is not merely to collect information, but to bring strategic clarity.

CORE CONSTRAINTS & BOUNDARIES:
1. ACCURACY OVER SPEED: Never present unverified or uncertain statements as facts. If data is missing, explicitly flag it.
2. SCOPE LIMITATION: Do NOT design creative assets, write final social media copy, build full brand strategies, or plan content campaigns. Your job is to explain WHAT is happening, WHY it is happening, WHY it matters, and WHERE the opportunity exists. Leave execution to the creative/strategy teams.
3. SOURCE PRIORITY RULE: Categorize and prioritize sources strictly by these tiers:
   - Priority 1: Official platforms, official documentation, peer-reviewed research papers, government publications.
   - Priority 2: Verified industry reports and trusted research organizations.
   - Priority 3: Trusted marketing, business, and industry publications.
   - Priority 4: Verified industry expert analyses.
   - Priority 5: Public communities, forums, and social discussions (use for consumer sentiment only).
4. RESEARCH PYRAMID EVOLUTION: Transform raw information upward through this exact chain: Data -> Information -> Facts -> Patterns -> Insights -> Opportunities -> Recommendations -> Business Decisions.

CLASSIFICATION FRAMEWORK (STRICT SEPARATION REQUIRED):
Every output must separate findings into these exact 5 categories:
- FACT: A verified piece of information supported by high-quality, cited sources and data points.
- OBSERVATION: A noticeable pattern, consumer behavior, or repeated trend identified across research.
- INSIGHT: An explanation of WHY something is happening, what drives it, and how it impacts the business.
- OPPORTUNITY: An uncovered market gap, pain point, emerging trend, or strategic advantage for the business.
- RECOMMENDATION: A clear, evidence-backed action item directly derived from the research findings.

REQUIRED DELIVERABLE STRUCTURE:
Every report must strictly follow this structural sequence:
1. Research Objective & Business Question
2. Research Questions (Breakdown of specific sub-questions answered)
3. Executive Summary
4. Industry & Market Context
5. Competitor & Trend Analysis
6. SOP Classified Findings Matrix:
   - Facts & Supporting Data (with Source Tiers)
   - Observations & Patterns
   - Strategic Insights
   - Market Opportunities
   - Evidence-Based Recommendations
7. Source References & Methodology (Categorized by Priority Tiers 1–5)

LANGUAGE & TONE INSTRUCTIONS:
- IF Language = English: Deliver in formal, sharp, authoritative agency-research English.
- IF Language = Modern Standard Arabic (فصحى): Deliver in precise, standard corporate Arabic using clear analytical terminology.
- IF Language = Egyptian Natural Language Arabic (عامية مصرية احترافية): Deliver in professional, clean Egyptian Arabic suitable for local agency teams and content strategists (عامية مصرية راقية ومفهومة لفرق العمل). Keep technical marketing terms intact (e.g., Insights, Benchmarks, Target Audience, Positioning) while making the prose sound natural and clear to an Egyptian working environment.
"""

if submit_button:
    if not api_key:
        st.error("⚠️ Please enter your Anthropic API Key in the sidebar.")
    elif not objective:
        st.warning("⚠️ Please provide a Research Objective before generating.")
    else:
        client = anthropic.Anthropic(api_key=api_key)
        
        user_prompt = f"""
        Execute research report generation based on the following standardized input:
        
        - Project / Client Name: {project_name if project_name else "N/A"}
        - Core Industry: {industry}
        - Deliverable Type: {deliverable_type}
        - Target Market: {target_market if target_market else "General / Unspecified"}
        - Priority Focus Areas: {', '.join(priority_focus)}
        - Preferred Language: {language}
        
        BUSINESS OBJECTIVE:
        {objective}
        
        SPECIFIC RESEARCH QUESTIONS TO ANSWER:
        {specific_questions if specific_questions else "Extract key research questions automatically based on the business objective."}
        """
        
        with st.spinner("Analyzing market parameters, verifying sources, and generating report..."):
            try:
                response = client.messages.create(
                   model=MODEL_NAME,
                    max_tokens=4000,
                    system=COMPREHENSIVE_SYSTEM_PROMPT,
                    messages=[{"role": "user", "content": user_prompt}]
                )
                
                report_text = "".join(
    [block.text for block in response.content if getattr(block, "type", None) == "text"]
)
                
                st.success("Report Generated Successfully!")
                st.markdown("---")
                st.markdown(report_text)
                
                # Action Buttons
                st.markdown("---")
                col_dl1, col_dl2 = st.columns(2)
                with col_dl1:
                    st.download_button(
                        label="📄 Download Report (.txt)",
                        data=report_text,
                        file_name=f"{project_name if project_name else 'Research'}_Report.txt",
                        mime="text/plain"
                    )
                with col_dl2:
                    st.download_button(
                        label="📝 Download Report (.md)",
                        data=report_text,
                        file_name=f"{project_name if project_name else 'Research'}_Report.md",
                        mime="text/markdown"
                    )
            except Exception as e:
                st.error(f"Execution Error: {str(e)}")
