"""
app.py
──────
SiteInsight · Streamlit Dashboard — Redesigned
Theme: White + Amber/Orange accents
Navigation: Left sidebar with section selector
Changes:
  1. White + amber theme throughout
  2. Left sidebar navigation, summary metrics at top
  3. Top risks show phase + activity, not just ID
  4. Detection flags as compact colour-coded table
  5. Site photos fixed (use_container_width) + 4 curated images
  6. MOM plain language + proper .docx download
  7. S-curve added to graphs tab

Run:
    python -m streamlit run app.py
"""

import os
import io
import json
import smtplib
import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
import matplotlib
matplotlib.use("Agg")

from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from supabase import create_client, Client
from groq import Groq
from dotenv import load_dotenv
from datetime import datetime, date

load_dotenv()

SUPABASE_URL   = os.environ["SUPABASE_URL"]
SUPABASE_KEY   = os.environ["SUPABASE_KEY"]
GROQ_API_KEY   = os.environ["GROQ_API_KEY"]
EMAIL_SENDER   = os.environ.get("EMAIL_SENDER", "")
EMAIL_PASSWORD = os.environ.get("EMAIL_PASSWORD", "")
GROQ_MODEL     = "llama-3.3-70b-versatile"

# ── Theme colours ─────────────────────────────────────────────
AMBER  = "#E87722"
AMBER2 = "#F9A825"
NAVY   = "#1E2761"
RED    = "#D32F2F"
GREEN  = "#2E7D32"
GRAY   = "#666666"
LGRAY  = "#F5F5F5"

# ── Page config ───────────────────────────────────────────────
st.set_page_config(
    page_title="SiteInsight · WPR Intelligence",
    page_icon="🏗",
    layout="wide",
)

# ── Custom CSS — white + amber theme ─────────────────────────
st.markdown("""
<style>
[data-testid="stSidebar"] { background-color: #1E2761 !important; }
[data-testid="stSidebar"] * { color: #FFFFFF !important; }
.badge-high   { background:#FFEBEE; color:#C62828; padding:2px 8px; border-radius:4px; font-size:0.75rem; font-weight:600; }
.badge-medium { background:#FFF8E1; color:#E65100; padding:2px 8px; border-radius:4px; font-size:0.75rem; font-weight:600; }
.badge-low    { background:#E8F5E9; color:#1B5E20; padding:2px 8px; border-radius:4px; font-size:0.75rem; font-weight:600; }
</style>
""", unsafe_allow_html=True)

# ── Site photos — 4 curated realistic images ─────────────────
SITE_PHOTOS = [
    {
        "url": "https://images.unsplash.com/photo-1504307651254-35680f356dfd?w=600&q=80",
        "caption": "Phase I · Basement formwork and shuttering in progress",
        "phase": "Phase I",
    },
    {
        "url": "https://images.unsplash.com/photo-1590579491624-f98f36d4c763?w=600&q=80",
        "caption": "Phase II · Reinforcement steel works and bar bending",
        "phase": "Phase II",
    },
    {
        "url": "https://images.unsplash.com/photo-1581094794329-c8112a89af12?w=600&q=80",
        "caption": "Phase III · Excavation and earthworks",
        "phase": "Phase III",
    },
    {
        "url": "https://images.unsplash.com/photo-1558618666-fcd25c85cd64?w=600&q=80",
        "caption": "Site · Waterproofing and finishing works",
        "phase": "Site",
    },
]

# ── Clients ───────────────────────────────────────────────────
@st.cache_resource
def get_supabase() -> Client:
    return create_client(SUPABASE_URL, SUPABASE_KEY)

@st.cache_resource
def get_groq():
    return Groq(api_key=GROQ_API_KEY)

supabase    = get_supabase()
groq_client = get_groq()

# ── Data loaders ──────────────────────────────────────────────
@st.cache_data(ttl=60)
def load_weeks():
    r = supabase.table("wpr_headers").select(
        "week_number,load_type").order("week_number").execute()
    return r.data or []

@st.cache_data(ttl=60)
def load_header(week):
    r = supabase.table("wpr_headers").select("*").eq("week_number", week).execute()
    return r.data[0] if r.data else {}

@st.cache_data(ttl=60)
def load_activities(week):
    r = supabase.table("wpr_activities").select("*").eq(
        "week_number", week).order("act_id").execute()
    return pd.DataFrame(r.data) if r.data else pd.DataFrame()

@st.cache_data(ttl=30)
def load_detections(week):
    r = supabase.table("detection_results").select("*").eq(
        "week_number", week).order("rule_id").execute()
    return pd.DataFrame(r.data) if r.data else pd.DataFrame()

@st.cache_data(ttl=60)
def load_narrative(week):
    r = supabase.table("ai_narratives").select("*").eq("week_number", week).execute()
    return r.data[0] if r.data else {}

@st.cache_data(ttl=60)
def load_dq_flags(week):
    r = supabase.table("dq_flags").select("*").eq("week_number", week).execute()
    return pd.DataFrame(r.data) if r.data else pd.DataFrame()

@st.cache_data(ttl=300)
def load_all_history():
    r = supabase.table("wpr_activities").select(
        "week_number,act_id,phase,activity,variance_pct,weeks_slip,"
        "cum_actual_pct,cum_planned_pct,delay_reason,is_critical_path,"
        "total_scope,responsible_person"
    ).order("week_number").execute()
    return pd.DataFrame(r.data) if r.data else pd.DataFrame()

@st.cache_data(ttl=300)
def load_baselines():
    r = supabase.table("baselines").select("*").order("id").execute()
    return pd.DataFrame(r.data) if r.data else pd.DataFrame()

# ── Groq helper ───────────────────────────────────────────────
def call_groq(prompt: str, max_tokens: int = 1500) -> str:
    response = groq_client.chat.completions.create(
        model=GROQ_MODEL,
        messages=[{"role": "user", "content": prompt}],
        temperature=0.3,
        max_tokens=max_tokens,
    )
    return response.choices[0].message.content.strip()

# ── Format narrative list fields ─────────────────────────────
def format_narrative_list(raw) -> list:
    """Convert Groq JSON list or string into clean Python list."""
    if isinstance(raw, list):
        return raw
    if isinstance(raw, str):
        raw = raw.strip()
        if raw.startswith("["):
            try:
                parsed = json.loads(raw)
                if isinstance(parsed, list):
                    return parsed
            except Exception:
                pass
        # Plain text — split by newline or numbered items
        lines = [l.strip().lstrip("-•123456789.)").strip()
                 for l in raw.split("\n") if l.strip()]
        return [l for l in lines if l]
    return [str(raw)]

# ── Email helper ──────────────────────────────────────────────
def send_email(to_email: str, subject: str, body: str) -> bool:
    try:
        msg = MIMEMultipart("alternative")
        msg["Subject"] = subject
        msg["From"]    = EMAIL_SENDER
        msg["To"]      = to_email
        msg.attach(MIMEText(body, "plain"))
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(EMAIL_SENDER, EMAIL_PASSWORD)
            server.sendmail(EMAIL_SENDER, to_email, msg.as_string())
        return True
    except Exception as e:
        st.error(f"Email failed to {to_email}: {e}")
        return False

# ── Helper functions ──────────────────────────────────────────
def compute_weekly_velocity(hdf, aid):
    act_hist = hdf[hdf["act_id"] == aid].sort_values("week_number")
    if len(act_hist) < 2:
        return 0.0
    return float(act_hist["cum_actual_pct"].diff().dropna().mean())

def compute_projected_finish(current_pct, current_week, weekly_rate, total_weeks=80):
    if weekly_rate <= 0:
        return total_weeks
    remaining = 1.0 - current_pct
    return int(current_week + remaining / weekly_rate)

def compute_trend_slope(hdf, aid, lookback=4):
    act_hist = hdf[hdf["act_id"] == aid].sort_values("week_number").tail(lookback)
    if len(act_hist) < 2:
        return 0.0
    variances = act_hist["variance_pct"].values
    return float(variances[-1] - variances[0]) / len(variances)

def compute_risk(act, history_df, lookback=4):
    aid      = act["act_id"]
    slope    = compute_trend_slope(history_df, aid, lookback)
    variance = float(act.get("variance_pct", 0) or 0)
    slip     = int(act.get("weeks_slip", 0) or 0)
    critical = bool(act.get("is_critical_path", False))
    score = 0; flags = []
    if slope < -0.015:
        score += 3; flags.append(f"Worsening trend ({abs(slope)*100:.1f}%/wk)")
    if variance < -0.05:
        score += 2; flags.append(f"{abs(variance)*100:.1f}% behind plan")
    if slip >= 2:
        score += 2; flags.append(f"{slip} weeks slip")
    if critical:
        score += 2; flags.append("Critical path")
    if slope < 0 and variance < -0.03:
        score += 1; flags.append("Negative trend + delay")
    return score, flags, slope

def update_flag_status(flag_id, new_status):
    supabase.table("detection_results").update(
        {"status": new_status}).eq("id", flag_id).execute()

# ── MOM docx builder ─────────────────────────────────────────
def build_mom_docx(mom_text: str, project_name: str,
                   meeting_date, week_num: int) -> bytes:
    try:
        from docx import Document as DocxDoc
        from docx.shared import Pt, RGBColor, Inches
        from docx.enum.text import WD_ALIGN_PARAGRAPH

        doc = DocxDoc()
        doc.core_properties.author = "SiteInsight"

        # Page margins
        for section in doc.sections:
            section.top_margin    = Inches(1)
            section.bottom_margin = Inches(1)
            section.left_margin   = Inches(1.2)
            section.right_margin  = Inches(1.2)

        # Header
        h = doc.add_paragraph()
        h.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = h.add_run("MINUTES OF MEETING")
        run.bold = True
        run.font.size = Pt(16)
        run.font.color.rgb = RGBColor(0x1E, 0x27, 0x61)

        sub = doc.add_paragraph()
        sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
        sr = sub.add_run(f"{project_name}  ·  Week {week_num:02d}  ·  {meeting_date}")
        sr.font.size = Pt(11)
        sr.font.color.rgb = RGBColor(0xE8, 0x77, 0x22)

        doc.add_paragraph()

        # Parse and render MOM text
        for line in mom_text.split("\n"):
            line = line.strip()
            if not line:
                doc.add_paragraph()
                continue
            if line.isupper() and len(line) > 3:
                # Section heading
                p = doc.add_paragraph()
                r = p.add_run(line)
                r.bold = True
                r.font.size = Pt(11)
                r.font.color.rgb = RGBColor(0x1E, 0x27, 0x61)
                p.paragraph_format.space_before = Pt(12)
                p.paragraph_format.space_after  = Pt(4)
                # underline with border
                p.paragraph_format.border_bottom = True
            elif line.startswith("- "):
                p = doc.add_paragraph(style="List Bullet")
                r = p.add_run(line[2:])
                r.font.size = Pt(10)
                p.paragraph_format.space_after = Pt(3)
            else:
                p = doc.add_paragraph()
                r = p.add_run(line)
                r.font.size = Pt(10)
                p.paragraph_format.space_after = Pt(3)

        # Footer
        doc.add_paragraph()
        f = doc.add_paragraph()
        f.alignment = WD_ALIGN_PARAGRAPH.CENTER
        fr = f.add_run(f"Generated by SiteInsight  ·  {datetime.today().strftime('%d %b %Y')}")
        fr.font.size = Pt(8)
        fr.font.color.rgb = RGBColor(0x99, 0x99, 0x99)
        fr.italic = True

        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        return buf.getvalue()
    except ImportError:
        return None

# ════════════════════════════════════════════════════════════════
# SIDEBAR
# ════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("## 🏗 SiteInsight")
    st.markdown("*WPR Intelligence Platform*")
    st.divider()

    weeks_data = load_weeks()
    if not weeks_data:
        st.error("No data loaded. Run load_history.py first.")
        st.stop()

    week_options = [w["week_number"] for w in weeks_data]
    live_weeks   = [w["week_number"] for w in weeks_data
                    if w.get("load_type") == "live"]

    selected_week = st.selectbox(
        "Select week",
        options=week_options,
        index=len(week_options) - 1,
        format_func=lambda w:
            f"Week {w:02d} {'🟢' if w in live_weeks else '📁'}"
    )

    st.divider()

    # Navigation selector
    st.markdown("**Navigate to**")
    nav = st.radio(
        "nav",
        ["📋 Intelligence Report",
         "📝 Generate MOM",
         "📊 Generate PPT",
         "📈 Generate Graphs",
         "🔄 Recovery Simulator",
         "⚠️ Early Warning"],
        label_visibility="collapsed",
    )

    st.divider()
    baselines_df = load_baselines()
    if not baselines_df.empty:
        st.markdown("**Baselines**")
        for _, row in baselines_df.iterrows():
            active = "✅ " if row["is_active"] else ""
            st.caption(
                f"{active}**{row['baseline_version']}** · "
                f"ends {row['planned_end_date']}"
            )

    st.divider()
    if st.button("🔄 Refresh"):
        st.cache_data.clear()
        st.rerun()

# ── Load data ─────────────────────────────────────────────────
header     = load_header(selected_week)
activities = load_activities(selected_week)
detections = load_detections(selected_week)
narrative  = load_narrative(selected_week)
dq_flags   = load_dq_flags(selected_week)
history_df = load_all_history()

# Risk engine
risk_list = []
if not activities.empty and not history_df.empty:
    for _, act in activities.iterrows():
        score, flags, slope = compute_risk(act, history_df)
        if score > 0:
            risk_list.append({
                "act_id":      act["act_id"],
                "phase":       act.get("phase", ""),
                "activity":    act["activity"],
                "score":       score,
                "flags":       flags,
                "slope":       slope,
                "variance":    act["variance_pct"],
                "critical":    act.get("is_critical_path", False),
                "responsible": act.get("responsible_person", "—") or "—",
            })
risk_df = pd.DataFrame(risk_list)
if not risk_df.empty:
    risk_df = risk_df.sort_values("score", ascending=False)

project_name = header.get("project_name", "Horizon Commercial Tower")
baseline_ver = header.get("baseline_version", "B1")
contractor   = header.get("contractor", "—")
report_date  = header.get("report_date", "—")

# ── Page header ───────────────────────────────────────────────
st.markdown(
    f"<h1>🏗 {project_name}</h1>",
    unsafe_allow_html=True
)
st.markdown(
    f"<p style='color:#666;margin-top:-12px;'>Week {selected_week:02d} &nbsp;·&nbsp; "
    f"Baseline: <b>{baseline_ver}</b> &nbsp;·&nbsp; "
    f"Contractor: <b>{contractor}</b> &nbsp;·&nbsp; "
    f"Report date: <b>{report_date}</b></p>",
    unsafe_allow_html=True
)

if header.get("baseline_revised_this_week"):
    st.warning(
        f"⚠️ Baseline revised this week · "
        f"{header.get('baseline_revision_note','')}"
    )

# ── KPI summary — always visible at top ──────────────────────
if not activities.empty:
    total_acts      = len(activities)
    behind          = int((activities["variance_pct"] < -0.05).sum())
    critical_behind = 0
    if "is_critical_path" in activities.columns:
        critical_behind = int((
            activities[activities["is_critical_path"] == True]
            ["variance_pct"] < -0.05
        ).sum())
    avg_variance = activities["variance_pct"].mean()
    max_slip     = int(activities["weeks_slip"].max()) \
                   if "weeks_slip" in activities.columns else 0

    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("Total activities",  total_acts)
    k2.metric("Behind plan",       behind,
              delta=f"{behind/total_acts*100:.0f}% of project",
              delta_color="inverse")
    k3.metric("Critical path risk", critical_behind,
              delta="HIGH ALERT" if critical_behind > 0 else "Clear",
              delta_color="inverse" if critical_behind > 0 else "normal")
    k4.metric("Avg variance",      f"{avg_variance*100:.1f}%",
              delta_color="inverse")
    k5.metric("Max weeks slip",    f"{max_slip} wks",
              delta_color="inverse" if max_slip > 0 else "normal")

st.divider()


# ════════════════════════════════════════════════════════════════
# SECTION: INTELLIGENCE REPORT
# ════════════════════════════════════════════════════════════════
if nav == "📋 Intelligence Report":

    # ── Top Risks ─────────────────────────────────────────────
    st.markdown("### 🚨 Top Risks This Week")
    if not risk_df.empty:
        top_risks = risk_df.head(5)
        for _, r in top_risks.iterrows():
            risk_level = "HIGH" if r["score"] >= 7 else \
                         "MEDIUM" if r["score"] >= 4 else "LOW"
            badge_class = f"badge-{risk_level.lower()}"
            location = f"{r['phase']} · {r['activity']}"
            flags_str = " · ".join(r["flags"])
            resp = r["responsible"] if r["responsible"] != "—" else "Not assigned"

            col_badge, col_body = st.columns([1, 8])
            with col_badge:
                st.markdown(
                    f'<span class="{badge_class}">{risk_level}</span>',
                    unsafe_allow_html=True
                )
            with col_body:
                st.markdown(
                    f"**{location}**  \n"
                    f"_{flags_str}_  \n"
                    f"👤 Responsible: **{resp}**"
                )
            st.divider()
    else:
        st.success("No risks identified this week.")

    # ── Executive Summary ─────────────────────────────────────
    st.markdown("### 📋 Executive Summary")
    if narrative:
        exec_sum = narrative.get("executive_summary", "")
        if exec_sum:
            st.info(exec_sum)

        # Trend commentary
        if not history_df.empty:
            wk_trend = (history_df.groupby("week_number")["variance_pct"]
                        .mean().reset_index())
            last3 = wk_trend.tail(3)["variance_pct"].values
            if len(last3) >= 2:
                direction = last3[-1] - last3[0]
                if direction < -0.01:
                    st.error(
                        f"📉 Project variance is **worsening** — dropped "
                        f"{abs(direction)*100:.1f}% over last 3 weeks. "
                        f"Intervention required immediately."
                    )
                elif direction > 0.01:
                    st.success(
                        f"📈 Project variance is **recovering** — improved "
                        f"{direction*100:.1f}% over last 3 weeks."
                    )
                else:
                    st.warning(
                        "⚠️ Project variance is **stable but negative** — "
                        "no recovery visible yet."
                    )
    else:
        st.info("No narrative generated yet. Run pipeline.py for this week.")

    st.divider()

    # ── Key Risks + Recommendations ──────────────────────────
    col_r, col_rec = st.columns(2)
    with col_r:
        st.markdown("### ⚠️ Key Risks")
        if narrative:
            risks = format_narrative_list(narrative.get("key_risks", []))
            for item in risks:
                if item:
                    st.markdown(f"- {item}")
        else:
            st.caption("No data available.")

    with col_rec:
        st.markdown("### ✅ Recommendations")
        if narrative:
            recs = format_narrative_list(narrative.get("recommendations", []))
            for item in recs:
                if item:
                    st.markdown(f"- {item}")
        else:
            st.caption("No data available.")

    st.divider()

    # ── Detection Flags — compact table ──────────────────────
    st.markdown(f"### 🔍 Detection Flags ({len(detections)} this week)")

    if selected_week > 1:
        prev_det = supabase.table("detection_results").select(
            "status").eq("week_number", selected_week - 1).execute()
        if prev_det.data:
            prev_df     = pd.DataFrame(prev_det.data)
            resolved    = (prev_df["status"] == "Resolved").sum()
            in_progress = (prev_df["status"] == "In Progress").sum()
            still_open  = (prev_df["status"] == "Open").sum()
            st.caption(
                f"From Week {selected_week-1:02d}: "
                f"✓ {resolved} resolved · "
                f"⏳ {in_progress} in progress · "
                f"🔴 {still_open} still open"
            )

    if detections.empty:
        st.success("No flags raised this week.")
    else:
        rule_names = {
            "R1": "Schedule Slip", "R2": "Scope Creep",
            "R3": "Stale Issue",   "R4": "Critical Path at Risk"
        }
        sev_color = {"high": "🔴", "medium": "🟡", "low": "🟢"}
        status_options = ["Open", "In Progress", "Resolved"]

        # Build display dataframe
        rows = []
        for _, h in detections.iterrows():
            act_id = h["act_id"]
            resp   = "—"
            if not activities.empty and "responsible_person" in activities.columns:
                match = activities[activities["act_id"] == act_id]
                if not match.empty:
                    rp = match.iloc[0].get("responsible_person")
                    if rp and str(rp) not in ["nan","None",""]:
                        resp = str(rp)

            # Get phase + activity for location
            location = act_id
            if not activities.empty:
                match = activities[activities["act_id"] == act_id]
                if not match.empty:
                    row_act = match.iloc[0]
                    location = f"{row_act.get('phase','')} · {row_act.get('activity','')}"

            rows.append({
                "Severity": sev_color.get(h.get("severity","low"),"⚪"),
                "Rule":     f"{h['rule_id']} · {rule_names.get(h['rule_id'],'')}",
                "Location": location,
                "Detail":   h.get("detail",""),
                "Responsible": resp,
                "Status":   h.get("status","Open") or "Open",
                "_id":      h.get("id"),
                "_rule":    h.get("rule_id"),
                "_act":     act_id,
            })

        flags_display = pd.DataFrame(rows)

        # Show compact table
        st.dataframe(
            flags_display[["Severity","Rule","Location","Responsible","Status"]],
            use_container_width=True,
            hide_index=True,
        )

        # Status update — below the table
        with st.expander("Update flag status"):
            for row in rows:
                col_loc, col_status, col_btn = st.columns([3, 1, 1])
                with col_loc:
                    st.caption(f"{row['Severity']} {row['Location']}")
                with col_status:
                    new_s = st.selectbox(
                        "s",
                        status_options,
                        index=status_options.index(row["Status"]),
                        key=f"s_{row['_id']}_{row['_rule']}_{row['_act']}",
                        label_visibility="collapsed",
                    )
                with col_btn:
                    if new_s != row["Status"] and row["_id"]:
                        if st.button("Save", key=f"b_{row['_id']}"):
                            update_flag_status(row["_id"], new_s)
                            st.cache_data.clear()
                            st.rerun()

    st.divider()

    # ── Activity Table ────────────────────────────────────────
    st.markdown("### 📊 Activity Progress")
    if not activities.empty:
        phases    = ["All"] + sorted(
            activities["phase"].dropna().unique().tolist())
        sel_phase = st.selectbox("Filter by phase", phases)
        disp      = activities.copy()
        if sel_phase != "All":
            disp = disp[disp["phase"] == sel_phase]

        show_cols = [c for c in [
            "act_id","phase","activity","unit",
            "cum_actual_pct","cum_planned_pct","variance_pct",
            "weeks_slip","delay_reason","responsible_person",
            "is_critical_path","remarks"
        ] if c in disp.columns]
        disp = disp[show_cols].copy()
        for pct_col in ["cum_actual_pct","cum_planned_pct","variance_pct"]:
            if pct_col in disp.columns:
                disp[pct_col] = (disp[pct_col]*100).round(1).astype(str)+"%"
        if "is_critical_path" in disp.columns:
            disp["is_critical_path"] = disp["is_critical_path"].map(
                {True:"✅", False:"—"})
        disp.columns = [c.replace("_"," ").title() for c in disp.columns]
        st.dataframe(disp, use_container_width=True, hide_index=True)

    st.divider()

    # ── Site Photos — 4 curated images ───────────────────────
    st.markdown("### 📸 Site Photos")
    st.caption(
        "Representative site photos by phase. "
        "In production, actual photos are uploaded by the site engineer."
    )
    p1, p2, p3, p4 = st.columns(4)
    cols = [p1, p2, p3, p4]
    for i, photo in enumerate(SITE_PHOTOS):
        with cols[i]:
            st.image(
                photo["url"],
                caption=photo["caption"],
                use_container_width=True,
            )


# ════════════════════════════════════════════════════════════════
# SECTION: GENERATE MOM
# ════════════════════════════════════════════════════════════════
elif nav == "📝 Generate MOM":
    st.markdown("### 📝 Minutes of Meeting Generator")
    st.caption(
        "Generates a professional MOM in plain language. "
        "Download as a formatted Word document."
    )

    with st.form("mom_form"):
        mc1, mc2 = st.columns(2)
        meeting_date        = mc1.date_input("Meeting date", value=date.today())
        meeting_chairperson = mc2.text_input("Chaired by", value="Project Manager")

        attendees = st.text_area(
            "Attendees (one per line)",
            value="Project Manager\nPlanning Engineer\nSite Engineer\nContractor Representative",
            height=80,
        )

        st.markdown("**Sections to include**")
        sc1, sc2, sc3 = st.columns(3)
        inc_delays   = sc1.checkbox("Delayed activities", value=True)
        inc_critical = sc2.checkbox("Critical path risks", value=True)
        inc_stale    = sc3.checkbox("Recurring issues",   value=True)
        sc4, sc5, sc6 = st.columns(3)
        inc_actions  = sc4.checkbox("Action items",       value=True)
        inc_baseline = sc5.checkbox("Baseline status",    value=True)
        inc_next     = sc6.checkbox("Next week targets",  value=True)

        gen_mom = st.form_submit_button("Generate MOM", type="primary")

    if gen_mom:
        if activities.empty:
            st.warning("No activity data for this week.")
        else:
            with st.spinner("Generating MOM via Groq..."):
                behind_acts = activities[activities["variance_pct"] < -0.05]
                det_r3 = detections[
                    detections["rule_id"] == "R3"
                ] if not detections.empty else pd.DataFrame()

                sections = (
                    (["delayed_activities"] if inc_delays   else []) +
                    (["critical_path"]      if inc_critical else []) +
                    (["stale_issues"]       if inc_stale    else []) +
                    (["action_items"]       if inc_actions  else []) +
                    (["baseline_status"]    if inc_baseline else []) +
                    (["next_week_targets"]  if inc_next     else [])
                )

                # Build activity context — plain language
                behind_str = "None"
                if not behind_acts.empty:
                    rows_list = []
                    for _, r in behind_acts.iterrows():
                        rp = r.get("responsible_person","—") or "—"
                        rows_list.append(
                            f"{r['phase']} · {r['activity']} · "
                            f"Delay reason: {r['delay_reason']} · "
                            f"Responsible: {rp}"
                        )
                    behind_str = "\n".join(rows_list)

                prompt = f"""You are a senior planning engineer writing formal Minutes of Meeting for a construction project review.

PROJECT: {project_name}
WEEK: {selected_week} | DATE: {meeting_date} | CHAIRED BY: {meeting_chairperson}
BASELINE: {baseline_ver} | CONTRACTOR: {contractor}

DELAYED ACTIVITIES (for context only — do NOT repeat raw numbers in the MOM):
{behind_str}

RECURRING ISSUES:
{det_r3['detail'].tolist() if not det_r3.empty else 'None'}

SUMMARY: {narrative.get('executive_summary','') if narrative else ''}

Generate a professional MOM with sections: {', '.join(sections)}

STRICT WRITING RULES:
- Write in plain construction project language — no variance percentages, no slip numbers
- Describe issues in plain words e.g. "Basement 1 Shuttering is significantly behind schedule due to monsoon rains"
- Every action item must state: what needs to be done, who is responsible (use name), deadline is "by next site review meeting"
- For recurring issues: state the issue, who is responsible, and consequence if unresolved
- MOM should read like something a site engineer would actually write — clear, direct, professional
- Use ALL CAPS for section headings only
- No bullet points for the intro paragraphs — use them only for action items

Start with:
MINUTES OF MEETING
Project: {project_name}
Date: {meeting_date}
Week: {selected_week}
Chaired by: {meeting_chairperson}
Attendees: {', '.join([a.strip() for a in attendees.split(chr(10)) if a.strip()])}
"""
                mom_text = call_groq(prompt, max_tokens=2000)

            st.divider()
            st.markdown("**Generated MOM**")
            st.text_area("", value=mom_text, height=400,
                         label_visibility="collapsed")

            # Download as .docx
            docx_bytes = build_mom_docx(
                mom_text, project_name, meeting_date, selected_week)

            dl1, dl2 = st.columns(2)
            with dl1:
                if docx_bytes:
                    st.download_button(
                        "⬇️ Download MOM as Word (.docx)",
                        docx_bytes,
                        f"MOM_{project_name.replace(' ','_')}_Wk{selected_week:02d}.docx",
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    )
                else:
                    st.info("Install python-docx for Word download: pip install python-docx")
            with dl2:
                st.download_button(
                    "⬇️ Download MOM as .txt",
                    mom_text,
                    f"MOM_SiteInsight_Wk{selected_week:02d}.txt",
                    "text/plain",
                )

            # Email notifications
            st.divider()
            st.markdown("**📧 Send Email Notifications**")
            st.caption("Each responsible person receives only their own action items.")

            if not EMAIL_SENDER:
                st.warning("EMAIL_SENDER not set in .env — add it to enable notifications.")
            else:
                person_actions = {}
                if not behind_acts.empty and "responsible_person" in behind_acts.columns:
                    for _, r in behind_acts.iterrows():
                        rp = r.get("responsible_person")
                        if rp and str(rp) not in ["nan","None",""]:
                            person = str(rp)
                            if person not in person_actions:
                                person_actions[person] = []
                            person_actions[person].append(
                                f"• {r['phase']} · {r['activity']} — "
                                f"{r['delay_reason']}. "
                                f"Action required by next site review meeting."
                            )

                if not person_actions:
                    st.info("No responsible persons found. Add names to Wk11+ Excel files.")
                else:
                    email_map = {}
                    for person in person_actions:
                        ec1, ec2 = st.columns([1,2])
                        ec1.write(f"👤 **{person}**")
                        email_map[person] = ec2.text_input(
                            f"Email for {person}",
                            placeholder="name@company.com",
                            key=f"email_{person}",
                            label_visibility="collapsed",
                        )

                    if st.button("📧 Send notifications", type="primary"):
                        sent = 0
                        for person, actions in person_actions.items():
                            recipient = email_map.get(person,"").strip()
                            if not recipient or "@" not in recipient:
                                st.warning(f"No valid email for {person} — skipped")
                                continue
                            body = f"""Dear {person},

Please find below your action items from the Week {selected_week:02d} site review meeting for {project_name}.

YOUR ACTION ITEMS:
{"".join([chr(10)+a for a in actions])}

All actions are to be completed by the next site review meeting.

Meeting date: {meeting_date}
Chaired by: {meeting_chairperson}

Regards,
SiteInsight · {project_name}
"""
                            subject = (
                                f"Action Items — Week {selected_week:02d} "
                                f"Site Review · {project_name}"
                            )
                            if send_email(recipient, subject, body):
                                st.success(f"✅ Sent to {person} ({recipient})")
                                sent += 1
                        if sent > 0:
                            st.success(f"Notifications sent to {sent} person(s).")


# ════════════════════════════════════════════════════════════════
# SECTION: GENERATE PPT
# ════════════════════════════════════════════════════════════════
elif nav == "📊 Generate PPT":
    st.markdown("### 📊 PPT Composer")
    st.caption(
        "Select exactly what goes into your presentation. "
        "Add custom notes. Set slide order. Download and edit in PowerPoint."
    )

    if activities.empty:
        st.info("No activity data available for this week.")
    else:
        st.divider()

        # ── STEP 1: CONTENT SELECTOR ─────────────────────────
        st.markdown("#### Step 1 — Select content for your slides")

        selected_slides_data = {}

        # Cover (always on)
        st.markdown("**Cover Slide**")
        inc_cover = st.checkbox("Include cover slide", value=True, key="ppt_cover")
        if inc_cover:
            selected_slides_data["cover"] = {"note": ""}

        st.divider()

        # Executive Summary
        st.markdown("**Executive Summary**")
        inc_exec = st.checkbox("Include executive summary", value=True, key="ppt_exec")
        if inc_exec:
            exec_text = narrative.get("executive_summary","") if narrative else ""
            exec_note = st.text_area(
                "Edit or add to executive summary",
                value=exec_text,
                height=80,
                key="ppt_exec_note",
            )
            selected_slides_data["executive_summary"] = {"note": exec_note}

        st.divider()

        # Top Risks
        st.markdown("**Top Risks**")
        inc_risks = st.checkbox("Include top risks slide", value=True, key="ppt_risks_main")
        selected_risks = []
        if inc_risks and not risk_df.empty:
            st.caption("Select which risks to include:")
            for idx, (_, r) in enumerate(risk_df.head(8).iterrows()):
                risk_level = "HIGH" if r["score"] >= 7 else "MEDIUM" if r["score"] >= 4 else "LOW"
                rc1, rc2 = st.columns([3, 4])
                with rc1:
                    tick = st.checkbox(
                        f"{risk_level} · {r['phase']} · {r['activity']}",
                        value=(risk_level == "HIGH"),
                        key=f"ppt_risk_{idx}",
                    )
                with rc2:
                    note = st.text_input(
                        "Note",
                        placeholder="Add context e.g. Recovery plan submitted",
                        key=f"ppt_risk_note_{idx}",
                        label_visibility="collapsed",
                    )
                if tick:
                    selected_risks.append({
                        "location": f"{r['phase']} · {r['activity']}",
                        "level":    risk_level,
                        "flags":    " · ".join(r["flags"]),
                        "responsible": r["responsible"],
                        "note":     note,
                    })
            if selected_risks:
                selected_slides_data["top_risks"] = {"items": selected_risks}

        st.divider()

        # Phase Status
        st.markdown("**Phase-wise Status**")
        inc_phase = st.checkbox("Include phase status slide", value=True, key="ppt_phase_main")
        selected_phases = []
        if inc_phase and not activities.empty:
            phases_avail = sorted(activities["phase"].dropna().unique().tolist())
            st.caption("Select which phases to include:")
            for phase in phases_avail:
                ph_grp = activities[activities["phase"] == phase]
                avg_actual  = ph_grp["cum_actual_pct"].mean() * 100
                avg_planned = ph_grp["cum_planned_pct"].mean() * 100
                behind_n    = (ph_grp["variance_pct"] < -0.05).sum()
                pc1, pc2 = st.columns([3, 4])
                with pc1:
                    ph_tick = st.checkbox(
                        f"{phase} — Actual {avg_actual:.0f}% vs Planned {avg_planned:.0f}%",
                        value=True,
                        key=f"ppt_phase_{phase}",
                    )
                with pc2:
                    ph_note = st.text_input(
                        "Note",
                        placeholder="Add phase-specific note",
                        key=f"ppt_phase_note_{phase}",
                        label_visibility="collapsed",
                    )
                if ph_tick:
                    selected_phases.append({
                        "phase":   phase,
                        "actual":  avg_actual,
                        "planned": avg_planned,
                        "behind":  int(behind_n),
                        "total":   len(ph_grp),
                        "note":    ph_note,
                    })
            if selected_phases:
                selected_slides_data["phase_status"] = {"items": selected_phases}

        st.divider()

        # Delayed Activities
        st.markdown("**Delayed Activities**")
        inc_delayed = st.checkbox("Include delayed activities slide", value=True, key="ppt_delayed_main")
        selected_delayed = []
        if inc_delayed and not activities.empty:
            behind_acts_ppt = activities[activities["variance_pct"] < -0.05]
            if not behind_acts_ppt.empty:
                st.caption("Select which activities to highlight:")
                for idx, (_, r) in enumerate(behind_acts_ppt.iterrows()):
                    rp = r.get("responsible_person","—") or "—"
                    dc1, dc2 = st.columns([3, 4])
                    with dc1:
                        d_tick = st.checkbox(
                            f"{r['phase']} · {r['activity']} ({r['variance_pct']*100:.1f}%)",
                            value=True,
                            key=f"ppt_del_{idx}",
                        )
                    with dc2:
                        d_note = st.text_input(
                            "Note",
                            placeholder="e.g. Material ordered, delivery next week",
                            key=f"ppt_del_note_{idx}",
                            label_visibility="collapsed",
                        )
                    if d_tick:
                        selected_delayed.append({
                            "location":    f"{r['phase']} · {r['activity']}",
                            "delay_reason": r.get("delay_reason","—"),
                            "responsible": rp,
                            "note":        d_note,
                        })
                if selected_delayed:
                    selected_slides_data["delayed_activities"] = {"items": selected_delayed}

        st.divider()

        # Graphs
        st.markdown("**Graphs**")
        inc_graphs = st.checkbox("Include graphs slide(s)", value=True, key="ppt_graphs_main")
        selected_graphs_ppt = []
        if inc_graphs:
            st.caption("Select graphs to embed — all selected graphs placed in 1-2 slides:")
            gg1, gg2, gg3 = st.columns(3)
            with gg1:
                if st.checkbox("S-Curve",         value=True,  key="ppt_g_scurve"):   selected_graphs_ppt.append("scurve")
                if st.checkbox("Variance trend",   value=True,  key="ppt_g_var"):      selected_graphs_ppt.append("variance")
            with gg2:
                if st.checkbox("Phase completion", value=True,  key="ppt_g_phase"):    selected_graphs_ppt.append("phase")
                if st.checkbox("Weeks slip",       value=False, key="ppt_g_slip"):     selected_graphs_ppt.append("slip")
            with gg3:
                if st.checkbox("Delay reasons",    value=False, key="ppt_g_delay"):    selected_graphs_ppt.append("delay")
                if st.checkbox("Critical path",    value=False, key="ppt_g_critical"): selected_graphs_ppt.append("critical")
            if selected_graphs_ppt:
                selected_slides_data["graphs"] = {"items": selected_graphs_ppt}

        st.divider()

        # Next Week Targets
        st.markdown("**Next Week Targets**")
        inc_next = st.checkbox("Include next week targets slide", value=True, key="ppt_next_main")
        if inc_next:
            next_note = st.text_area(
                "Next week targets / additional notes",
                placeholder="e.g. Target: recover 5% on Phase I shuttering. Mr. D to submit recovery plan.",
                height=80,
                key="ppt_next_note",
            )
            selected_slides_data["next_week"] = {"note": next_note}

        st.divider()

        # ── STEP 2: SLIDE ORDER ───────────────────────────────
        st.markdown("#### Step 2 — Set slide order")
        st.caption("Enter a number for each section to set the order (1 = first slide after cover).")

        slide_order_map = {}
        order_keys = [k for k in selected_slides_data.keys() if k != "cover"]
        slide_labels = {
            "executive_summary":  "Executive Summary",
            "top_risks":          "Top Risks",
            "phase_status":       "Phase-wise Status",
            "delayed_activities": "Delayed Activities",
            "graphs":             "Graphs",
            "next_week":          "Next Week Targets",
        }

        if order_keys:
            ord_cols = st.columns(min(len(order_keys), 3))
            for i, key in enumerate(order_keys):
                with ord_cols[i % 3]:
                    slide_order_map[key] = st.number_input(
                        slide_labels.get(key, key),
                        min_value=1,
                        max_value=10,
                        value=i + 1,
                        key=f"order_{key}",
                    )

            ordered_keys = sorted(
                order_keys,
                key=lambda k: slide_order_map.get(k, 99)
            )
        else:
            ordered_keys = []

        st.divider()

        # ── STEP 3: DETAILS ───────────────────────────────────
        st.markdown("#### Step 3 — Presentation details")
        det1, det2 = st.columns(2)
        ppt_by       = det1.text_input("Prepared by", value="Planning Team", key="ppt_by")
        ppt_audience = det2.selectbox(
            "Audience",
            ["Project Director","Client","Internal Team","Site Review Meeting"],
            key="ppt_audience",
        )

        st.divider()

        # ── GENERATE ──────────────────────────────────────────
        if st.button("🎯 Generate Presentation", type="primary"):
            if not selected_slides_data:
                st.warning("Please select at least one section.")
            else:
                with st.spinner("Generating slide content via Groq..."):

                    # Build prompt from selected content only
                    content_block = ""
                    for key in (["cover"] + ordered_keys):
                        data = selected_slides_data.get(key, {})
                        if key == "cover":
                            content_block += f"COVER: {project_name} · Week {selected_week} · {ppt_audience}\n"
                        elif key == "executive_summary":
                            content_block += f"EXECUTIVE SUMMARY: {data.get('note','')}\n"
                        elif key == "top_risks":
                            content_block += "TOP RISKS:\n"
                            for item in data.get("items",[]):
                                content_block += (
                                    f"  - {item['level']} · {item['location']} · "
                                    f"{item['flags']} · Responsible: {item['responsible']}"
                                )
                                if item["note"]:
                                    content_block += f" · Note: {item['note']}"
                                content_block += "\n"
                        elif key == "phase_status":
                            content_block += "PHASE STATUS:\n"
                            for item in data.get("items",[]):
                                content_block += (
                                    f"  - {item['phase']}: Actual {item['actual']:.0f}% vs "
                                    f"Planned {item['planned']:.0f}%, "
                                    f"{item['behind']}/{item['total']} behind"
                                )
                                if item["note"]:
                                    content_block += f" · Note: {item['note']}"
                                content_block += "\n"
                        elif key == "delayed_activities":
                            content_block += "DELAYED ACTIVITIES:\n"
                            for item in data.get("items",[]):
                                content_block += (
                                    f"  - {item['location']} · "
                                    f"Reason: {item['delay_reason']} · "
                                    f"Responsible: {item['responsible']}"
                                )
                                if item["note"]:
                                    content_block += f" · Note: {item['note']}"
                                content_block += "\n"
                        elif key == "next_week":
                            content_block += f"NEXT WEEK TARGETS: {data.get('note','')}\n"

                    prompt = f"""You are preparing a PowerPoint presentation for {ppt_audience}.
PROJECT: {project_name} | WEEK: {selected_week} | BASELINE: {baseline_ver}
PREPARED BY: {ppt_by}

SELECTED CONTENT (use ONLY this — do not add anything not listed here):
{content_block}

Generate slide content for these sections in this order: {", ".join(["cover"] + ordered_keys)}
Skip "graphs" section — that is handled separately.

Return JSON only — no markdown fences:
{{
  "cover":               {{"title": "...", "subtitle": "...", "details": ["..."]}},
  "executive_summary":   {{"title": "...", "bullets": ["..."]}},
  "top_risks":           {{"title": "...", "bullets": ["..."]}},
  "phase_status":        {{"title": "...", "bullets": ["..."]}},
  "delayed_activities":  {{"title": "...", "bullets": ["..."]}},
  "next_week":           {{"title": "...", "bullets": ["..."]}}
}}

RULES:
- Use ONLY the content provided above — no generic filler
- 3-5 bullets per slide, max 20 words each
- Each bullet must be specific — name the activity, phase, or person
- If engineer added a custom note for an item, include it in the bullet
- Audience is {ppt_audience} — write accordingly
- Return only the sections that are in the selected content
"""
                    raw = call_groq(prompt, max_tokens=1500)
                    try:
                        if "```" in raw:
                            raw = raw.split("```")[1]
                            if raw.startswith("json"):
                                raw = raw[4:]
                        slide_content = json.loads(raw.strip())
                    except Exception:
                        slide_content = {}

                # ── Generate graphs as images ─────────────────
                graph_images = {}
                if "graphs" in selected_slides_data:
                    g_list = selected_slides_data["graphs"]["items"]

                    if "scurve" in g_list and not history_df.empty:
                        scurve = history_df.groupby("week_number").agg(
                            actual=("cum_actual_pct","mean"),
                            planned=("cum_planned_pct","mean")
                        ).reset_index()
                        scurve[["actual","planned"]] *= 100
                        fig, ax = plt.subplots(figsize=(10,4))
                        ax.plot(scurve["week_number"], scurve["planned"],
                                color=NAVY, linewidth=2, label="Planned %",
                                linestyle="--", marker="o", markersize=3)
                        ax.plot(scurve["week_number"], scurve["actual"],
                                color=AMBER, linewidth=2, label="Actual %",
                                marker="o", markersize=3)
                        ax.fill_between(scurve["week_number"],
                                        scurve["planned"], scurve["actual"],
                                        where=scurve["actual"] < scurve["planned"],
                                        alpha=0.12, color=RED)
                        ax.axvline(selected_week, color=AMBER, linewidth=1.5,
                                   linestyle=":", alpha=0.8)
                        ax.set_xlabel("Week"); ax.set_ylabel("Progress (%)")
                        ax.set_title("S-Curve: Planned vs Actual", fontweight="bold")
                        ax.legend(); ax.grid(alpha=0.3)
                        ax.set_facecolor("#FAFAFA"); fig.patch.set_facecolor("white")
                        plt.tight_layout()
                        buf = io.BytesIO()
                        fig.savefig(buf, format="png", dpi=120, bbox_inches="tight")
                        buf.seek(0)
                        graph_images["scurve"] = buf.getvalue()
                        plt.close(fig)

                    if "variance" in g_list and not history_df.empty:
                        trend = (history_df.groupby("week_number")["variance_pct"]
                                 .mean().reset_index())
                        trend["variance_pct"] *= 100
                        fig, ax = plt.subplots(figsize=(10,3.5))
                        colors_v = [RED if v < 0 else GREEN for v in trend["variance_pct"]]
                        ax.bar(trend["week_number"], trend["variance_pct"],
                               color=colors_v, width=0.6, zorder=3)
                        ax.axhline(0, color=GRAY, linewidth=0.8, linestyle="--")
                        ax.set_xlabel("Week"); ax.set_ylabel("Avg Variance (%)")
                        ax.set_title("Variance Trend — All Weeks", fontweight="bold")
                        ax.grid(axis="y", alpha=0.3)
                        ax.set_facecolor("#FAFAFA"); fig.patch.set_facecolor("white")
                        plt.tight_layout()
                        buf = io.BytesIO()
                        fig.savefig(buf, format="png", dpi=120, bbox_inches="tight")
                        buf.seek(0)
                        graph_images["variance"] = buf.getvalue()
                        plt.close(fig)

                    if "phase" in g_list and not activities.empty:
                        pd_d = activities.groupby("phase").agg(
                            actual=("cum_actual_pct","mean"),
                            planned=("cum_planned_pct","mean")
                        ).reset_index()
                        pd_d[["actual","planned"]] *= 100
                        fig, ax = plt.subplots(figsize=(9,4))
                        xp = range(len(pd_d)); wp = 0.35
                        ax.bar([i-wp/2 for i in xp], pd_d["planned"],
                               wp, label="Planned %", color=NAVY, alpha=0.85)
                        ax.bar([i+wp/2 for i in xp], pd_d["actual"],
                               wp, label="Actual %", color=AMBER, alpha=0.85)
                        ax.set_xticks(list(xp)); ax.set_xticklabels(pd_d["phase"])
                        ax.set_ylabel("Completion (%)")
                        ax.set_title("Phase-wise Status", fontweight="bold")
                        ax.legend(); ax.grid(axis="y", alpha=0.3)
                        ax.set_facecolor("#FAFAFA"); fig.patch.set_facecolor("white")
                        plt.tight_layout()
                        buf = io.BytesIO()
                        fig.savefig(buf, format="png", dpi=120, bbox_inches="tight")
                        buf.seek(0)
                        graph_images["phase"] = buf.getvalue()
                        plt.close(fig)

                    if "slip" in g_list and not activities.empty:
                        slip_d = activities[activities["weeks_slip"] > 0].sort_values(
                            "weeks_slip", ascending=True)
                        if not slip_d.empty:
                            labels_s = [f"{r.get('phase','')} · {r.get('activity','')}"
                                        for _, r in slip_d.iterrows()]
                            fig, ax = plt.subplots(figsize=(9, max(3.5, len(slip_d)*0.45)))
                            colors_s = [RED if row.get("is_critical_path") else AMBER
                                        for _, row in slip_d.iterrows()]
                            ax.barh(labels_s, slip_d["weeks_slip"],
                                    color=colors_s, height=0.55)
                            ax.set_xlabel("Weeks Slip")
                            ax.set_title("Weeks Slip by Activity", fontweight="bold")
                            ax.grid(axis="x", alpha=0.3)
                            ax.set_facecolor("#FAFAFA"); fig.patch.set_facecolor("white")
                            plt.tight_layout()
                            buf = io.BytesIO()
                            fig.savefig(buf, format="png", dpi=120, bbox_inches="tight")
                            buf.seek(0)
                            graph_images["slip"] = buf.getvalue()
                            plt.close(fig)

                    if "delay" in g_list and not activities.empty:
                        delayed_d = activities[
                            (activities["variance_pct"] < 0) &
                            (activities["delay_reason"].notna()) &
                            (activities["delay_reason"] != "No Delay")
                        ]
                        if not delayed_d.empty:
                            counts_d = delayed_d["delay_reason"].value_counts()
                            fig, ax = plt.subplots(figsize=(7,4))
                            palette_d = [NAVY,AMBER,RED,AMBER2,GREEN,"#7B61FF"]
                            ax.pie(counts_d.values, labels=counts_d.index,
                                   autopct="%1.0f%%",
                                   colors=palette_d[:len(counts_d)],
                                   startangle=140, pctdistance=0.82)
                            ax.set_title("Delay Reasons", fontweight="bold")
                            fig.patch.set_facecolor("white"); plt.tight_layout()
                            buf = io.BytesIO()
                            fig.savefig(buf, format="png", dpi=120, bbox_inches="tight")
                            buf.seek(0)
                            graph_images["delay"] = buf.getvalue()
                            plt.close(fig)

                    if "critical" in g_list and not activities.empty and "is_critical_path" in activities.columns:
                        cp_d = activities[activities["is_critical_path"] == True].copy()
                        if not cp_d.empty:
                            cp_d = cp_d.sort_values("variance_pct")
                            labels_c = [f"{r.get('phase','')} · {r.get('activity','')}"
                                        for _, r in cp_d.iterrows()]
                            fig, ax = plt.subplots(figsize=(9, max(3.5, len(cp_d)*0.5)))
                            xc = range(len(cp_d))
                            ax.barh([i+0.2 for i in xc], cp_d["cum_planned_pct"]*100,
                                    height=0.35, label="Planned %", color=NAVY, alpha=0.8)
                            ax.barh([i-0.2 for i in xc], cp_d["cum_actual_pct"]*100,
                                    height=0.35, label="Actual %", color=AMBER, alpha=0.8)
                            ax.set_yticks(list(xc)); ax.set_yticklabels(labels_c, fontsize=7)
                            ax.set_xlabel("Completion (%)")
                            ax.set_title("Critical Path: Planned vs Actual", fontweight="bold")
                            ax.legend(); ax.grid(axis="x", alpha=0.3)
                            ax.set_facecolor("#FAFAFA"); fig.patch.set_facecolor("white")
                            plt.tight_layout()
                            buf = io.BytesIO()
                            fig.savefig(buf, format="png", dpi=120, bbox_inches="tight")
                            buf.seek(0)
                            graph_images["critical"] = buf.getvalue()
                            plt.close(fig)

                # ── Build PPTX ────────────────────────────────
                try:
                    from pptx import Presentation as PPTXPres
                    from pptx.util import Inches, Pt, Emu
                    from pptx.dml.color import RGBColor as RGB

                    prs              = PPTXPres()
                    prs.slide_width  = Inches(13.33)
                    prs.slide_height = Inches(7.5)
                    blank            = prs.slide_layouts[6]

                    C_NAVY  = RGB(0x1E, 0x27, 0x61)
                    C_WHITE = RGB(0xFF, 0xFF, 0xFF)
                    C_AMBER = RGB(0xE8, 0x77, 0x22)
                    C_GRAY  = RGB(0x55, 0x55, 0x55)

                    def make_slide(title_text, bullets,
                                   is_cover=False, cover_data=None):
                        sl  = prs.slides.add_slide(blank)
                        # Background
                        bg  = sl.shapes.add_shape(
                            1, 0, 0, prs.slide_width, prs.slide_height)
                        bg.fill.solid()
                        bg.fill.fore_color.rgb = C_NAVY if is_cover else C_WHITE
                        bg.line.fill.background()
                        # Amber top bar
                        bar = sl.shapes.add_shape(
                            1, 0, 0, prs.slide_width, Inches(0.08))
                        bar.fill.solid()
                        bar.fill.fore_color.rgb = C_AMBER
                        bar.line.fill.background()
                        # Title
                        tb = sl.shapes.add_textbox(
                            Inches(0.5), Inches(0.2),
                            Inches(12), Inches(1.2))
                        tp = tb.text_frame.paragraphs[0]
                        tp.text = title_text
                        tp.font.size     = Pt(32 if is_cover else 24)
                        tp.font.bold     = True
                        tp.font.color.rgb = C_WHITE if is_cover else C_NAVY

                        if is_cover and cover_data:
                            # Subtitle
                            sb = sl.shapes.add_textbox(
                                Inches(0.5), Inches(1.8),
                                Inches(12), Inches(1))
                            sp2 = sb.text_frame.paragraphs[0]
                            sp2.text = cover_data.get("subtitle","")
                            sp2.font.size     = Pt(18)
                            sp2.font.color.rgb = C_AMBER
                            # Details
                            db  = sl.shapes.add_textbox(
                                Inches(0.5), Inches(3.0),
                                Inches(12), Inches(3))
                            dtf = db.text_frame
                            dtf.word_wrap = True
                            for i2, line in enumerate(
                                cover_data.get("details",[])
                            ):
                                dp2 = (dtf.paragraphs[0] if i2 == 0
                                       else dtf.add_paragraph())
                                dp2.text = line
                                dp2.font.size     = Pt(14)
                                dp2.font.color.rgb = C_WHITE
                        else:
                            # Bullets
                            bb  = sl.shapes.add_textbox(
                                Inches(0.5), Inches(1.5),
                                Inches(12), Inches(5.5))
                            btf = bb.text_frame
                            btf.word_wrap = True
                            for i2, bullet in enumerate(bullets):
                                bp2 = (btf.paragraphs[0] if i2 == 0
                                       else btf.add_paragraph())
                                bp2.text = f"▸  {bullet}"
                                bp2.font.size     = Pt(15)
                                bp2.font.color.rgb = C_GRAY
                                bp2.space_after   = Pt(10)
                        # Footer
                        ft = sl.shapes.add_textbox(
                            Inches(0.5), Inches(7.1),
                            Inches(12), Inches(0.3))
                        fp = ft.text_frame.paragraphs[0]
                        fp.text = (
                            f"SiteInsight · {project_name} · "
                            f"Week {selected_week:02d} · CONFIDENTIAL"
                        )
                        fp.font.size     = Pt(8)
                        fp.font.color.rgb = RGB(0x99, 0x99, 0x99)
                        return sl

                    # Cover slide
                    if "cover" in selected_slides_data:
                        cdata = slide_content.get("cover", {})
                        make_slide(
                            cdata.get("title", project_name),
                            [],
                            is_cover=True,
                            cover_data={
                                "subtitle": cdata.get("subtitle",
                                    f"Week {selected_week:02d} WPR Intelligence Report"),
                                "details": cdata.get("details", [
                                    f"Baseline: {baseline_ver}",
                                    f"Contractor: {contractor}",
                                    f"Prepared by: {ppt_by}",
                                    f"Audience: {ppt_audience}",
                                ]),
                            }
                        )

                    # Content slides in user-defined order
                    slide_key_map = {
                        "executive_summary":  "executive_summary",
                        "top_risks":          "top_risks",
                        "phase_status":       "phase_status",
                        "delayed_activities": "delayed_activities",
                        "next_week":          "next_week",
                    }
                    for key in ordered_keys:
                        if key == "graphs":
                            continue
                        if key not in selected_slides_data:
                            continue
                        sc_key   = slide_key_map.get(key, key)
                        sc_data  = slide_content.get(sc_key, {})
                        bullets  = sc_data.get("bullets", [])
                        title    = sc_data.get("title", key.replace("_"," ").title())
                        make_slide(title, bullets)

                    # Graphs slide(s) — position as per order
                    if graph_images:
                        graph_keys = list(graph_images.keys())
                        # Split into groups of 2 per slide
                        for chunk_start in range(0, len(graph_keys), 2):
                            chunk = graph_keys[chunk_start:chunk_start+2]
                            sl    = prs.slides.add_slide(blank)
                            # Background
                            bg2 = sl.shapes.add_shape(
                                1, 0, 0, prs.slide_width, prs.slide_height)
                            bg2.fill.solid()
                            bg2.fill.fore_color.rgb = C_WHITE
                            bg2.line.fill.background()
                            # Amber bar
                            bar2 = sl.shapes.add_shape(
                                1, 0, 0, prs.slide_width, Inches(0.08))
                            bar2.fill.solid()
                            bar2.fill.fore_color.rgb = C_AMBER
                            bar2.line.fill.background()
                            # Title
                            tb2 = sl.shapes.add_textbox(
                                Inches(0.5), Inches(0.1),
                                Inches(12), Inches(0.6))
                            tp2 = tb2.text_frame.paragraphs[0]
                            tp2.text = f"Project Graphs — Week {selected_week:02d}"
                            tp2.font.size     = Pt(20)
                            tp2.font.bold     = True
                            tp2.font.color.rgb = C_NAVY
                            # Place graphs
                            positions = [
                                (Inches(0.3),  Inches(0.8),
                                 Inches(6.3),  Inches(6.2)),
                                (Inches(6.8),  Inches(0.8),
                                 Inches(6.3),  Inches(6.2)),
                            ]
                            for gi, gkey in enumerate(chunk):
                                img_bytes = graph_images[gkey]
                                img_buf   = io.BytesIO(img_bytes)
                                lf, tp3, wd, ht = positions[gi]
                                sl.shapes.add_picture(img_buf, lf, tp3, wd, ht)
                            # Footer
                            ft2 = sl.shapes.add_textbox(
                                Inches(0.5), Inches(7.1),
                                Inches(12), Inches(0.3))
                            fp2 = ft2.text_frame.paragraphs[0]
                            fp2.text = (
                                f"SiteInsight · {project_name} · "
                                f"Week {selected_week:02d} · CONFIDENTIAL"
                            )
                            fp2.font.size     = Pt(8)
                            fp2.font.color.rgb = RGB(0x99, 0x99, 0x99)

                    # Save
                    ppt_buf = io.BytesIO()
                    prs.save(ppt_buf)
                    ppt_buf.seek(0)

                    st.success(
                        f"Presentation ready — "
                        f"{len(prs.slides)} slides generated."
                    )
                    st.download_button(
                        "⬇️ Download PowerPoint (.pptx)",
                        ppt_buf,
                        f"SiteInsight_Wk{selected_week:02d}_{ppt_audience.replace(' ','_')}.pptx",
                        "application/vnd.openxmlformats-officedocument"
                        ".presentationml.presentation",
                    )

                except ImportError:
                    st.warning(
                        "python-pptx not installed. "
                        "Run: pip install python-pptx --only-binary=:all:"
                    )

# SECTION: GENERATE GRAPHS
# ════════════════════════════════════════════════════════════════
elif nav == "📈 Generate Graphs":
    st.markdown("### 📈 Graph Generator")

    with st.form("graph_form"):
        gc1, gc2 = st.columns(2)
        with gc1:
            g_scurve   = st.checkbox("S-Curve (planned vs actual)",    value=True)
            g_variance = st.checkbox("Variance trend — all weeks",      value=True)
            g_phase    = st.checkbox("Phase-wise completion",           value=True)
        with gc2:
            g_slip     = st.checkbox("Weeks slip by activity",          value=True)
            g_delay    = st.checkbox("Delay reasons breakdown",         value=True)
            g_critical = st.checkbox("Critical path: planned vs actual",value=True)
            g_all      = st.checkbox("Generate ALL graphs",             value=False)
        gen_graphs = st.form_submit_button("Generate Graphs", type="primary")

    if gen_graphs:
        if g_all:
            g_scurve = g_variance = g_phase = g_slip = g_delay = g_critical = True

        st.divider()

        # 0 · S-Curve
        if g_scurve and not history_df.empty:
            st.markdown("#### S-Curve — Planned vs Actual Progress")
            scurve = history_df.groupby("week_number").agg(
                actual=("cum_actual_pct","mean"),
                planned=("cum_planned_pct","mean")
            ).reset_index()
            scurve[["actual","planned"]] *= 100

            fig, ax = plt.subplots(figsize=(11,4.5))
            ax.plot(scurve["week_number"], scurve["planned"],
                    color=NAVY, linewidth=2.5, label="Planned %",
                    linestyle="--", marker="o", markersize=3)
            ax.plot(scurve["week_number"], scurve["actual"],
                    color=AMBER, linewidth=2.5, label="Actual %",
                    marker="o", markersize=3)
            ax.fill_between(scurve["week_number"],
                            scurve["planned"], scurve["actual"],
                            where=scurve["actual"] < scurve["planned"],
                            alpha=0.12, color=RED, label="Variance gap")
            ax.axvline(selected_week, color=AMBER, linewidth=1.5,
                       linestyle=":", alpha=0.8, label=f"Wk{selected_week}")
            ax.set_xlabel("Week Number", fontsize=10)
            ax.set_ylabel("Cumulative Progress (%)", fontsize=10)
            ax.set_title(
                f"{project_name} · S-Curve: Planned vs Actual Progress",
                fontweight="bold", fontsize=11)
            ax.legend(fontsize=9)
            ax.grid(alpha=0.3)
            ax.set_facecolor("#FAFAFA")
            fig.patch.set_facecolor("white")
            plt.tight_layout()
            buf = io.BytesIO()
            fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
            buf.seek(0)
            st.image(buf, use_container_width=True)
            st.download_button("⬇️ Download S-Curve", buf,
                               f"scurve_Wk{selected_week:02d}.png",
                               "image/png", key="dl_sc")
            plt.close(fig); st.divider()

        # 1 · Variance trend
        if g_variance and not history_df.empty:
            st.markdown("#### Variance trend — all weeks")
            trend = (history_df.groupby("week_number")["variance_pct"]
                     .mean().reset_index())
            trend["variance_pct"] *= 100
            fig, ax = plt.subplots(figsize=(10,3.5))
            colors = [RED if v < 0 else GREEN for v in trend["variance_pct"]]
            ax.bar(trend["week_number"], trend["variance_pct"],
                   color=colors, width=0.6, zorder=3)
            ax.axhline(0, color=GRAY, linewidth=0.8, linestyle="--")
            ax.axvline(selected_week-0.5, color=AMBER, linewidth=1.5,
                       linestyle="--", label=f"Wk{selected_week}")
            ax.set_xlabel("Week"); ax.set_ylabel("Avg Variance (%)")
            ax.set_title(f"{project_name} · Average Variance by Week",
                         fontweight="bold")
            ax.legend(); ax.grid(axis="y", alpha=0.3)
            ax.set_facecolor("#FAFAFA"); fig.patch.set_facecolor("white")
            plt.tight_layout()
            buf = io.BytesIO()
            fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
            buf.seek(0)
            st.image(buf, use_container_width=True)
            st.download_button("⬇️ Download", buf,
                               f"variance_Wk{selected_week:02d}.png",
                               "image/png", key="dl_v")
            plt.close(fig); st.divider()

        # 2 · Phase-wise completion
        if g_phase and not activities.empty:
            st.markdown("#### Phase-wise completion — this week")
            pd_data = activities.groupby("phase").agg(
                actual=("cum_actual_pct","mean"),
                planned=("cum_planned_pct","mean")
            ).reset_index()
            pd_data[["actual","planned"]] *= 100
            fig, ax = plt.subplots(figsize=(9,4))
            x = range(len(pd_data)); w = 0.35
            b1 = ax.bar([i-w/2 for i in x], pd_data["planned"],
                        w, label="Planned %", color=NAVY, alpha=0.85)
            b2 = ax.bar([i+w/2 for i in x], pd_data["actual"],
                        w, label="Actual %", color=AMBER, alpha=0.85)
            ax.set_xticks(list(x))
            ax.set_xticklabels(pd_data["phase"])
            ax.set_ylabel("Completion (%)")
            ax.set_title(f"Phase-wise Status · Week {selected_week:02d}",
                         fontweight="bold")
            ax.legend(); ax.grid(axis="y", alpha=0.3)
            ax.set_facecolor("#FAFAFA"); fig.patch.set_facecolor("white")
            for b in list(b1)+list(b2):
                ax.text(b.get_x()+b.get_width()/2, b.get_height()+0.5,
                        f"{b.get_height():.0f}%",
                        ha="center", va="bottom", fontsize=8)
            plt.tight_layout()
            buf = io.BytesIO()
            fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
            buf.seek(0)
            st.image(buf, use_container_width=True)
            st.download_button("⬇️ Download", buf,
                               f"phase_Wk{selected_week:02d}.png",
                               "image/png", key="dl_p")
            plt.close(fig); st.divider()

        # 3 · Slip tracker
        if g_slip and not activities.empty:
            st.markdown("#### Weeks slip by activity")
            slip_data = activities[
                activities["weeks_slip"] > 0
            ].sort_values("weeks_slip", ascending=True)
            if slip_data.empty:
                st.success("No activities with weeks slip this week.")
            else:
                fig, ax = plt.subplots(
                    figsize=(9, max(3.5, len(slip_data)*0.45)))
                colors = [RED if row.get("is_critical_path") else AMBER
                          for _, row in slip_data.iterrows()]
                # Use phase · activity as label
                labels = [
                    f"{r.get('phase','')} · {r.get('activity','')}"
                    for _, r in slip_data.iterrows()
                ]
                bars = ax.barh(labels, slip_data["weeks_slip"],
                               color=colors, height=0.55)
                ax.set_xlabel("Weeks Slip")
                ax.set_title(f"Weeks Slip · Week {selected_week:02d}",
                             fontweight="bold")
                ax.grid(axis="x", alpha=0.3)
                ax.set_facecolor("#FAFAFA"); fig.patch.set_facecolor("white")
                from matplotlib.patches import Patch
                ax.legend(handles=[
                    Patch(color=RED,   label="Critical"),
                    Patch(color=AMBER, label="Non-critical")
                ])
                for b in bars:
                    ax.text(b.get_width()+0.05,
                            b.get_y()+b.get_height()/2,
                            f"{int(b.get_width())}w",
                            va="center", fontsize=8)
                plt.tight_layout()
                buf = io.BytesIO()
                fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
                buf.seek(0)
                st.image(buf, use_container_width=True)
                st.download_button("⬇️ Download", buf,
                                   f"slip_Wk{selected_week:02d}.png",
                                   "image/png", key="dl_s")
                plt.close(fig); st.divider()

        # 4 · Delay reasons
        if g_delay and not activities.empty:
            st.markdown("#### Delay reasons breakdown")
            delayed = activities[
                (activities["variance_pct"] < 0) &
                (activities["delay_reason"].notna()) &
                (activities["delay_reason"] != "No Delay")
            ]
            if delayed.empty:
                st.success("No delay reasons to display.")
            else:
                counts  = delayed["delay_reason"].value_counts()
                fig, ax = plt.subplots(figsize=(7,4.5))
                palette = [NAVY, AMBER, RED, AMBER2, GREEN,
                           "#7B61FF","#FF6B6B","#4ECDC4"]
                wedges, texts, autotexts = ax.pie(
                    counts.values, labels=counts.index,
                    autopct="%1.0f%%",
                    colors=palette[:len(counts)],
                    startangle=140, pctdistance=0.82,
                )
                for t in texts:      t.set_fontsize(9)
                for at in autotexts:
                    at.set_fontsize(8); at.set_color("white")
                    at.set_fontweight("bold")
                ax.set_title(f"Delay Reasons · Week {selected_week:02d}",
                             fontweight="bold")
                fig.patch.set_facecolor("white"); plt.tight_layout()
                buf = io.BytesIO()
                fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
                buf.seek(0)
                st.image(buf, use_container_width=True)
                st.download_button("⬇️ Download", buf,
                                   f"delay_Wk{selected_week:02d}.png",
                                   "image/png", key="dl_d")
                plt.close(fig); st.divider()

        # 5 · Critical path
        if (g_critical and not activities.empty
                and "is_critical_path" in activities.columns):
            st.markdown("#### Critical path: planned vs actual")
            cp = activities[activities["is_critical_path"] == True].copy()
            if cp.empty:
                st.info("No critical path activities found.")
            else:
                cp = cp.sort_values("variance_pct")
                labels = [
                    f"{r.get('phase','')} · {r.get('activity','')}"
                    for _, r in cp.iterrows()
                ]
                fig, ax = plt.subplots(
                    figsize=(9, max(3.5, len(cp)*0.55)))
                x = range(len(cp))
                ax.barh([i+0.2 for i in x], cp["cum_planned_pct"]*100,
                        height=0.35, label="Planned %",
                        color=NAVY, alpha=0.8)
                ax.barh([i-0.2 for i in x], cp["cum_actual_pct"]*100,
                        height=0.35, label="Actual %",
                        color=AMBER, alpha=0.8)
                ax.set_yticks(list(x))
                ax.set_yticklabels(labels, fontsize=8)
                ax.set_xlabel("Completion (%)")
                ax.set_title(f"Critical Path · Week {selected_week:02d}",
                             fontweight="bold")
                ax.legend(); ax.grid(axis="x", alpha=0.3)
                ax.set_facecolor("#FAFAFA"); fig.patch.set_facecolor("white")
                plt.tight_layout()
                buf = io.BytesIO()
                fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
                buf.seek(0)
                st.image(buf, use_container_width=True)
                st.download_button("⬇️ Download", buf,
                                   f"critical_Wk{selected_week:02d}.png",
                                   "image/png", key="dl_c")
                plt.close(fig)


# ════════════════════════════════════════════════════════════════
# SECTION: RECOVERY SIMULATOR
# ════════════════════════════════════════════════════════════════
elif nav == "🔄 Recovery Simulator":
    st.markdown("### 🔄 Recovery Simulator")
    st.caption(
        "Select a delayed activity and simulate catch-up scenarios "
        "based on different acceleration targets."
    )

    if activities.empty or history_df.empty:
        st.info("No activity data available. Load history first.")
    else:
        behind_acts = activities[activities["variance_pct"] < -0.05].copy()
        if behind_acts.empty:
            st.success("No delayed activities this week.")
        else:
            act_options = [
                f"{row['phase']} · {row['activity']} "
                f"({row['variance_pct']*100:.1f}%)"
                for _, row in behind_acts.iterrows()
            ]
            selected_act_label = st.selectbox(
                "Select delayed activity to simulate", act_options)
            selected_act_id = behind_acts.iloc[
                act_options.index(selected_act_label)]["act_id"]
            act_row = behind_acts[
                behind_acts["act_id"] == selected_act_id].iloc[0]

            current_pct   = float(act_row["cum_actual_pct"])
            planned_pct   = float(act_row["cum_planned_pct"])
            variance      = float(act_row["variance_pct"])
            weeks_slip    = int(act_row["weeks_slip"])
            delay_reason  = act_row.get("delay_reason","—")
            resp_person   = act_row.get("responsible_person","—") or "—"
            hist_velocity = compute_weekly_velocity(history_df, selected_act_id)

            st.divider()
            col_s1, col_s2, col_s3, col_s4, col_s5 = st.columns(5)
            col_s1.metric("Current actual", f"{current_pct*100:.1f}%")
            col_s2.metric("Planned target", f"{planned_pct*100:.1f}%")
            col_s3.metric("Variance",       f"{variance*100:.1f}%",
                          delta_color="inverse")
            col_s4.metric("Weeks slip",     f"{weeks_slip} wks",
                          delta_color="inverse")
            col_s5.metric("Responsible",    resp_person)
            st.caption(
                f"Delay reason: **{delay_reason}** · "
                f"Historical rate: **{hist_velocity*100:.2f}%/wk**"
            )

            st.divider()
            sc1, sc2 = st.columns(2)
            with sc1:
                acceleration = st.slider(
                    "Target weekly progress rate (%)",
                    min_value=1.0, max_value=25.0,
                    value=max(float(hist_velocity*100*1.5), 5.0),
                    step=0.5,
                )
            with sc2:
                total_project_weeks = st.number_input(
                    "Total project weeks",
                    min_value=20, max_value=200, value=80)

            accel_rate        = acceleration / 100
            realistic_rate    = hist_velocity if hist_velocity > 0 else 0.02
            optimistic_rate   = accel_rate
            conservative_rate = accel_rate * 0.6

            proj_realistic    = compute_projected_finish(
                current_pct, selected_week, realistic_rate, total_project_weeks)
            proj_optimistic   = compute_projected_finish(
                current_pct, selected_week, optimistic_rate, total_project_weeks)
            proj_conservative = compute_projected_finish(
                current_pct, selected_week, conservative_rate, total_project_weeks)

            st.divider()
            st.markdown("**Projected finish week**")
            r1, r2, r3 = st.columns(3)
            r1.metric("Conservative",
                      f"Week {min(proj_conservative, total_project_weeks)}",
                      delta_color="inverse")
            r2.metric("Target (your rate)",
                      f"Week {min(proj_optimistic, total_project_weeks)}",
                      delta_color="normal")
            r3.metric("At current rate",
                      f"Week {min(proj_realistic, total_project_weeks)}",
                      delta_color="inverse")

            st.divider()
            future_weeks = list(range(
                selected_week,
                min(selected_week+30, total_project_weeks+1)))

            def project_curve(start_pct, rate, weeks):
                pcts = []; p = start_pct
                for _ in weeks:
                    p = min(p+rate, 1.0); pcts.append(p*100)
                return pcts

            fig, ax = plt.subplots(figsize=(10,4))
            ax.plot(future_weeks,
                    project_curve(current_pct, optimistic_rate, future_weeks),
                    color=GREEN, linewidth=2,
                    label=f"Target ({acceleration:.1f}%/wk)")
            ax.plot(future_weeks,
                    project_curve(current_pct, conservative_rate, future_weeks),
                    color=AMBER2, linewidth=1.5, linestyle="--",
                    label=f"Conservative ({acceleration*0.6:.1f}%/wk)")
            ax.plot(future_weeks,
                    project_curve(current_pct, realistic_rate, future_weeks),
                    color=RED, linewidth=1.5, linestyle=":",
                    label=f"Current rate ({hist_velocity*100:.1f}%/wk)")
            ax.axhline(100, color=GRAY, linewidth=0.8,
                       linestyle="--", label="100% complete")
            ax.axhline(planned_pct*100, color=NAVY, linewidth=0.8,
                       linestyle="-.",
                       label=f"Planned ({planned_pct*100:.0f}%)")
            ax.set_xlabel("Week"); ax.set_ylabel("Cumulative Progress (%)")
            act_label = f"{act_row.get('phase','')} · {act_row['activity']}"
            ax.set_title(f"Recovery Simulation · {act_label}",
                         fontweight="bold")
            ax.legend(fontsize=9); ax.grid(alpha=0.3)
            ax.set_facecolor("#FAFAFA"); fig.patch.set_facecolor("white")
            plt.tight_layout()
            buf = io.BytesIO()
            fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
            buf.seek(0)
            st.image(buf, use_container_width=True)
            st.download_button("⬇️ Download simulation chart", buf,
                               f"recovery_{selected_act_id}_Wk{selected_week:02d}.png",
                               "image/png", key="dl_sim")
            plt.close(fig)

            st.divider()
            if st.button("Generate AI recovery recommendation", type="primary"):
                with st.spinner("Analysing via Groq..."):
                    prompt = f"""You are a senior construction planning consultant.
ACTIVITY: {act_row['activity']} ({selected_act_id})
PHASE: {act_row['phase']} | RESPONSIBLE: {resp_person}
CURRENT: {current_pct*100:.1f}% | PLANNED: {planned_pct*100:.1f}%
VARIANCE: {variance*100:.1f}% | WEEKS SLIP: {weeks_slip}
DELAY REASON: {delay_reason}
HISTORICAL RATE: {hist_velocity*100:.2f}%/wk | TARGET: {acceleration:.1f}%/wk
At current rate: Week {proj_realistic} | At target: Week {proj_optimistic}

Write a specific recovery recommendation (max 200 words):
1. Is {acceleration:.1f}%/week realistic given the delay reason?
2. What must {resp_person} do THIS WEEK specifically?
3. Consequence if not recovered by Week {selected_week+1}?
4. One direct instruction to the contractor.
Direct construction language only."""
                    rec = call_groq(prompt, max_tokens=400)
                st.markdown("**AI Recovery Recommendation**")
                st.warning(rec)


# ════════════════════════════════════════════════════════════════
# SECTION: EARLY WARNING
# ════════════════════════════════════════════════════════════════
elif nav == "⚠️ Early Warning":
    st.markdown("### ⚠️ Early Warning System")
    st.caption(
        "Identifies activities likely to slip NEXT week based on "
        "trend velocity — before it appears in the WPR."
    )

    LOOKBACK = 4

    if history_df.empty or activities.empty:
        st.info("Need at least 3 weeks of history for trend analysis.")
    else:
        warnings_list = []
        for _, act in activities.iterrows():
            aid      = act["act_id"]
            variance = float(act.get("variance_pct", 0) or 0)
            critical = bool(act.get("is_critical_path", False))
            slip     = int(act.get("weeks_slip", 0) or 0)
            resp     = act.get("responsible_person","—") or "—"
            slope    = compute_trend_slope(history_df, aid, LOOKBACK)
            risk_score = 0; risk_flags = []
            if slope < -0.015:
                risk_score += 3
                risk_flags.append(
                    f"Variance worsening {abs(slope)*100:.1f}%/wk over {LOOKBACK} wks")
            if variance < -0.05:
                risk_score += 2
                risk_flags.append(f"Already {abs(variance)*100:.1f}% behind plan")
            if slip >= 2:
                risk_score += 2
                risk_flags.append(f"{slip} weeks slip accumulated")
            if critical:
                risk_score += 2
                risk_flags.append("On critical path")
            if slope < 0 and variance < -0.03:
                risk_score += 1
                risk_flags.append("Negative trend + negative variance")
            if risk_score >= 3:
                risk_level = ("HIGH"   if risk_score >= 7 else
                              "MEDIUM" if risk_score >= 4 else "LOW")
                warnings_list.append({
                    "act_id":      aid,
                    "activity":    act["activity"],
                    "phase":       act["phase"],
                    "variance":    variance,
                    "slope":       slope,
                    "slip":        slip,
                    "critical":    critical,
                    "responsible": resp,
                    "risk_score":  risk_score,
                    "risk_level":  risk_level,
                    "flags":       risk_flags,
                })

        warnings_list.sort(key=lambda x: x["risk_score"], reverse=True)

        if not warnings_list:
            st.success("No early warning signals detected.")
        else:
            high   = [w for w in warnings_list if w["risk_level"]=="HIGH"]
            medium = [w for w in warnings_list if w["risk_level"]=="MEDIUM"]
            low    = [w for w in warnings_list if w["risk_level"]=="LOW"]

            w1, w2, w3 = st.columns(3)
            w1.metric("HIGH risk",   len(high),
                      delta="Immediate action" if high else None,
                      delta_color="inverse")
            w2.metric("MEDIUM risk", len(medium),
                      delta="Monitor closely" if medium else None,
                      delta_color="inverse")
            w3.metric("LOW risk",    len(low))
            st.divider()

            level_icons = {"HIGH":"🔴","MEDIUM":"🟡","LOW":"🟢"}
            for w in warnings_list:
                with st.expander(
                    f"{level_icons[w['risk_level']]} {w['risk_level']} · "
                    f"{w['phase']} · {w['activity']} "
                    f"(Score: {w['risk_score']})",
                    expanded=(w["risk_level"]=="HIGH"),
                ):
                    e1, e2, e3, e4, e5 = st.columns(5)
                    e1.metric("Variance",
                              f"{w['variance']*100:.1f}%",
                              delta_color="inverse")
                    e2.metric("Trend slope",
                              f"{w['slope']*100:.2f}%/wk",
                              delta_color=(
                                  "inverse" if w["slope"] < 0 else "normal"))
                    e3.metric("Weeks slip",
                              f"{w['slip']} wks",
                              delta_color=(
                                  "inverse" if w["slip"] > 0 else "normal"))
                    e4.metric("Critical",
                              "YES" if w["critical"] else "No")
                    e5.metric("Responsible", w["responsible"])
                    st.markdown("**Why flagged:**")
                    for flag in w["flags"]:
                        st.markdown(f"- {flag}")

            # Trend chart
            st.divider()
            st.markdown("#### Variance trend — at-risk activities")
            at_risk_ids = [w["act_id"] for w in warnings_list[:6]]
            hist_subset = history_df[history_df["act_id"].isin(at_risk_ids)]
            if not hist_subset.empty:
                fig, ax = plt.subplots(figsize=(10,4))
                for aid in at_risk_ids:
                    act_data = hist_subset[
                        hist_subset["act_id"]==aid
                    ].sort_values("week_number")
                    if len(act_data) < 2:
                        continue
                    risk_level = next(
                        (w["risk_level"] for w in warnings_list
                         if w["act_id"]==aid), "LOW")
                    color = (RED   if risk_level=="HIGH" else
                             AMBER if risk_level=="MEDIUM" else GRAY)
                    # Use phase · activity as legend label
                    w_item = next(
                        (w for w in warnings_list if w["act_id"]==aid), {})
                    leg_label = (
                        f"{w_item.get('phase','')} · "
                        f"{w_item.get('activity','')}"
                    )
                    ax.plot(act_data["week_number"],
                            act_data["variance_pct"]*100,
                            label=leg_label,
                            linewidth=2.5 if risk_level=="HIGH" else 1.5,
                            color=color, marker="o", markersize=3)
                ax.axhline(0,   color=GRAY, linewidth=0.8, linestyle="--")
                ax.axhline(-5,  color=RED,  linewidth=0.5,
                           linestyle=":", alpha=0.5)
                ax.axhline(-10, color=RED,  linewidth=0.8,
                           linestyle=":", alpha=0.7)
                ax.axvline(selected_week, color=AMBER,
                           linewidth=1, linestyle="--", alpha=0.7)
                ax.set_xlabel("Week"); ax.set_ylabel("Variance (%)")
                ax.set_title("Variance Trajectory — At-risk Activities",
                             fontweight="bold")
                ax.legend(fontsize=7, ncol=2)
                ax.grid(alpha=0.3); ax.set_facecolor("#FAFAFA")
                fig.patch.set_facecolor("white"); plt.tight_layout()
                buf = io.BytesIO()
                fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
                buf.seek(0)
                st.image(buf, use_container_width=True)
                st.download_button("⬇️ Download early warning chart", buf,
                                   f"early_warning_Wk{selected_week:02d}.png",
                                   "image/png", key="dl_ew")
                plt.close(fig)

            st.divider()
            if st.button("Generate proactive weekly brief", type="primary"):
                with st.spinner("Generating via Groq..."):
                    high_names = [
                        f"{w['phase']} · {w['activity']} "
                        f"(responsible: {w['responsible']}, "
                        f"slope {w['slope']*100:.2f}%/wk)"
                        for w in high
                    ]
                    med_names = [
                        f"{w['phase']} · {w['activity']} "
                        f"(responsible: {w['responsible']})"
                        for w in medium
                    ]
                    prompt = f"""You are a proactive construction project intelligence system.
Week {selected_week} of {project_name}.

EARLY WARNING ({LOOKBACK}-week trend analysis):
HIGH RISK — will worsen next week: {high_names if high_names else 'None'}
MEDIUM RISK — monitor: {med_names if med_names else 'None'}

Write a PROACTIVE brief (max 250 words):
1. Predict what happens in Week {selected_week+1} if no action taken
2. Name each responsible person and their specific required action THIS WEEK
3. Identify milestone / critical path impact
4. End with ONE thing the PM must do before next site meeting

This is intelligence — tell the team what WILL happen, not what HAS happened.
Direct, action-oriented construction language."""
                    brief = call_groq(prompt, max_tokens=500)
                st.markdown("**Proactive Weekly Brief — Week ahead**")
                st.warning(brief)
