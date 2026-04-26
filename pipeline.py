"""
pipeline.py
───────────
SiteInsight · Live WPR Pipeline
Runs for Wk11 onward. Each call processes ONE new WPR file:
  1. Parse Excel
  2. DQ validation gate
  3. Run 4-rule detection engine
  4. Call Groq API for narrative
  5. Store everything in Supabase

Usage:
    python pipeline.py --file ./wpr_files/WPR_Wk11.xlsx
    python pipeline.py --file ./wpr_files/WPR_Wk11.xlsx --dry-run
"""

import os
import json
import argparse
from pathlib import Path
from datetime import datetime

from groq import Groq
from supabase import create_client, Client
from dotenv import load_dotenv

# Reuse parsers from load_history
from load_history import (
    parse_header,
    parse_activities,
    run_dq_checks,
    load_wpr_file as _load_wpr_file,
)

load_dotenv()

SUPABASE_URL = os.environ["SUPABASE_URL"]
SUPABASE_KEY = os.environ["SUPABASE_KEY"]
GROQ_API_KEY = os.environ["GROQ_API_KEY"]

supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)
groq_client      = Groq(api_key=GROQ_API_KEY)

GROQ_MODEL = "llama-3.3-70b-versatile"   # fast + free tier friendly

# ── Detection Engine ─────────────────────────────────────────

def run_detection(activities: list[dict], week_number: int,
                  history: list[dict]) -> list[dict]:
    """
    4 detection rules. Returns list of detection_result dicts.

    R1 — Schedule slip          : variance < -0.05 AND weeks_slip > 0
    R2 — Scope creep            : act_id not seen in any previous week
    R3 — Stale issue            : same act_id had same delay_reason last week
                                  AND variance is still negative
    R4 — Critical path at risk  : is_critical_path=True AND variance < -0.10
    """
    hits = []

    # Build lookup sets from history
    prev_week    = week_number - 1
    hist_act_ids = {r["act_id"] for r in history}
    prev_week_acts = {r["act_id"]: r for r in history
                      if r.get("week_number") == prev_week}

    for act in activities:
        aid      = act["act_id"]
        variance = act.get("variance_pct") or 0
        slip     = act.get("weeks_slip")   or 0
        reason   = act.get("delay_reason", "")
        critical = act.get("is_critical_path", False)
        phase    = act.get("phase", "")
        activity = act.get("activity", "")
        baseline = act.get("baseline_version", "")

        base = {
            "week_number":      week_number,
            "act_id":           aid,
            "variance_pct":     variance,
            "weeks_slip":       slip,
            "baseline_version": baseline,
        }

        # R1 · Schedule slip
        if variance < -0.05 and slip > 0:
            severity = "high" if variance < -0.15 else "medium"
            hits.append({**base,
                "rule_id":   "R1",
                "rule_name": "Schedule slip",
                "severity":  severity,
                "detail":    (f"{phase} · {activity} is {abs(variance)*100:.1f}% behind plan "
                              f"({slip} wk slip). Reason: {reason}"),
            })

        # R2 · Scope creep — activity never seen before
        if aid not in hist_act_ids:
            hits.append({**base,
                "rule_id":   "R2",
                "rule_name": "Scope creep",
                "severity":  "medium",
                "detail":    (f"New activity '{activity}' ({aid}) not in any previous WPR. "
                              f"Verify against {baseline} baseline."),
            })

        # R3 · Stale issue — same delay reason, still behind
        prev = prev_week_acts.get(aid)
        if prev:
            prev_reason   = prev.get("delay_reason", "")
            prev_variance = prev.get("variance_pct") or 0
            if (prev_reason == reason
                    and reason not in ("No Delay", "", None)
                    and variance < 0
                    and prev_variance < 0):
                hits.append({**base,
                    "rule_id":   "R3",
                    "rule_name": "Stale issue",
                    "severity":  "medium",
                    "detail":    (f"{activity} ({aid}) has carried same delay reason "
                                  f"'{reason}' for 2+ consecutive weeks with no recovery."),
                })

        # R4 · Critical path at risk
        if critical and variance < -0.10:
            hits.append({**base,
                "rule_id":   "R4",
                "rule_name": "Critical path at risk",
                "severity":  "high",
                "detail":    (f"CRITICAL PATH: {activity} ({aid}) is "
                              f"{abs(variance)*100:.1f}% behind. "
                              f"Immediate recovery plan required."),
            })

    return hits


# ── Groq Narrative Generator ─────────────────────────────────

def build_prompt(week_number: int, header: dict, activities: list[dict],
                 detections: list[dict], dq_flags: list[dict]) -> str:
    """Build the prompt sent to Groq."""
    slip_acts   = [d for d in detections if d["rule_id"] == "R1"]
    creep_acts  = [d for d in detections if d["rule_id"] == "R2"]
    stale_acts  = [d for d in detections if d["rule_id"] == "R3"]
    critical    = [d for d in detections if d["rule_id"] == "R4"]

    behind_acts = [a for a in activities if (a.get("variance_pct") or 0) < 0]
    ahead_acts  = [a for a in activities if (a.get("variance_pct") or 0) > 0.05]

    prompt = f"""You are a senior construction project controls analyst.
Analyse the Week {week_number} Weekly Progress Report for {header.get('project_name')} 
and produce a structured intelligence report.

REPORT CONTEXT:
- Active baseline : {header.get('baseline_version')}
- Baseline revised this week: {header.get('baseline_revised_this_week')}
- Contractor: {header.get('contractor')}

DETECTION ENGINE RESULTS ({len(detections)} flags):
"""
    for d in detections:
        prompt += f"  [{d['rule_id']} · {d['severity'].upper()}] {d['detail']}\n"

    if not detections:
        prompt += "  No flags raised this week.\n"

    prompt += f"""
ACTIVITIES BEHIND PLAN ({len(behind_acts)}):
"""
    for a in behind_acts:
        prompt += (f"  {a.get('phase','')} · {a['activity']} · "
                   f"Variance: {a.get('variance_pct',0)*100:.1f}% · "
                   f"Reason: {a.get('delay_reason','')} · "
                   f"Responsible: {a.get('responsible_person','—')}\n")

    if dq_flags:
        prompt += f"\nDATA QUALITY FLAGS ({len(dq_flags)}):\n"
        for f in dq_flags:
            prompt += f"  {f.get('act_id')} · {f.get('rule_triggered')} · {f.get('raw_value')}\n"

    prompt += """
Produce a JSON response with exactly these keys:
{
  "executive_summary": "2-3 sentence overall project status for this week",
  "key_risks": "Bullet-point list of top 3 risks identified",
  "recommendations": "Bullet-point list of top 3 recommended actions for next week"
}

Rules:
- Be specific: name activities, phases, percentages
- Use plain English, no jargon
- Keep each section under 150 words
- Return ONLY the JSON object, no preamble, no markdown fences
"""
    return prompt


def call_groq(prompt: str) -> dict:
    """Call Groq API and return parsed JSON + token counts."""
    response = groq_client.chat.completions.create(
        model=GROQ_MODEL,
        messages=[{"role": "user", "content": prompt}],
        temperature=0.3,
        max_tokens=800,
    )
    raw = response.choices[0].message.content.strip()
    usage = response.usage

    # Strip markdown fences if model adds them despite instructions
    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]
    raw = raw.strip()

    try:
        parsed = json.loads(raw)
    except json.JSONDecodeError:
        parsed = {
            "executive_summary": raw,
            "key_risks":         "Could not parse structured response.",
            "recommendations":   "Review raw Groq output.",
        }

    # Ensure key_risks and recommendations are stored as clean strings
    if isinstance(parsed.get("key_risks"), list):
        parsed["key_risks"] = json.dumps(parsed["key_risks"])
    if isinstance(parsed.get("recommendations"), list):
        parsed["recommendations"] = json.dumps(parsed["recommendations"])

    return {
        "parsed":            parsed,
        "full_response":     raw,
        "prompt_tokens":     usage.prompt_tokens,
        "completion_tokens": usage.completion_tokens,
    }


# ── Main Pipeline ─────────────────────────────────────────────

def run_pipeline(filepath: Path, dry_run: bool = False) -> dict:
    """Full pipeline for one WPR file."""
    from openpyxl import load_workbook

    print(f"\n{'='*55}")
    print(f"  SiteInsight Pipeline · {filepath.name}")
    print(f"  Mode: {'DRY RUN' if dry_run else 'LIVE'}")
    print(f"{'='*55}")

    wb = load_workbook(str(filepath), read_only=True)
    ws = wb["WPR Data"]

    # ── Step 1: Parse ─────────────────────────────────────────
    print("\n[1/5] Parsing WPR Excel...")
    header     = parse_header(ws)
    activities = parse_activities(ws, header["week_number"])
    week_number = header["week_number"]
    print(f"      Week {week_number} | {len(activities)} activities | "
          f"Baseline: {header['baseline_version']}")

    # ── Step 2: DQ Validation Gate ────────────────────────────
    print("\n[2/5] Running DQ validation...")
    dq_flags = run_dq_checks(activities, week_number)
    errors   = [f for f in dq_flags if f["severity"] == "error"]
    warnings = [f for f in dq_flags if f["severity"] == "warning"]
    print(f"      Errors: {len(errors)} | Warnings: {len(warnings)}")

    if errors:
        print("      [DQ ERRORS — these activities are flagged but pipeline continues]")
        for e in errors:
            print(f"      ✗ {e['act_id']} · {e['rule_triggered']} · {e['raw_value']}")

    # ── Step 3: Fetch History for Detection ───────────────────
    print(f"\n[3/5] Fetching history (Wk01–{week_number-1}) for detection context...")
    if not dry_run:
        resp = (supabase.table("wpr_activities")
                .select("*")
                .lt("week_number", week_number)
                .execute())
        history = resp.data or []
    else:
        history = []
    print(f"      {len(history)} historical activity records loaded")

    # ── Step 4: Detection Engine ──────────────────────────────
    print("\n[4/5] Running detection engine (R1–R4)...")
    detections = run_detection(activities, week_number, history)
    by_rule = {}
    for d in detections:
        by_rule.setdefault(d["rule_id"], []).append(d)
    for rule_id, hits in sorted(by_rule.items()):
        print(f"      {rule_id}: {len(hits)} hit(s)")
    print(f"      Total flags: {len(detections)}")

    # ── Step 5: Groq Narrative ────────────────────────────────
    print("\n[5/5] Generating narrative via Groq...")
    prompt  = build_prompt(week_number, header, activities, detections, dq_flags)
    if not dry_run:
        groq_result = call_groq(prompt)
        narrative   = groq_result["parsed"]
        print(f"      Tokens used: {groq_result['prompt_tokens']} prompt + "
              f"{groq_result['completion_tokens']} completion")
        print(f"      Executive summary: {narrative.get('executive_summary','')[:80]}...")
    else:
        print("      [DRY RUN] Skipping Groq call")
        groq_result = {"parsed": {}, "full_response": "", "prompt_tokens": 0,
                       "completion_tokens": 0}
        narrative   = {}

    # ── Store to Supabase ─────────────────────────────────────
    if not dry_run:
        print("\n[DB] Writing to Supabase...")

        # Header
        supabase.table("wpr_headers").upsert(
            {**header, "source_file": filepath.name, "load_type": "live"},
            on_conflict="week_number"
        ).execute()

        # Activities
        supabase.table("wpr_activities").upsert(
            activities, on_conflict="week_number,act_id"
        ).execute()

        # DQ flags
        if dq_flags:
            supabase.table("dq_flags").insert(dq_flags).execute()

        # Detection results
        if detections:
            supabase.table("detection_results").insert(detections).execute()

        # Narrative
        supabase.table("ai_narratives").upsert({
            "week_number":       week_number,
            "model_used":        GROQ_MODEL,
            "prompt_tokens":     groq_result["prompt_tokens"],
            "completion_tokens": groq_result["completion_tokens"],
            "executive_summary": narrative.get("executive_summary"),
            "key_risks":         narrative.get("key_risks"),
            "recommendations":   narrative.get("recommendations"),
            "full_response":     groq_result["full_response"],
        }, on_conflict="week_number").execute()

        print("      Done.")

    print(f"\n{'='*55}")
    print(f"  Pipeline complete · Week {week_number}")
    print(f"  Activities: {len(activities)} | DQ flags: {len(dq_flags)} | "
          f"Detections: {len(detections)}")
    print(f"{'='*55}\n")

    return {
        "week_number": week_number,
        "activities":  len(activities),
        "dq_flags":    len(dq_flags),
        "detections":  len(detections),
        "status":      "ok" if not dry_run else "dry_run",
    }


def main():
    parser = argparse.ArgumentParser(description="SiteInsight live WPR pipeline")
    parser.add_argument("--file",    required=True, help="Path to WPR Excel file")
    parser.add_argument("--dry-run", action="store_true",
                        help="Parse and detect only, skip Groq + Supabase writes")
    args = parser.parse_args()

    filepath = Path(args.file)
    if not filepath.exists():
        print(f"[ERROR] File not found: {filepath}")
        return

    run_pipeline(filepath, dry_run=args.dry_run)


if __name__ == "__main__":
    main()
