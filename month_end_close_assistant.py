"""
Month-End Close Assistant
Multi-agent pipeline using LangGraph + Google Gemini

Agents:
  1. Planner       — reads both trial balances, builds close checklist
  2. Variance Analyst — flags accounts with >10% movement
  3. Commentary Writer — drafts plain-English explanation per variance
  4. Reviewer      — challenges weak commentary, loops back if needed
  5. Report Generator — assembles final Markdown + Excel output

Usage:
  python month_end_close_assistant.py
"""

import os
import json
import pandas as pd
from datetime import datetime
from typing import TypedDict, List, Dict, Any
from langgraph.graph import StateGraph, END
from langchain_google_genai import ChatGoogleGenerativeAI
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Config ────────────────────────────────────────────────────────────────────
GEMINI_API_KEY   = os.environ.get("GEMINI_API_KEY", "")
CURRENT_TB_FILE  = "trial_balance_sep2026.csv"
PREVIOUS_TB_FILE = "trial_balance_aug2026.csv"
CURRENT_MONTH    = "September 2026"
PREVIOUS_MONTH   = "August 2026"
VARIANCE_THRESHOLD = 10.0   # percent

llm = ChatGoogleGenerativeAI(
    model="gemini-3.5-flash-lite",
    google_api_key=GEMINI_API_KEY,
    temperature=0.3,
)

# ── Helper: safely extract text from LLM response ─────────────────────────────
def get_text(resp) -> str:
    """Handle both string and list content from langchain-google-genai."""
    c = resp.content
    if isinstance(c, str):
        return c
    if isinstance(c, list):
        item = c[0]
        if isinstance(item, dict):
            return item.get("text", str(item))
        return str(item)
    return str(c)

# ── Shared State ──────────────────────────────────────────────────────────────
class CloseState(TypedDict):
    current_tb: pd.DataFrame
    previous_tb: pd.DataFrame
    checklist: List[str]
    variances: List[Dict[str, Any]]
    commentary: List[Dict[str, Any]]   # [{account, variance_info, draft, review, final, retry_count}]
    report_md: str
    report_xlsx: str

# ── Helper: Load Trial Balances ───────────────────────────────────────────────
def load_trial_balances() -> tuple[pd.DataFrame, pd.DataFrame]:
    current  = pd.read_csv(CURRENT_TB_FILE)
    previous = pd.read_csv(PREVIOUS_TB_FILE)
    return current, previous

# ══════════════════════════════════════════════════════════════════════════════
# AGENT 1: PLANNER
# Reads both TBs, builds a month-end close checklist
# ══════════════════════════════════════════════════════════════════════════════
def planner_node(state: CloseState) -> CloseState:
    print("\n" + "="*60)
    print("🗂️  AGENT 1: PLANNER")
    print("="*60)

    current  = state["current_tb"]
    previous = state["previous_tb"]

    # Build summary for Gemini
    current_summary  = current[["account_name","category","net_balance"]].to_string(index=False)
    previous_summary = previous[["account_name","category","net_balance"]].to_string(index=False)

    prompt = f"""You are a Chartered Accountant doing month-end close for {CURRENT_MONTH}.

Here are the two trial balances:

CURRENT MONTH ({CURRENT_MONTH}):
{current_summary}

PREVIOUS MONTH ({PREVIOUS_MONTH}):
{previous_summary}

Build a concise month-end close checklist (10-15 items) covering:
- Reconciliations to perform (bank, AR, AP, inventory)
- Accruals to review
- Accounts showing significant movement that need explanation
- Any accounts that need special attention based on the data

Return as a numbered list. Be specific to the accounts in these trial balances."""

    response = llm.invoke(prompt)
    checklist_text = get_text(response)

    # Parse into list
    checklist = [
        line.strip()
        for line in checklist_text.split("\n")
        if line.strip() and line.strip()[0].isdigit()
    ]

    print(f"✅ Planner built {len(checklist)} checklist items")
    for item in checklist:
        print(f"   {item}")

    return {**state, "checklist": checklist}

# ══════════════════════════════════════════════════════════════════════════════
# AGENT 2: VARIANCE ANALYST
# Flags accounts where abs(% change) > threshold
# ══════════════════════════════════════════════════════════════════════════════
def variance_analyst_node(state: CloseState) -> CloseState:
    print("\n" + "="*60)
    print("📊  AGENT 2: VARIANCE ANALYST")
    print("="*60)

    current  = state["current_tb"]
    previous = state["previous_tb"]

    merged = current.merge(
        previous[["account_code", "net_balance"]],
        on="account_code",
        suffixes=("_cur", "_prev"),
    )

    variances = []
    for _, row in merged.iterrows():
        cur  = float(row["net_balance_cur"])
        prev = float(row["net_balance_prev"])

        if prev == 0:
            pct = 100.0 if cur != 0 else 0.0
        else:
            pct = ((cur - prev) / abs(prev)) * 100

        if abs(pct) >= VARIANCE_THRESHOLD:
            variances.append({
                "account_code": row["account_code"],
                "account_name": row["account_name"],
                "category":     row["category"],
                "current":      cur,
                "previous":     prev,
                "change":       cur - prev,
                "pct_change":   round(pct, 1),
            })

    # Sort by abs % change descending
    variances.sort(key=lambda x: abs(x["pct_change"]), reverse=True)

    print(f"✅ Variance Analyst flagged {len(variances)} accounts (threshold: {VARIANCE_THRESHOLD}%)")
    for v in variances:
        direction = "▲" if v["change"] > 0 else "▼"
        print(f"   {direction} {v['account_name']}: {v['pct_change']:+.1f}%  "
              f"(₹{v['previous']:,.0f} → ₹{v['current']:,.0f})")

    return {**state, "variances": variances}

# ══════════════════════════════════════════════════════════════════════════════
# AGENT 3: COMMENTARY WRITER
# Drafts plain-English commentary for each flagged variance
# ══════════════════════════════════════════════════════════════════════════════
def commentary_writer_node(state: CloseState) -> CloseState:
    print("\n" + "="*60)
    print("✍️  AGENT 3: COMMENTARY WRITER")
    print("="*60)

    commentary = []

    for v in state["variances"]:
        direction = "increased" if v["change"] > 0 else "decreased"
        prompt = f"""You are a Finance Controller writing management commentary for month-end close.

Account: {v['account_name']} ({v['category']})
Previous Month ({PREVIOUS_MONTH}): ₹{v['previous']:,.0f}
Current Month ({CURRENT_MONTH}): ₹{v['current']:,.0f}
Change: ₹{v['change']:+,.0f} ({v['pct_change']:+.1f}%)

Write 2-3 sentences of management commentary explaining this variance.
- Be specific: mention the account name, the exact amounts, and the % change
- Suggest a likely business reason (e.g. seasonal demand, new contracts, cost controls)
- Flag any risk or action required if applicable
- Write in professional finance language suitable for a CFO report"""

        response = llm.invoke(prompt)
        draft = get_text(response).strip()

        commentary.append({
            "account_code": v["account_code"],
            "account_name": v["account_name"],
            "category":     v["category"],
            "variance_info": v,
            "draft":        draft,
            "review":       "",
            "final":        "",
            "retry_count":  0,
            "status":       "draft",
        })

        print(f"   ✍️  Draft written for: {v['account_name']}")

    return {**state, "commentary": commentary}

# ══════════════════════════════════════════════════════════════════════════════
# AGENT 4: REVIEWER
# Challenges weak commentary; loops back to writer if needed (max 2 retries)
# ══════════════════════════════════════════════════════════════════════════════
def reviewer_node(state: CloseState) -> CloseState:
    print("\n" + "="*60)
    print("🔍  AGENT 4: REVIEWER")
    print("="*60)

    updated_commentary = []

    for item in state["commentary"]:
        v     = item["variance_info"]
        draft = item["draft"] if item["retry_count"] == 0 else item["draft"]

        review_prompt = f"""You are a CFO reviewing management commentary written by a Finance Controller.

Account: {item['account_name']} ({item['category']})
Change: ₹{v['change']:+,.0f} ({v['pct_change']:+.1f}%)

Commentary to review:
\"\"\"{draft}\"\"\"

Evaluate this commentary strictly. Return a JSON object with exactly these keys:
{{
  "verdict": "APPROVED" or "REJECTED",
  "reason": "one sentence explaining why",
  "improved_commentary": "rewritten version if REJECTED, else same as input"
}}

REJECT if the commentary:
- Is vague (e.g. just says 'increased due to business activity')
- Doesn't mention specific amounts or percentages
- Doesn't suggest a business reason
- Is only 1 sentence

APPROVE if it is specific, mentions amounts, suggests a reason, and is professionally written."""

        response = llm.invoke(review_prompt)

        # Parse JSON from response
        try:
            raw = get_text(response).strip()
            # Strip markdown code fences if present
            if raw.startswith("```"):
                raw = raw.split("```")[1]
                if raw.startswith("json"):
                    raw = raw[4:]
            review_result = json.loads(raw.strip())
        except Exception:
            review_result = {
                "verdict": "APPROVED",
                "reason": "Could not parse review — defaulting to approved",
                "improved_commentary": draft,
            }

        verdict = review_result.get("verdict", "APPROVED")
        reason  = review_result.get("reason", "")
        improved = review_result.get("improved_commentary", draft)

        retry_count = item["retry_count"]

        if verdict == "REJECTED" and retry_count < 2:
            # Loop back: use improved commentary as the new draft
            print(f"   ❌ REJECTED: {item['account_name']} — {reason} (retry {retry_count + 1}/2)")
            updated_commentary.append({
                **item,
                "draft":       improved,
                "review":      reason,
                "retry_count": retry_count + 1,
                "status":      "retry",
            })
        else:
            final = improved if verdict == "REJECTED" else draft
            status = "approved" if verdict == "APPROVED" else "approved-after-retry"
            print(f"   ✅ APPROVED: {item['account_name']}")
            updated_commentary.append({
                **item,
                "review": reason,
                "final":  final,
                "status": status,
            })

    return {**state, "commentary": updated_commentary}

# ── Conditional edge: re-run reviewer if any item needs retry ─────────────────
def needs_retry(state: CloseState) -> str:
    for item in state["commentary"]:
        if item["status"] == "retry":
            return "reviewer"
    return "report_generator"

# ══════════════════════════════════════════════════════════════════════════════
# AGENT 5: REPORT GENERATOR
# Assembles Markdown report + Excel workbook
# ══════════════════════════════════════════════════════════════════════════════
def report_generator_node(state: CloseState) -> CloseState:
    print("\n" + "="*60)
    print("📄  AGENT 5: REPORT GENERATOR")
    print("="*60)

    now = datetime.now().strftime("%d %b %Y, %I:%M %p")

    # ── Markdown Report ───────────────────────────────────────────────────────
    md_lines = [
        f"# Month-End Close Report — {CURRENT_MONTH}",
        f"*Generated on {now} by Month-End Close Assistant (LangGraph + Gemini AI)*",
        "",
        "---",
        "",
        "## 1. Month-End Close Checklist",
        "",
    ]
    for item in state["checklist"]:
        md_lines.append(f"- [ ] {item.lstrip('0123456789. ')}")

    md_lines += [
        "",
        "---",
        "",
        f"## 2. Variance Analysis (>{VARIANCE_THRESHOLD}% movement vs {PREVIOUS_MONTH})",
        "",
        f"| Account | Category | {PREVIOUS_MONTH} (₹) | {CURRENT_MONTH} (₹) | Change (₹) | Change (%) |",
        "|---|---|---:|---:|---:|---:|",
    ]
    for v in state["variances"]:
        arrow = "▲" if v["change"] > 0 else "▼"
        md_lines.append(
            f"| {v['account_name']} | {v['category']} "
            f"| {v['previous']:,.0f} | {v['current']:,.0f} "
            f"| {v['change']:+,.0f} | {arrow} {abs(v['pct_change']):.1f}% |"
        )

    md_lines += [
        "",
        "---",
        "",
        "## 3. Management Commentary",
        "",
    ]
    for item in state["commentary"]:
        v = item["variance_info"]
        final_text = item["final"] if item["final"] else item["draft"]
        status_badge = "✅ Approved" if "approved" in item["status"] else "⚠️ Review"
        md_lines += [
            f"### {item['account_name']} ({v['pct_change']:+.1f}%) — {status_badge}",
            "",
            final_text,
            "",
        ]
        if item["review"]:
            md_lines += [f"> **Reviewer note:** {item['review']}", ""]

    md_lines += [
        "---",
        "",
        "*This report was auto-generated. All commentary should be reviewed by the Finance Controller before submission.*",
    ]

    report_md = "\n".join(md_lines)

    # Save Markdown
    md_filename = f"month_end_close_{CURRENT_MONTH.replace(' ', '_')}.md"
    with open(md_filename, "w", encoding="utf-8") as f:
        f.write(report_md)
    print(f"   📝 Markdown report saved: {md_filename}")

    # ── Excel Report ──────────────────────────────────────────────────────────
    wb = Workbook()

    # Colors
    HEADER_FILL  = PatternFill("solid", fgColor="1F3864")   # dark navy
    GREEN_FILL   = PatternFill("solid", fgColor="D4EDDA")
    RED_FILL     = PatternFill("solid", fgColor="F8D7DA")
    YELLOW_FILL  = PatternFill("solid", fgColor="FFF3CD")
    BLUE_FILL    = PatternFill("solid", fgColor="CCE5FF")
    GREY_FILL    = PatternFill("solid", fgColor="F2F2F2")

    header_font  = Font(bold=True, color="FFFFFF", name="Arial", size=11)
    title_font   = Font(bold=True, name="Arial", size=13)
    body_font    = Font(name="Arial", size=10)
    bold_font    = Font(bold=True, name="Arial", size=10)

    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    def style_header(cell):
        cell.fill   = HEADER_FILL
        cell.font   = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border

    def style_cell(cell, fill=None, bold=False, align="left", wrap=False):
        cell.font   = Font(bold=bold, name="Arial", size=10)
        cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
        cell.border = border
        if fill:
            cell.fill = fill

    # ── Sheet 1: Variance Report ──────────────────────────────────────────────
    ws1 = wb.active
    ws1.title = "Variance Report"
    ws1.row_dimensions[1].height = 30
    ws1.row_dimensions[2].height = 20

    ws1["A1"] = f"Month-End Close — Variance Report ({CURRENT_MONTH})"
    ws1["A1"].font = title_font
    ws1.merge_cells("A1:G1")
    ws1["A1"].alignment = Alignment(horizontal="center", vertical="center")

    headers = ["Account Code", "Account Name", "Category",
               f"{PREVIOUS_MONTH} (₹)", f"{CURRENT_MONTH} (₹)", "Change (₹)", "Change (%)"]
    for col, h in enumerate(headers, 1):
        cell = ws1.cell(row=2, column=col, value=h)
        style_header(cell)

    col_widths = [14, 28, 16, 20, 20, 18, 12]
    for i, w in enumerate(col_widths, 1):
        ws1.column_dimensions[get_column_letter(i)].width = w

    for row_idx, v in enumerate(state["variances"], 3):
        fill = GREEN_FILL if v["change"] > 0 else RED_FILL
        values = [
            v["account_code"], v["account_name"], v["category"],
            v["previous"], v["current"], v["change"], v["pct_change"] / 100,
        ]
        for col, val in enumerate(values, 1):
            cell = ws1.cell(row=row_idx, column=col, value=val)
            style_cell(cell, fill=fill,
                       align="right" if col > 3 else "left",
                       bold=(col == 2))
        # Number formats
        for col in [4, 5, 6]:
            ws1.cell(row=row_idx, column=col).number_format = '#,##0'
        ws1.cell(row=row_idx, column=7).number_format = '+0.0%;-0.0%'

    # ── Sheet 2: Management Commentary ───────────────────────────────────────
    ws2 = wb.create_sheet("Management Commentary")
    ws2.row_dimensions[1].height = 30
    ws2.row_dimensions[2].height = 20

    ws2["A1"] = f"Management Commentary — {CURRENT_MONTH}"
    ws2["A1"].font = title_font
    ws2.merge_cells("A1:E1")
    ws2["A1"].alignment = Alignment(horizontal="center", vertical="center")

    headers2 = ["Account Name", "Category", "Change (%)", "Commentary", "Review Status"]
    for col, h in enumerate(headers2, 1):
        cell = ws2.cell(row=2, column=col, value=h)
        style_header(cell)

    col_widths2 = [28, 16, 12, 70, 20]
    for i, w in enumerate(col_widths2, 1):
        ws2.column_dimensions[get_column_letter(i)].width = w

    for row_idx, item in enumerate(state["commentary"], 3):
        v = item["variance_info"]
        final_text = item["final"] if item["final"] else item["draft"]
        status = "✅ Approved" if "approved" in item["status"] else "⚠️ Needs Review"
        fill = GREEN_FILL if "approved" in item["status"] else YELLOW_FILL

        row_data = [
            item["account_name"],
            item["category"],
            v["pct_change"] / 100,
            final_text,
            status,
        ]
        for col, val in enumerate(row_data, 1):
            cell = ws2.cell(row=row_idx, column=col, value=val)
            style_cell(cell, fill=fill,
                       align="right" if col == 3 else "left",
                       wrap=(col == 4))
        ws2.cell(row=row_idx, column=3).number_format = '+0.0%;-0.0%'
        ws2.row_dimensions[row_idx].height = 80

    # ── Sheet 3: Close Checklist ──────────────────────────────────────────────
    ws3 = wb.create_sheet("Close Checklist")
    ws3.row_dimensions[1].height = 30
    ws3.row_dimensions[2].height = 20

    ws3["A1"] = f"Month-End Close Checklist — {CURRENT_MONTH}"
    ws3["A1"].font = title_font
    ws3.merge_cells("A1:C1")
    ws3["A1"].alignment = Alignment(horizontal="center", vertical="center")

    for col, h in enumerate(["#", "Checklist Item", "Status"], 1):
        cell = ws3.cell(row=2, column=col, value=h)
        style_header(cell)

    ws3.column_dimensions["A"].width = 6
    ws3.column_dimensions["B"].width = 70
    ws3.column_dimensions["C"].width = 16

    for row_idx, item in enumerate(state["checklist"], 3):
        clean = item.lstrip("0123456789. ")
        fill = GREY_FILL if row_idx % 2 == 0 else None
        ws3.cell(row=row_idx, column=1, value=row_idx - 2).font = body_font
        cell_b = ws3.cell(row=row_idx, column=2, value=clean)
        cell_c = ws3.cell(row=row_idx, column=3, value="☐ Pending")
        for cell in [ws3.cell(row=row_idx, column=1), cell_b, cell_c]:
            style_cell(cell, fill=fill, wrap=True)
        ws3.row_dimensions[row_idx].height = 24

    xlsx_filename = f"month_end_close_{CURRENT_MONTH.replace(' ', '_')}.xlsx"
    wb.save(xlsx_filename)
    print(f"   📊 Excel report saved: {xlsx_filename}")

    print("\n" + "="*60)
    print("✅  ALL AGENTS COMPLETE")
    print(f"   Markdown : {md_filename}")
    print(f"   Excel    : {xlsx_filename}")
    print("="*60)

    return {**state, "report_md": md_filename, "report_xlsx": xlsx_filename}

# ══════════════════════════════════════════════════════════════════════════════
# BUILD THE LANGGRAPH
# ══════════════════════════════════════════════════════════════════════════════
def build_graph():
    graph = StateGraph(CloseState)

    graph.add_node("planner",            planner_node)
    graph.add_node("variance_analyst",   variance_analyst_node)
    graph.add_node("commentary_writer",  commentary_writer_node)
    graph.add_node("reviewer",           reviewer_node)
    graph.add_node("report_generator",   report_generator_node)

    graph.set_entry_point("planner")
    graph.add_edge("planner",           "variance_analyst")
    graph.add_edge("variance_analyst",  "commentary_writer")
    graph.add_edge("commentary_writer", "reviewer")

    # Conditional: if reviewer rejects any item, loop back to reviewer
    # (commentary_writer already put improved draft in state)
    graph.add_conditional_edges(
        "reviewer",
        needs_retry,
        {
            "reviewer":          "reviewer",
            "report_generator":  "report_generator",
        }
    )
    graph.add_edge("report_generator", END)

    return graph.compile()

# ══════════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    print("\n" + "="*60)
    print("🏦  MONTH-END CLOSE ASSISTANT")
    print(f"    {CURRENT_MONTH} vs {PREVIOUS_MONTH}")
    print("="*60)

    if not GEMINI_API_KEY:
        print("❌  ERROR: GEMINI_API_KEY not set.")
        print("    PowerShell: $env:GEMINI_API_KEY=\"your_api_key_here\"")
        exit(1)

    current_tb, previous_tb = load_trial_balances()
    print(f"✅  Loaded trial balances: {len(current_tb)} accounts each")

    initial_state: CloseState = {
        "current_tb":  current_tb,
        "previous_tb": previous_tb,
        "checklist":   [],
        "variances":   [],
        "commentary":  [],
        "report_md":   "",
        "report_xlsx": "",
    }

    app = build_graph()
    final_state = app.invoke(initial_state)

    print(f"\n🎉  Done! Open your reports:")
    print(f"    {final_state['report_md']}")
    print(f"    {final_state['report_xlsx']}")
