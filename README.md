A multi-agent AI pipeline that automates month-end financial close — reading trial balances, flagging variances, writing management commentary, reviewing it for quality, and generating a final Markdown + Excel report. Built with **LangGraph** and **Google Gemini**.

## What It Does

Upload two trial balance CSVs (current month and previous month) and the pipeline:

1. **Builds a close checklist** tailored to the actual accounts in your trial balance
2. **Flags all accounts** with movement greater than 10% vs. prior month
3. **Writes management commentary** for each variance in CFO-ready language
4. **Reviews the commentary** and rejects anything vague or incomplete — looping back until quality meets the bar
5. **Generates two output files**: a Markdown report and a color-coded Excel workbook

## The 5-Agent Pipeline

```
Planner → Variance Analyst → Commentary Writer → Reviewer → Report Generator
                                                      ↑____________|
                                               (loops back if rejected)
```

| Agent | Role | LLM? |
|---|---|---|
| 🗂️ Planner | Reads both trial balances, builds month-end checklist | ✅ Gemini |
| 📊 Variance Analyst | Flags accounts with >10% movement using pandas | ❌ Pure Python |
| ✍️ Commentary Writer | Drafts 2–3 sentence explanation per variance | ✅ Gemini |
| 🔍 Reviewer | Reviews each commentary; rejects if vague or incomplete | ✅ Gemini |
| 📄 Report Generator | Assembles Markdown + Excel output | ❌ Pure Python |

The **Reviewer ↔ Commentary Writer loop** is a conditional edge in LangGraph — if any commentary is rejected, the improved version cycles back through the Reviewer until approved (max 2 retries).

## Sample Output

Running on September 2026 vs August 2026 trial balances (24 accounts), the pipeline flagged **9 accounts** including:

| Account | Change | Direction |
|---|---|---|
| Marketing Expenses | +76.8% | ▲ |
| Cash and Bank | -38.2% | ▼ |
| Revenue - Services | +35.5% | ▲ |
| Retained Earnings | +19.5% | ▲ |
| Accounts Payable | +8.3% | ▲ |

### Excel Output (3 sheets)
1. **Variance Report** — All flagged accounts, color-coded green (increase) / red (decrease)
2. **Management Commentary** — AI-written, CFO-reviewed commentary per account
3. **Close Checklist** — Auto-generated checklist specific to your trial balance accounts

## Tech Stack

| Component | Technology |
|---|---|
| Agent Orchestration | LangGraph (`StateGraph`, conditional edges) |
| LLM | Google Gemini (`gemini-3.5-flash-lite`) via LangChain |
| Data Processing | pandas |
| Excel Output | openpyxl |
| Language | Python 3.10+ |

## LangGraph Concepts Used

- **StateGraph** — shared `CloseState` TypedDict passed between all agents
- **Nodes** — each agent is a node function that reads and updates state
- **Edges** — linear flow from Planner → Variance Analyst → Commentary Writer → Reviewer
- **Conditional Edges** — `needs_retry()` routes back to Reviewer if any commentary is rejected
- **Planner–Worker–Reviewer pattern** — core multi-agent design pattern for quality control loops

## How to Run

### Prerequisites
```bash
pip install langgraph langchain-google-genai pandas openpyxl
```

### Setup

1. Get a Gemini API key from [Google AI Studio](https://aistudio.google.com/)
2. Set the environment variable:
   ```powershell
   # PowerShell (Windows)
   $env:GEMINI_API_KEY="your_api_key_here"
   ```
   ```bash
   # Mac / Linux
   export GEMINI_API_KEY="your_api_key_here"
   ```

3. Place your trial balance CSVs in the same folder as the script:
   - `trial_balance_sep2026.csv` — current month
   - `trial_balance_aug2026.csv` — previous month

### Run
```bash
python month_end_close_assistant.py
```

The terminal shows real-time progress as each agent completes. Total runtime: ~2–4 minutes (18 Gemini API calls across Agents 1, 3, and 4).

## CSV Format

Both trial balance files must have these columns:

```
account_code, account_name, category, debit, credit, net_balance
```

`category` should be one of: `Asset`, `Liability`, `Equity`, `Income`, `Expense`

Sample data files are included to test the pipeline out of the box.

## Configuration

Edit these constants at the top of `month_end_close_assistant.py` to adapt to your data:

```python
CURRENT_TB_FILE    = "trial_balance_sep2026.csv"   # your current month file
PREVIOUS_TB_FILE   = "trial_balance_aug2026.csv"   # your previous month file
CURRENT_MONTH      = "September 2026"
PREVIOUS_MONTH     = "August 2026"
VARIANCE_THRESHOLD = 10.0   # flag accounts with >10% movement
```

## Author

**Ravi Sehgal** | Chartered Accountant & Finance Systems PM  
[LinkedIn](https://www.linkedin.com/in/ravisehgal) · [GitHub](https://github.com/Ravi4Sehgal)
