# Engineering Agent Template (OpenCode)

Universal starter template for engineering projects built with AI agents (OpenCode / Claude Code), aligned with the **"Agenci AI w pracy inżyniera"** series.

## Project Overview
- **Project Name:** [Fill during setup]
- **Domain:** [e.g. Hardware / Embedded / Mechanical / Automation]
- **Key Tools:** [e.g. Python, Docling, LTspice, KiCad]

---

### OPERATIONAL PROTOCOL
Every task MUST follow this exact 3-step cycle:

1. **START (Before touching any project code):**
   - Read & commit goal: update `.context/active-context.md` setting:
     - `Goal: <clear summary of user request in English>`
     - `Status: In Progress`
     - `Next Steps: <numbered steps in English>`
   *CRITICAL: Do NOT edit or create any project files until `.context/active-context.md` has `Status: In Progress`.*

2. **WORK (Implementation):**
   - Implement the planned changes.
   - Run tests and verify deliverables on disk.

3. **END (Before returning your final response):**
   - Close context: update `.context/active-context.md`:
     - Append completed task to `## Progress Log` in English (if > 5 entries, keep 5 newest and summarize older into one paragraph in the same write).
     - Clear `Goal` and `Next Steps`.
     - Set `Status: Completed`.

**Routine / repetitive tasks — fast path:**
A task is *routine* when it is a known, documented, periodic operation whose normal outcome is "done / N items" and generates no new knowledge — e.g. the daily Outlook email sync (`sync_outlook_emails.py`), running a documented script, a periodic staleness check. For routine tasks: skip the START goal churn and the END Progress Log entry — UNLESS something new or unexpected happened (a bug, a gotcha, a decision, or a structural change).

---

### ROUTINE TASKS & COMMANDS
- **Outlook Email Sync:** Configure `scripts/sync_outlook_config.json` (via `2_prompt_mail.md`), then run:
  `python scripts/sync_outlook_emails.py`
- **Run Tests:**
  `python -m unittest discover -s tests`

---

### CONTEXT INTEGRITY
- Keep `.context/active-context.md` headers, fields, and statuses strictly in English (`In Progress`, `Completed`, `None`).
- Never hardcode personal emails, credentials, or confidential project data in `scripts/sync_outlook_emails.py`; store project-specific sync settings in `scripts/sync_outlook_config.json` (gitignored).
