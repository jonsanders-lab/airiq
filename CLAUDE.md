# Jon Sanders — Global Claude Code Memory
# Place this file at: ~/.claude/CLAUDE.md
# This applies to ALL Claude Code sessions across every project.
# Last updated: 2026-06-01

---

## Who I Am

- **Jon Sanders**, VP of Sales at Hodge Compressors (subsidiary of Hodge Industrial Technologies, Hoschton, GA)
- I lead 2 Regional Sales Managers → ~16 field reps across 9 branches (Atlanta, Charlotte, Tampa, Greenville, Nashville, Dallas, Detroit, Cleveland, Chicago)
- I build AI-powered internal tools for the sales team alongside my day job — I am not a full-time developer, but I am technically hands-on and hold my tools to a high standard

---

## How I Work — Behavioral Rules (Karpathy Method)

### Rule 1: Verify Before Acting
- If a task is ambiguous, **ask one clarifying question** before writing a single line of code
- Never silently assume what I mean — a wrong assumption wastes more time than a question
- If you're about to make a decision that can't easily be undone, state it and confirm

### Rule 2: No Over-Engineering
- Build the simplest thing that works
- Prefer 50 lines over 500 lines
- No premature abstractions, no utility layers I didn't ask for
- If you think a more complex approach is genuinely better, say so briefly — but default to simple

### Rule 3: No Scope Creep
- Fix/build exactly what was asked
- Do NOT refactor adjacent code, rename unrelated variables, or "clean up while I'm in here"
- If you notice something worth fixing elsewhere, **mention it at the end** — don't touch it

### Rule 4: Data Accuracy is Non-Negotiable
- I work with real product specs, part numbers, model numbers, and pricing
- **Never fabricate, estimate, or infer technical specifications**
- If you don't have confirmed data, say so explicitly: "I don't have a verified spec for this"
- Always prefer data sourced from my Google Drive, uploaded files, or documents I provide

### Rule 5: Define Done
- Before starting any non-trivial task, confirm: "When this is complete, X will work / Y will display / Z will output correctly"
- This prevents building the wrong thing for 30 minutes

---

## My Tech Stack

| Tool | Purpose |
|------|---------|
| Node.js / Express | Backend (AirIQ, Wingman) |
| React (CDN, no build step) | Frontend |
| Railway | Hosting / deployment |
| monday.com | CRM |
| ServiceTitan | Business management / pricebook |
| Smartsheets | Sales reporting |
| QuickBooks | Accounting |
| Google Suite + Drive | Docs, specs, knowledge base |
| Slack | Team communication |

---

## Domain Knowledge — Compressed Air Systems

I have deep technical knowledge in this domain. When working on industry-related code or content:
- Rotary screw compressors (fixed speed and VSD)
- Refrigerated and desiccant dryers
- Compressed air piping systems (AIRpipe aluminum system)
- SCADA / compressed air auditing
- Competitive equipment: Sullair, Hertz, Atlas Copco

**Do not dumb down technical language.** Use correct industry terminology.

---

## My Active Projects (as of 2026-06-01)

| Project | Description | Repo / URL |
|---------|-------------|-----------|
| **AirIQ** | AI sales assistant suite (Tab 1: AI Q&A, Tab 2: Wingman field briefing, Tab 3: AIRpipe Piping Estimator) | jonsanders-lab/airiq · airiq-production.up.railway.app |
| **Hodge Knowledge Base** | Markdown-based LLM wiki for product specs, SOPs, confirmed part numbers | TBD — in planning |

---

## What I Care About

1. **Accuracy over speed** — I will catch errors; don't fabricate to fill a gap
2. **Rep-usable output** — tools must work for field reps with zero technical background
3. **"Parking lot mode"** — mobile-first, under 60 seconds to get useful output
4. **Margin protection** — quoting errors cost real money; anything touching pricing or labor must be exact

---

## Session Hygiene

- Use `/compact` when context gets long — don't let drift accumulate
- If a debugging loop goes more than 3 attempts without progress, **stop and reframe** — don't keep trying the same thing
- At the end of multi-step tasks, summarize what was done and what's left
