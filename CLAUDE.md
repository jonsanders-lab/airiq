## AUTO-UPDATE RULE (Claude Code must follow this)
At the end of EVERY session, before the final git push:
1. Update the "Last updated" date at the top of this file
2. Move any completed items from NEXT PRIORITIES to KEY COMPLETED FEATURES
3. Add any new features built this session to KEY COMPLETED FEATURES
4. Add any new business rules discovered to KEY BUSINESS RULES
5. Update NEXT PRIORITIES to reflect current state
6. Commit CLAUDE.md in the same final commit as the session work

This keeps the briefing current automatically. Never skip this step.

---

# AirIQ — Claude Code Project Briefing
**Last updated: June 15, 2026**
**VP of Sales: Jon Sanders — Hodge Industrial Technologies, Hoschton GA**
**9 branches: Atlanta, Charlotte, Tampa, Greenville, Nashville, Dallas, Detroit, Cleveland, Chicago**
**16 reps across 2 RSMs**

## PROJECT: AirIQ + Wingman
React/Node.js PWA deployed on Railway at airiq-production.up.railway.app
GitHub: jonsanders-lab/airiq
Local path: C:\Users\jonsa\Documents\airiq
Railway volume mounted at /app/data for persistent storage

## CURRENT TAB STRUCTURE
- Tab 1: AirIQ — AI sales assistant with live ST pricing and inventory
- Tab 2: Wingman — Mission Brief, Field Intel, Log a Win
- Tab 3: Field Log — daily stop tracker
- Tab 4: Sys Eng — 7-step system engineering tool with Design Advisor, Energy Calculator, BOM generator
- Tab 5: Site Design — IQF Site Survey form + Facility Drawing Tool
- Tab 6: Mkt Intel — placeholder

## KEY COMPLETED FEATURES
- ST pricebook pricing via Cloudflare Worker proxy
- Gmail inventory pull daily at 6:10am EST — 73 equipment items cached
- Railway persistent volume at /app/data — estimates.json and stories survive deploys
- Google Sheets integration — Site Survey saves to "IQF - AirIQ Site Surveys" (Sheet ID: 1Zrzr_PgEMmMmclJ2W1Dere-JEKYmaB4xa-N4MMUIUuI)
- Gmail OAuth refresh token covers both Gmail + Sheets scopes
- Voice input model number normalization (HD 30 → HD30)
- Series stock queries (HTM stock, HD stock) bypass Claude and hit inventory directly
- Pipe stick length: 19 ft (corrected from 13.12 ft)
- Lugging tool auto-check when pipe size >= 2.5"
- Facility Drawing Tool: fullscreen, zoom/pan, rotation, snap, auto-fittings, auto-couplings at 19ft, elevation notes, BOM from drawing
- Energy & Cost Calculator with VSD recommendation engine and tank cycle calculator
- Site Survey: 10-section IQF form, auto-save draft, Past Surveys panel, email on save (Gmail compose)

## KEY BUSINESS RULES (hardcoded)
- HTM series: 2-3 day assembly lead time even if ST shows stock
- Standard orders: 50% deposit. Hot rush: 100% upfront
- HTM can run on 208V despite tech data showing 230-460V
- All pricing from ServiceTitan only — never from vendor price books
- Pipe sticks: 19 ft. Auto-add coupling every 19 ft on pipe runs
- Lugging tool rental required for pipe 2.5" and larger
- Minimum labor: first compressor 8hrs, each additional 4hrs, each dryer 4hrs, each tank 4hrs
- Energy cost formula: motor_kW × load_factor × annual_hours × $0.085/kWh
- VSD uses load_factor^1.6 (square-law). Fixed speed unload draws 25%

## AIRPIPE PART NUMBERS
Pipe prefix by size: 3/4"=1, 1"=2, 1.5"=4, 2"=5, 2.5"=6, 3"=7, 4"=8, 6"=9, 8"=A
Blue=x000, Gray=x062, Green=x061
Union=x002, 90 Elbow=x003, Equal Tee=x005, End Cap=x006, Flex Hose=x055

## NEXT PRIORITIES (in order)
1. Drawing tool polish — auto-label fittings with AIRpipe part numbers + size
2. Drawing workflow — Site Design=as-is layout, Sys Eng=proposed system with auto-placed equipment
3. Live ST inventory lookup (replace daily Gmail report) — report ID 1823
4. Lead time display for out-of-stock units + Trello board link for ETA
5. ST customer lookup + autofill contact and address
6. Drive time calculation from branch to customer
7. Electrical sizing calculator
8. System Design Interview guided Q&A → Good/Better/Best
9. SOW generator
10. Sales presentation builder
11. Referral list in Wingman tab

## LONG TERM
- VR/AR headset app — rep places equipment and piping in augmented reality on-site for 99% accurate BOM
- Monday.com deal status integration
- All drawings Hodge-branded, no supplier names

## BEHAVIORAL RULES FOR CLAUDE CODE
- Always verify before editing — read the file section before changing it
- No scope creep — only change what was asked
- Pipe sticks are 19 ft — never revert to 13.12
- Never fabricate prices — all pricing from ST only
- After every build: git add, commit with descriptive message, push to main
- Working branch: main (feature branches merged)
- Stable baseline tag: v-stable-pre-systemdesign
- If index.html exceeds context, ask Jon to paste the relevant section rather than guessing
- Windows machine — use PowerShell syntax, never bash/Mac paths
