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
**Last updated: June 16, 2026 (session 5)**
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
- Tab 4: Site / IQF — IQF Site Survey form + Facility Drawing Tool
- Tab 5: Sys Eng — 7-step system engineering tool with Design Advisor, Energy Calculator, BOM generator
- Tab 6: Mkt Intel — placeholder

## KEY COMPLETED FEATURES
- ST pricebook pricing via Cloudflare Worker proxy
- Gmail inventory pull daily at 6:10am EST — 73 equipment items cached
- Railway persistent volume at /app/data — estimates.json and stories survive deploys
- Google Sheets integration — Site Survey saves to "IQF - AirIQ Site Surveys" (Sheet ID: 1Zrzr_PgEMmMmclJ2W1Dere-JEKYmaB4xa-N4MMUIUuI)
- Gmail OAuth refresh token covers both Gmail + Sheets scopes
- Voice input model number normalization (HD 30 → HD30)
- Series stock queries render as clean HTML tables (HTM stock, HD stock, HB stock etc) — bypass Claude, hit inventory directly
- Single model ST lookup renders as clean styled card with variant table (formatSTReply)
- Multi-model parallel ST lookup with 6s timeout and proper tool_use/tool_result handling for all tool blocks
- Pressure variant stocking rules in system prompt (HD/HT/HTV/HV stock 116psi only; HB stocks 116/145/181; HTM stocks 145psi only)
- HB belt drive equivalent output hallucination fixed in system prompt
- Tab order updated: AirIQ → Wingman → Field Log → Site/IQF → Sys Eng → Mkt Intel
- CLAUDE.md auto-update rule added (runs end of every session)
- BETA badge (neon green #39FF14, CSS glow) added to Site/IQF and Sys Eng tab headers
- FeedbackModal component: rep name (prefilled from localStorage), branch dropdown, type selector (Bug/Feature/General), textarea → POST /api/slack-feedback
- server.js /api/slack-feedback endpoint: proxies to SLACK_WEBHOOK_URL env var, formats structured Slack message
- Canvas leader lines: copper line + dot from symbol bottom to label, only renders when gap > 20px screen (high zoom or large scale)
- Pipe stick length: 19 ft (corrected from 13.12 ft)
- Lugging tool auto-check when pipe size >= 2.5"
- Facility Drawing Tool: fullscreen, zoom/pan, rotation, snap, auto-fittings, auto-couplings at 19ft, elevation notes, BOM from drawing
- Energy & Cost Calculator with VSD recommendation engine and tank cycle calculator
- Site Survey: 10-section IQF form, auto-save draft, Past Surveys panel, email on save (Gmail compose)
- 💬 FEEDBACK button (no BETA badge) added to AirIQ, Wingman, Field Log, and Mkt Intel tab headers — all post to /api/slack-feedback with correct tab name
- Equipment real-world dimensions: FD_DIMS/FD_TANK_DIMS lookup table (HD/HV/HT/HTM compressors, HAD dryers, receiver tanks by gallon) auto-scales compressor/dryer/tank symbols to true footprint on the drawing grid; editable W/D (inches) panel appears when one of those symbols is selected; auto-populates on Sys Eng pre-place and on manual placement when the typed label matches a known model
- FD_CONNECTORS given 4-directional (not just left/right) connector points for compressor/dryer/tank/filter/ows so pipes approaching from any side get a precise snap — fixes inconsistent NPT-adapter generation and tank diagonal-pipe artifacts
- Ortho-lock on pipe snap: if a pipe endpoint snaps within 1 grid cell on one axis, the pipe is forced perfectly horizontal/vertical instead of diagonal
- Right-click (PC) now opens the same symbol context menu (Edit Label/Rotate/Duplicate/Delete) as long-press on iPad — `onContextMenu` handler added to the drawing canvas
- 8-material pipe selector in drawing tool: AIRpipe group (Blue/Green/Gray) + Other Materials (Copper, Stainless, PVC, CPVC, Black Iron); stainless renders shimmer gradient, PVC/CPVC show ⚠️ warning borders, black iron double-strokes for visibility on dark canvas
- Field Log: GRAPH VIEW — bar chart toggle (LIST|GRAPH) in leader view with pace-colored bars, 20-stop dashed goal line, rep name labels; secondary weekly trend SVG mini-chart (Mon–today) with area fill, data points, day labels
- Field Log: CSV EXPORT (leader only) — EXPORT button opens date-range modal (Today/Week/Month/Year/Custom) → auto-downloads CSV with all 13 activity columns
- Field Log: LEADER AUTO-DASHBOARD — LEADERS const (Jon Sanders, Tony, Kyle); view state auto-inits to leader view when rep name matches
- server.js: GET /api/field-log/week (Mon–today team stop totals by day) and GET /api/field-log/export (per-rep-per-day aggregate, all columns, date-range params)
- Condensate drain lines in drawing tool: DRAIN toggle in toolbar, drain size selector (1/4"–10mm metric), dashed amber (#F59E0B) rendering with endpoint dots, orthogonal-only routing, no auto-NPT/coupling
- DRAIN FITTINGS palette section: PTL Elbow, PTL Tee, PTL Straight — amber symbols that inline-snap to drain pipes only (PTL_FITTINGS filter in findPipeSnap)
- BOM: split into COMPRESSED AIR and CONDENSATE DRAIN sections; drain rows show footage and PTL fitting counts
- Infinite canvas: drawing area is unbounded in all directions; grid renders dynamically based on viewport (wLeft/wRight/wTop/wBottom computed from zoom/pan); only visible elements rendered (viewport culling in element + leader-lines loops)
- Dynamic grid interval: auto-scales from 1ft (high zoom) to 1000ft (low zoom) so screen spacing stays >= 16px; ruler labels scale independently to >= 80px
- fitView: content-aware — computes bounding box of all drawn elements and fits to canvas with padding; falls back to origin if empty
- Zoom range: 10%–400% (clampZoom changed from 0.25 min to 0.1)
- Scale bar: bottom-right screen-space overlay shows current real-world scale (e.g. "├──┤ 50 ft"), updates live with zoom
- Ortho pipe routing: all air pipe draws route as two ortho segments (H-V default, Shift key flips to V-H); live L-shape ghost preview with corner dot and mode label; auto-elbow (elbow90 at correct orientation) placed at corner; NPT adapters and couplings distributed across both segments; existing diagonal pipes in saved drawings auto-converted to H-V on load via straightenPipes(); status bar shows SHIFT hint during drawing; drain lines unchanged

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
- Equipment footprints (W x D, inches) for drawing tool auto-scale: HD30/HV30 28x46, HD60/HV60/HT30/HT40 30x55, HD100 34x68, HD150 38x76, HT200 48x90, HT350 60x108, HTM60/75 48x72, HTM150 60x90; HAD dryers 8x12 (HAD18) up to 20x32 (HAD487); receiver tanks 80G=16x16 up to 660G=36x36 — see FD_DIMS/FD_TANK_DIMS in index.html

## AIRPIPE PART NUMBERS
Pipe prefix by size: 3/4"=1, 1"=2, 1.5"=4, 2"=5, 2.5"=6, 3"=7, 4"=8, 6"=9, 8"=A
Blue=x000, Gray=x062, Green=x061
Union=x002, 90 Elbow=x003, Equal Tee=x005, End Cap=x006, Flex Hose=x055

## NEXT PRIORITIES (in order)
1. Drawing workflow — Site/IQF tab=as-is facility layout, Sys Eng tab=proposed system with auto-placed equipment
2. Drawing tool polish — auto-label fittings with AIRpipe part numbers + size
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
