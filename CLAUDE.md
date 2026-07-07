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
**Last updated: July 6, 2026 (session 11)**
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
- Tab 4: Drawing — standalone facility drawing tool (FacilityDrawing extracted from Site/IQF)
- Tab 5: Site / IQF — IQF Site Survey form (no drawing; shows attached drawing thumbnail + PDF upload)
- Tab 6: Sys Eng — 7-step system engineering tool with Design Advisor, Energy Calculator, BOM generator
- Tab 7: Mkt Intel — placeholder

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
- DRAIN/PIPE mutual exclusion: changeTool() clears drain mode; DRAIN button sets tool state directly to avoid self-clearing; clicking PIPE always returns to regular pipe mode
- PDF export (white background, print-ready): jsPDF CDN loaded; doExportPDF() renders drawing on white canvas with light-gray grid (#EEEEEE), equipment symbols white-fill/dark-border (fdSym pm=true), dark pipe labels (#1a1a1a), Hodge footer bar; PDF button is primary, PNG kept as secondary
- Auto-reducing fittings: PIPE_SIZE_ORDER size index ['3/4"'…'4"']; 'reducer' symbol (tapered inline symbol) added to FITTINGS palette; findSnapPoint() returns sourceEl for pipe endpoint snaps; pipe completion detects end-to-end size mismatch and auto-inserts labeled reducer (e.g. "2×1") at connection point; works on both single-segment and L-shape routing paths
- PDF footer: branch (from airiq_site_survey_draft.branch or wingman_branch localStorage) and customer location (from airiq_site_survey_draft.equipLocation) added; fields joined gracefully, blank fields omitted
- Help / SOP tab: ❓ tab added to bottom tab bar; HelpTab component with 9 searchable SOP sections (Getting In, Wingman, Field Log, Sys Eng, Site/IQF, Drawing Tool Quick Reference table, Market Intel, Sending Feedback, Change Log); dark theme, real-time keyword filter, inline section cards
- Drawing tool light mode: isLight state + MutationObserver on html.class; canvas BG, grid, ruler, toolbar, palette, BOM panel all theme-aware; fdSym receives pm=isLight for symbol fill/stroke
- Tank symbol: redrawn as bird's eye top-down circle (outer ring=body, inner ring=cap/fitting) with 4-directional connectors; FD_BASE_PX updated
- Pipe body snap + tee auto-insert: snap to pipe mid-segment (bt 0.05–0.95), auto-inserts tee or reducing tee at junction; tee rotation derived from main pipe angle (+90° for start, +270° for end); reducer auto-insert on size mismatch; bflip forces perpendicular-first routing so L-shape elbow lands off main pipe
- Pipe label suppression: segments < 2ft show no label (avoids clutter at tee junctions)
- Symbol text readability: tank inner ring replaced with small filled dot (r=3); compressor ring shrunk r=12→r=6 to clear 'COMP' text
- Pipe/drain label backgrounds: isLight-aware (light bg + navy text in light mode)
- NPT adapter label dedup: hasNptLabel() checks els within 60px — only first NPT at any equipment port gets a label, subsequent ones blank
- Flow direction arrows: chevron at pipe midpoint, mat color, 1.5px, shows flow from x1/y1 to x2/y2
- Tee stem direction fix: teeBranchRot() helper; body snap tees correctly orient stem toward branch for all 4 insertion points; perpendicular endpoint snaps also auto-insert tee (cross product > 0.5)
- PDF footer: branch + customer location auto-populated from localStorage (airiq_site_survey_draft, airiq_rep_branch, wingman_branch fallback chain)
- Dark/light mode toggle: ☀️/🌙 button in app header; anti-FOUC script; CSS variables on html.theme-dark/html.theme-light; all tabs and drawing tool UI respond to theme; canvas re-draws on toggle via MutationObserver + isLight state; PalIcon uses pm=isLight for correct symbol rendering on light palette
- Tank symbol top-down view: fdSym 'tank' case redrawn as bird's eye circle (outer r=22, inner r=6 cap ring); FD_CONNECTORS.tank updated to radius 22; FD_BASE_PX.tank updated to [44,44]
- Drain/Pipe mutual exclusion (visual): PIPE button active highlight suppressed when drainMode=true so reps see clear state; clicking PIPE still calls changeTool which sets drainMode=false
- PDF export footer: already had branch + location from previous session; added airiq_rep_branch localStorage key as additional fallback
- Pipe body snap (T-branch): findSnapPoint priority 2.5 — projects cursor onto pipe body (bt 5%–95%), returns {sourceEl, onBody:true}; pipe completion auto-inserts tee at body snap (reducing tee labeled 'X×Y Red. Tee' if sizes differ, regular 'X Tee' if same size); works for both straight and L-shape routing
- MY LOG export: format selector (xlsx/csv) added to rep export modal; server builds .xlsx via XLSX package or returns JSON for client-side CSV
- Team View export: same format selector; server builds .xlsx with MM/DD/YYYY dates; CSV path unchanged
- DB-backed location autocomplete in Field Log: GET /api/field-log/locations; fetch-once per session; keyboard nav; touch+mouse handlers (onTouchStart/onTouchEnd + onMouseDown); locInputFocusedRef fixes first-focus async race
- Stop search / form pre-fill: GET /api/field-log/search?q= (ILIKE across company/contact/location/notes); 400ms debounce; results panel with onMouseDown+onTouchEnd; fillFromSearchResult pre-fills form; ESC/× clear
- GET /api/stories + POST /api/stories endpoints added; STORIES_FILE on Railway volume; Wingman LogAWin was silently failing without these
- Drawing Tab Separation (session 10): FacilityDrawing extracted from SiteSurvey into standalone "Drawing" tab (id: drawing, icon: ✏️); App-level drawingAttachment and drawingForBOM state; Drawing toolbar: ⬇ DXF (DXF R12 export), ▶ IQF (sendToIQF callback), ▶ SYS ENG (sendToSysEng callback); SiteSurvey: shows attached drawing thumbnail with REMOVE + PDF upload zone; Sys Eng: "Analyze Drawing for BOM" card with AI analysis, editable BOM table, Copy BOM; Estimator "Send to Drawing" buttons now route to drawing tab; session loadSession() switches to drawing tab
- MY LOG stop card date prefix (session 11): formatStopStamp() shows time-only for today's stops and "Jul 6, 2026 · 11:56 AM" for prior days; applied to both MY LOG views (grouped + chronological); Team View + formatTime unchanged
- Sales Blitz Mode (session 11): standalone BlitzMode component inside Field Log tab. SheetJS (xlsx.full.min.js) added to CDN scripts (was NOT previously loaded). ⚡ BLITZ MODE header button visible to ALL reps when an active blitz exists, and always to blitz leaders (Jon S, Tony, Kyle, Morty, Hudson); button is orange when a blitz is active, navy-outlined otherwise; hidden entirely from non-leaders when no active blitz. Two sub-tabs: MY GROUP (group/day selectors built dynamically from the uploaded file, default group from localStorage blitzGroup + current blitz day by calendar diff; primary stops in numeric order, backups collapsed behind a toggle; stop cards show # badge orange/green, company, industry pill, contact+title, tel: phone, maps link — Apple Maps on iOS / Google Maps else, air-likelihood + site-confidence badges, priority score; outcome buttons + notes + LOG STOP; progress bar "X of Y stops logged today") and BLITZ DASHBOARD (scoreboard: one row per group, per-day stops/opps, total stops/opps, conv%; overall totals row; auto-refresh every 60s). Leaders-only: xlsx upload panel (reads "All 240 Stops" sheet or first sheet with a Group column, maps 17 columns) + End Blitz button (returns to normal Field Log). Blitz logging is fully separate from field_log_entries and never triggers monday.com.
- server.js Blitz endpoints (session 11): blitz_sessions + blitz_stops tables (initBlitzTables on startup); POST /api/blitz/upload (deactivates prior active session, inserts stops in a transaction, returns {session_id, totalStops}); GET /api/blitz/active; POST /api/blitz/log/:stopId; GET /api/blitz/dashboard (per group/day stops+opps + overall totals; opp = outcome not null and not 'nothing'); DELETE /api/blitz/session
- Reps Morty + Hudson added (session 11): REP_START_DATES (Morty 2010-01-01, Hudson 2025-04-14), Field Log rep dropdown, and global full-name REPS array. NO_GOAL_REPS = ['Morty','Hudson','Tony','Kyle','Jon S'] — getRepGoals() returns {daily:null, weekly:0, monthly:0} for these regardless of tenure/day

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
1. Drawing tool polish — auto-label fittings with AIRpipe part numbers + size
2. Live ST inventory lookup (replace daily Gmail report) — report ID 1823
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
