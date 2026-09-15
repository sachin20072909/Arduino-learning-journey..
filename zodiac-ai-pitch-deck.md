# ZODIAC AI — Pitch Deck & Technical Overview

**Deck:** 11 slides · ~20-minute presentation + Q&A
**Theme:** Dark-mode "Cyber-Grid" — cyberpunk terminal interfaces, neon green + electric blue accents, glowing data streams
**Tone:** Highly technical, professional, forward-thinking

---

## 🎨 GLOBAL DESIGN SYSTEM (apply to all slides)

| Token | Value | Usage |
|---|---|---|
| Background | `#0A0E14` (near-black) | Base canvas on every slide |
| Neon Green | `#00FF41` | Terminal text, primary accents, "success" states |
| Electric Blue | `#00D4FF` | Data streams, diagram traces, secondary accents |
| Grid | 1px cyan lines @ 8% opacity, one-point perspective floor | Background depth layer |
| Glow | 8–16px outer glow on all neon elements | Signature "live data" feel |
| Typography | Monospace (JetBrains Mono / Fira Code) for code & labels; geometric sans (Space Grotesk / Orbitron) for headlines | Terminal-vs-human contrast |
| Motifs | Scanlines, blinking cursors, status bars (`> ask anything_`), matrix-style data cascades, holographic wireframes | Recurring visual vocabulary |

**Accent logic:** Green = *ground truth / executed / safe*. Blue = *in-flight / routed / rendered*. Red = *blocked / rejected / hallucinated*.

---
---

## SLIDE 1 — TITLE

**Slide Title:** ZODIAC AI — Chat with Your Database

**Visual Direction:**
Full-bleed near-black scene. Center-right: a holographic wireframe database cylinder pulsing neon green, with particle data-streams arcing out of it into a floating translucent chat terminal. Behind everything, an electric-blue cyber grid floor receding to a vanishing point. Faint overhead overlay: the 12-point Zodiac constellation connected by thin neon lines. The word **ZODIAC** set in glitch-styled display type with subtle chromatic aberration; a terminal status bar along the bottom reads `v1.0 // TEXT-TO-SQL AGENT // STATUS: ONLINE`.

**Main Content:**
- **ZODIAC AI** — *Chat with Your Database*
- A conversational agent that talks directly to relational databases — and answers with **exact SQL, interactive Plotly charts, live Mermaid diagrams, and mathematically grounded explanations.**
- Tagline (glowing, bottom third): **NOT GUESSES. GROUND TRUTH.**
- Footer chips: `TEXT-TO-SQL` · `AGENTIC LOOP` · `MULTI-PROVIDER` · `8-STEP REASONING`

**Speaker Notes:**
Open with the one-liner: "Every company runs on data — almost nobody can talk to it directly." Introduce Zodiac AI as a conversational agent that sits between humans and the relational database: you ask in plain English, and it returns the exact SQL, the chart, the diagram, and the numbers — every figure traceable to an executed query. Set the hook fast and promise the architecture tour to come.

---
---

## SLIDE 2 — EXECUTIVE VISION

**Slide Title:** The Wall Between You and Your Data

**Visual Direction:**
Split-screen composition divided by a vertical, pulsing electric-blue lightning seam. **Left half (the old way):** desaturated grey-blue — a cluttered mosaic of stale BI dashboards, CSV exports, and a support-ticket queue icon, glitching slightly, colors drained. **Right half (the Zodiac way):** vivid neon — a single clean chat exchange: *"What drove revenue last quarter?"* → a response card containing a glowing SQL snippet, a mini interactive chart, and a bullet explanation. The contrast should feel like turning the lights on.

**Main Content:**

**The Problem — Traditional BI**
- Every question becomes a ticket: analysts queue, answers ship days later
- Static dashboards answer *yesterday's* questions, not today's
- SQL fluency is the bottleneck — the data exists; direct access doesn't
- Naive LLM chatbots hallucinate numbers with total confidence

**The Solution — Zodiac AI**
- Ask in plain English → receive the **exact SQL**, executed live against the database
- Answers arrive as **interactive Plotly charts** and **live Mermaid diagrams**
- Every number is **computed from real result sets** — mathematically grounded
- Analytics at the speed of conversation

**Speaker Notes:**
Frame it as economics: the real cost of BI is the *latency between question and answer*. When that latency is days, people stop asking. Generic AI chatbots feel instant but fabricate — speed without trust. Zodiac collapses the loop to seconds while keeping every number traceable to an executed query. Position it carefully: "We're not replacing analysts — we're unblocking everyone else."

---
---

## SLIDE 3 — TECH STACK

**Slide Title:** The Stack — Engineered for Speed, Built to Swap

**Visual Direction:**
A vertical layered-architecture diagram rendered as a neon circuit-board cross-section, like a chip die-shot. Four glowing horizontal strata, top to bottom: `FRONTEND` (React atom hologram + lightning bolt), `API LAYER` (rocket glyph), `AI LAYER` (two neural chips labeled GROQ and GEMINI connected through a small router switch), `DATA LAYER` (green SQLite cylinder). Luminous data beams travel vertically through glowing vias between layers. Each stratum carries a monospace spec label on its right edge.

**Main Content:**
- **Frontend — React + TanStack Query:** declarative UI, async server-state management, smart caching & background refetching for live results
- **Backend — FastAPI (Python):** async endpoints, Pydantic-validated payloads, native streaming for agent responses
- **Database — SQLite:** zero-config, file-based, instantly portable — a schema that ships with the repo
- **AI Layer — Groq + Gemini:** dual LLM providers hidden behind a single router interface
- **Design principle:** every layer is modular — swap any component (Postgres for SQLite, another provider for Groq) without rewriting the agent

**Speaker Notes:**
Two reasons for these choices: iteration velocity and architectural portability. React + TanStack Query elegantly handles the chat-plus-live-chart rendering pattern. FastAPI gives async Python with strict payload validation. SQLite makes the demo infinitely portable — no server to stand up. Most importantly, the AI layer is deliberately abstracted: the agent never talks to a vendor directly — it talks to the router.

---
---

## SLIDE 4 — MULTI-PROVIDER AI ARCHITECTURE

**Slide Title:** Two Brains, One Loop — Never Offline

**Visual Direction:**
Two-panel composition. **Left panel — the fallback router:** a central glowing node labeled `PROVIDER ROUTER` with two outgoing neural paths; the primary path burns bright neon green labeled `GROQ — ULTRA-LOW LATENCY`, shown mid-failover with a red glitch-sever on the green path and traffic rerouting onto the electric-blue `GEMINI` path. **Right panel — the agent loop:** an orbital ring of 8 nodes connected by a directional energy pulse, orbiting a core labeled `AGENT CONTEXT`; the segment from node 4→7 glows brighter, labeled `ITERATE UNTIL GROUNDED`.

**Main Content:**

**Provider-agnostic fallback router**
- One unified LLM interface — the agent is provider-agnostic by design
- **Primary: Groq** — near-instant inference on Llama-class models
- **Automatic failover** — rate limits, timeouts, outages seamlessly reroute to **Gemini**
- Zero single points of AI failure — run with one key or both

**The 8-step agent loop**
1. **Ingest** — parse the prompt + conversation context
2. **Introspect** — pull the live schema via `get_schema`
3. **Plan** — decompose the question, select tools
4. **Act** — execute the next tool call
5. **Observe** — inspect results and errors
6. **Validate** — is the answer complete and grounded?
7. **Iterate** — refine and loop (3 → 6) until it is
8. **Synthesize** — ship SQL + chart + diagram + explanation

**Speaker Notes:**
Two ideas on this slide. First, resilience: LLM providers are utilities, and utilities go down. The router treats Groq and Gemini as interchangeable behind one interface — when Groq rate-limits mid-session, the user never notices. Second, the loop: this is *not* one-shot prompt → answer. The agent reasons, acts, observes, and iterates — if the first query errors or returns partial data, it self-corrects and runs again. Answers are **earned through iteration**, not guessed on the first pass.

---
---

## SLIDE 5 — CORE AGENT TOOLS · PART 1

**Slide Title:** `get_schema` & `execute_query` — Read the Schema, Run the SQL

**Visual Direction:**
Two floating terminal windows in dark space. **Left terminal** (header: `get_schema`): a glowing tree view — `sales_db → customers / products / orders / order_items` — with column names and types cascading in green monospace. **Right terminal** (header: `execute_query`): a `SELECT ... JOIN ...` statement executing with a neon shield icon and a green `READ-ONLY ✓` badge; beneath it, a red-stamped rejected line: `DROP TABLE orders;` → `⚠ BLOCKED — WRITE GUARD`. Query results stream below like a matrix cascade.

**Main Content:**
- **`get_schema`** — the agent introspects the *live* database at runtime: tables, columns, types, and relationships. No hardcoded, stale schema assumptions
- **`execute_query`** — generates and runs **exact SQL** against the real database
- **Read-only by design** — SELECT-only enforcement; writes and schema mutations rejected at the guard layer
- **Radical transparency** — every answer ships with the SQL that produced it: copy it, audit it, rerun it

**Speaker Notes:**
Trust is the product, and these two tools are its foundation. Because the agent reads the schema fresh at runtime, it cannot reference a column that doesn't exist. Because it executes real SQL, answers come from the database — not from the model's parametric memory. The read-only guard means the worst-case interaction is an empty result set. Safety plus auditability: the user always sees the exact SQL behind every number.

---
---

## SLIDE 6 — CORE AGENT TOOLS · PART 2

**Slide Title:** `generate_chart` & `generate_flowchart` — Answers You Can See

**Visual Direction:**
A split dashboard inside a futuristic browser chrome. **Left:** an interactive Plotly chart mid-render — glowing bar/line series with neon-gradient fills and a hover tooltip caught on a data point. **Right:** a Mermaid relationship diagram rendered as a neon blueprint — nodes and arrows in electric blue with green highlight nodes. Behind both panels, faint source code (`fig = px.bar(...)` / `graph TD;`) streams vertically like a data waterfall and *morphs into* the visuals — making the code-to-render transformation visible.

**Main Content:**
- **`generate_chart`** — the agent authors **Plotly chart specs** from real query results; the browser renders them fully interactive
  - Zoom, hover, pan, filter — exploration, not a static screenshot
  - Bar, line, scatter, pie — chart type matched to the shape of the question
- **`generate_flowchart`** — the agent writes **Mermaid syntax**; diagrams render live in the chat
  - Schema maps, entity relationships, logic and pipeline flows
- One prompt can return both: **the numbers *and* the picture of the numbers**

**Speaker Notes:**
Text-only answers waste an LLM's real strength: structured generation. Zodiac treats visuals as first-class outputs — the model doesn't *describe* a chart, it *writes the chart definition*, and Plotly and Mermaid render it deterministically. Because the chart spec is generated from executed query results, the visual carries the same grounding as the SQL behind it. Deterministic rendering is also what keeps the visuals hallucination-free.

---
---

## SLIDE 7 — CORE AGENT TOOLS · PART 3

**Slide Title:** `explain_data` — Grounded, Not Guessed

**Visual Direction:**
Center: a neon magnifying lens hovering over a glowing result-set table; specific cells illuminate green and lift out of the grid, flowing upward into a paragraph of text where the same numbers appear highlighted — literally visualizing *"numbers traced to rows."* Orbiting the lens: math glyphs (∑, μ, x̄, Δ%) in faint blue. In the lower corner, a ghostly translucent chatbot bubble showing a made-up statistic is slashed through in red — the anti-hallucination motif.

**Main Content:**
- **`explain_data`** — computes numeric summaries **directly from the executed result set**: totals, averages, min/max, distributions, shares
- Every figure in the prose **traces back to an actual query result** — text is generated from computed numbers, never from model memory
- Explanations include the math: **growth rates, proportions, outliers** — derived first, then narrated
- **Anti-hallucination by architecture:** the model can only narrate what the database has proven

**Speaker Notes:**
This slide separates Zodiac from a wrapped chat API. Generic LLMs interpolate plausible-sounding statistics; Zodiac *computes* the statistics first and writes the narrative around them. The pipeline is strict: **query → compute → narrate.** If a number can't be traced to a result set, it doesn't make it into the answer. Say it plainly: "We don't ask the model to remember your revenue. We ask your database."

---
---

## SLIDE 8 — DATABASE ARCHITECTURE

**Slide Title:** Under the Hood — The Sample Sales Schema

**Visual Direction:**
An entity-relationship diagram reimagined as a **neon circuit schematic on a dark PCB**. Four table-nodes rendered as glowing chips — `CUSTOMERS` (green), `PRODUCTS` (blue), `ORDERS` (green), `ORDER_ITEMS` (blue) — connected by luminous traces with pulsing cardinality glyphs (`1 — ∞`) at each junction. Each chip exposes pin-labels for its key columns (`customer_id`, `order_id`, `product_id`, `quantity`, `price`). Faint grid and ruler ticks frame the schematic like engineering blueprints.

**Main Content:**
- A normalized e-commerce sales model — **four tables, one star:**
  - **`customers`** — who is buying
  - **`products`** — what is being sold
  - **`orders`** — the transaction header (who, when, status)
  - **`order_items`** — line-level detail (what, how many, at what price)
- Relationships: `customers 1—∞ orders` · `orders 1—∞ order_items` · `products 1—∞ order_items`
- **Join-rich by design** — the schema that stresses every text-to-SQL edge case
- The agent maps all of it at runtime via `get_schema` — nothing is hardcoded

**Speaker Notes:**
The demo dataset is deliberately canonical e-commerce: instantly understandable to any audience, and join-heavy — which is exactly where naive text-to-SQL implementations break. Every interesting question (top customers, revenue per product, average order value) crosses at least one join. The argument: if the agent is correct on this schema, it generalizes. And remember — the schema is never baked into prompts; it's discovered live.

---
---

## SLIDE 9 — VALID QUERY CATEGORIES & SAMPLE PROMPTS

**Slide Title:** Ask It Anything (Structured)

**Visual Direction:**
A dark-mode chat UI mockup with a vertical stack of user message bubbles in neon-green outline, each paired with a small electric-blue category chip (`AGGREGATION`, `RANKING`, `TREND`, `JOIN`, `SCHEMA`, `QUALITY`). One exchange is expanded to show the full multi-part response: a SQL code block, a mini Plotly chart, and an explanation paragraph — the "answer bundle." At the bottom, an empty input field with a blinking cursor invites the next question: `> ask anything_`

**Main Content:**
- **Aggregations** — *"What was total revenue per month this year?"*
- **Rankings** — *"Show me the top 10 customers by lifetime value."*
- **Trends & comparisons** — *"Plot quarterly sales growth as a line chart."*
- **Multi-table joins** — *"Which products are most frequently bought together?"*
- **Schema exploration** — *"Diagram how my tables relate to each other."*
- **Data quality** — *"Are there any orders with no line items?"*
- Every answer returns as a bundle: **exact SQL + interactive chart + grounded explanation**

**Speaker Notes:**
Walk one or two bubbles aloud, then let the audience read the rest. Point out the data-quality category last — it proves the agent can *reason about* the database, not just retrieve from it. Close on the bundle: a single natural-language prompt yields the SQL, a rendered chart, and a grounded explanation. That triple — query, visual, narrative — *is* the product experience.

---
---

## SLIDE 10 — GETTING STARTED

**Slide Title:** Zero to Query in Five Minutes

**Visual Direction:**
A full terminal window styled as a **boot sequence**: command lines appearing with a typing cursor, each executed line flashing green with a `✓`, and a progress bar motif `[████████░░] LOADING` mid-frame. The final command launches the app — a small browser-window thumbnail pops out of the terminal showing the Zodiac chat UI. At right, a `.env` file rendered as a glowing document with masked keys (`GROQ_API_KEY=••••••••`, `GEMINI_API_KEY=••••••••`).

**Main Content:**
- **Prerequisites:** Python 3.10+ · Node.js 18+ · Git — and at least one API key
- **Clone & enter:**
  - `git clone https://github.com/<org>/zodiac-ai.git && cd zodiac-ai`
- **Configure environment (`.env`):**
  - `GROQ_API_KEY=...` — primary provider
  - `GEMINI_API_KEY=...` — optional fallback (resilience if you want it)
- **Install & run:**
  - Backend: `pip install -r requirements.txt` → launch the FastAPI server
  - Frontend: `npm install` → `npm run dev`
- Open the local URL → start asking

**Speaker Notes:**
The pitch of this slide is *friction* — or the absence of it. SQLite means there is no database server to stand up: the schema ships inside the repo. One API key is enough to run; the second is optional insurance that activates the failover router. From `git clone` to first grounded answer: under five minutes. Invite the audience to do it live during Q&A.

---
---

## SLIDE 11 — CONCLUSION & Q&A

**Slide Title:** The Database Is Now a Conversation

**Visual Direction:**
A wide cinematic shot of the **Zodiac constellation** — 12 stars connected by neon lines — reflected on the cyber-grid floor, closing the visual loop with Slide 1. Data streams flow from the constellation down into a small, glowing database cylinder. A terminal beneath reads `$ status: ALL SYSTEMS OPERATIONAL` with a blinking `> ask anything_`. Oversized **Q&A** set in glitch typography, green with a blue chromatic offset.

**Main Content:**
- **What we built:** a conversational agent that answers with exact SQL, interactive Plotly charts, live Mermaid diagrams, and mathematically grounded explanations
- **Why it matters:** analytics at conversational speed — **without the hallucination tax**
- **How it holds up:** read-only guardrails · runtime schema awareness · multi-provider failover · iterative 8-step reasoning
- **What's next:** broader connectors (Postgres, MySQL) · richer chart grammar · conversational follow-ups · shared query memory
- **CTA:** Clone it. Query it. Break it. → **Q&A**

**Speaker Notes:**
Close the loop with the opening line: databases used to be things you *queried* — now they're things you *talk to*. Recap the four pillars in one breath: exact SQL, live visuals, grounded numbers, resilient architecture. Then open the floor — and if the room is quiet, seed the first question yourself: "Ask me what happens when Groq goes down mid-query."

---
---

## ⏱️ SUGGESTED PACING

| Slides | Time | Mode |
|---|---|---|
| 1–2 | ~3 min | Vision & hook |
| 3–4 | ~4 min | Architecture deep-dive |
| 5–7 | ~6 min | Tool-by-tool walkthrough (the differentiators) |
| 8–9 | ~4 min | Live schema & demo prompts (run live if possible) |
| 10–11 | ~3 min | Adoption path & close |
| Q&A | remaining | Seed question ready: provider failover behavior |
