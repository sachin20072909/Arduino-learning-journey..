# CityVerse AI — Hackathon Pitch Deck

**CityVerse_AI_Pitch_Deck.pptx** — a 12-slide, competition-ready pitch deck for the
"Smart Digital Twin Platform for Smart Cities" problem statement.

## What's inside

| # | Slide | Highlights |
|---|-------|-----------|
| 01 | Opening Impact | AI-generated hero artwork, HUD frame, cinematic build-in |
| 02 | The Problem | Split-screen story: chaos vs. siloed tools, pulsing alert chip |
| 03 | Our Vision | Physical → Twin → AI → Decisions pipeline with flowing data dots |
| 04 | Meet CityVerse AI | 5-module orbit around a live city hub |
| 05 | How It Works | 5-stage sensor-to-dashboard pipeline, animated data particles |
| 06 | Live Digital City | GIS map + pulsing sensors + live vitals cards |
| 07 | AI Intelligence | Neural core art + 3 prediction cards with confidence bars |
| 08 | Emergency Simulation | Simulated flood rerouting map + 5-step response timeline |
| 09 | Governance Dashboard | Full command-center: heatmap, energy bars, AQI trend, AI recommendation |
| 10 | Technology Stack | 5-layer architecture cards |
| 11 | Impact | Big gradient metrics + beneficiary row |
| 12 | Closing | Dawn city, closing quote, cinematic fade-to-black |

## Motion design

- **Morph transitions** on every slide (by-object, ~1.1s) with a Fade fallback for
  older PowerPoint versions.
- **100+ choreographed on-slide animations**: fade, wipe, zoom (scale+fade),
  wheel (gauge sweep), motion paths (data particles, scan-lines, rises), and
  looping pulses / breathing alerts.
- All animations are **one-click-per-slide builds**: press anywhere once and the
  scene choreographs itself automatically (cascading "after previous" timing).

## Speaker notes

Every slide carries a 20–30 second presenter script → ~5.5 minutes total.

## File layout

```
CityVerse_AI_Pitch_Deck.pptx   ← the deck (fully editable, all native shapes/text)
build_deck.py                  ← regenerates the deck
deck_lib.py                    ← design system (glassmorphism, glows, anim XML)
assets/                        ← AI-generated artwork (hero, maps, neural core)
```

Regenerate after edits: `python3 build_deck.py` (needs `python-pptx`).
