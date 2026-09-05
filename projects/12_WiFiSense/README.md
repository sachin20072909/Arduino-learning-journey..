# 12 — WiFiSense (hackathon pitch deck)

**"Privacy-Preserving Human Activity Detection Using WiFi CSI"**
_Tagline: Sense movement. Not identities._

A 10-slide, 16:9 dark-tech hackathon deck for a system that tells whether a
person is present or moving by reading **WiFi Channel State Information (CSI)**
from 2–3 ESP32 boards — no camera, no microphone, no wearables.

```
projects/12_WiFiSense/
├── WiFiSense_Hackathon_Deck.pptx   ← the deliverable (10 slides, 16:9)
├── deck/
│   ├── prepare_assets.py           ← crops + tones the artwork → assets/processed
│   ├── build_deck.py               ← assembles the .pptx
│   └── qa_check.py                 ← geometry audit (bounds, distortion, overflow)
└── assets/                         ← source art (not committed, see .gitignore)
```

## Slide map

| # | Slide | One idea it communicates |
|---|-------|--------------------------|
| 1 | Title | WiFiSense — sense movement, not identities |
| 2 | Problem | Cameras are not always the answer (privacy, blind spots, darkness, wearables) |
| 3 | Solution | Don't watch the room — read the WiFi |
| 4 | How CSI works | Five steps from packets on air to an ML-readable pattern |
| 5 | The hard problem | A changed signal is **not** always a human |
| 6 | ML pipeline | Raw CSI → filtering → features → classification → confidence |
| 7 | Architecture | ESP32 links → edge processing → dashboard → alerts |
| 8 | Use cases | Six places this fits, with an honest scope note |
| 9 | Live demo | Five steps in 90 s, incl. the bag/chair false-positive test |
| 10 | Impact & future | Why it matters, what comes next, closing line |

## What the deck deliberately does

- **Reports confidence, not certainty.** The dashboard shows probabilities;
  below threshold the system says "uncertain".
- **Treats non-human disturbance as a first-class problem.** Slide 5 is
  dedicated to bags, chairs, fans and doors, and slide 9's step 5 demos it.
- **Does not over-claim.** No "100 % accurate", no "works through every wall",
  and no claim that WiFi sensing itself is new — the contribution is presented
  as a low-cost ESP32 implementation with explicit non-human rejection.
- **Text-light, diagram-heavy.** Every slide carries one idea and at least one
  large visual.

## Rebuilding the deck

```bash
pip install python-pptx pillow numpy

python3 deck/prepare_assets.py     # crops + tone-curves the art
python3 deck/prepare_assets.py --stats   # optional: per-asset brightness report
python3 deck/build_deck.py         # writes WiFiSense_Hackathon_Deck.pptx
python3 deck/qa_check.py           # geometry audit — should print "clean"
```

The AI artwork lives in `assets/gen/` (excluded from git because of size);
`assets/processed/` holds the derived crops. The generated art is very dark, so
`prepare_assets.py` *lifts* it with a gamma curve before use and re-scrims it
only where headline text sits on top.

## Design system

| Role | Value |
|------|-------|
| Background | `#05070C` |
| Panels | `#0B1220` / `#0E1728` |
| Primary accent | cyan `#22D3EE` |
| Secondary | violet `#8B5CF6`, sky `#0EA5E9` |
| Status | green `#34D399`, amber `#FBBF24`, red `#FB7185` |
| Type | Segoe UI — 30 pt titles, ≤14 pt body, letter-spaced eyebrows |

## Hardware / software story the deck pitches

- **Hardware:** 2–3 ESP32 dev boards, USB cables, a laptop, optional router
  and power banks.
- **Software:** ESP-IDF CSI firmware, Python, NumPy / Pandas, scikit-learn,
  FastAPI backend, lightweight web dashboard.
