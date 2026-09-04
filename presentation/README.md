# PhishGuard — Hackathon Pitch Deck

Browser-based phishing detector with **on-device machine learning**.  
12-slide product pitch for a **5–7 minute** judging session.

## Present

1. Open `presentation/index.html` in Chrome (or start the local preview server).
2. Press **F11** (or the browser fullscreen control) for pitch mode.
3. Navigate with **← / →**, **Space**, or the on-screen arrows.

Placeholders to fill before you present:

- `[Team Name]`
- `[Member 1] · [Member 2] · [Member 3]`
- `[Hackathon Name]`
- `github.com/[your-team]/phishguard`
- `[contact@email]`

## Pitch beats (about 30–40s per slide)

1. Title — privacy-first hook  
2. Problem — one click, stolen credentials  
3. Solution — Chrome extension, local verdicts  
4. Architecture — URL + page features → on-device model  
5. Signals — 30–40 URL and content features  
6. ML — compact candidates, still in model-selection  
7. Privacy — cloud vs local comparison  
8. UX — extension popup mockup  
9. Evaluation — metrics and **targets**, not unverified scores  
10. Data — PhishTank, OpenPhish, Tranco  
11. Limits & roadmap — honest gaps  
12. Close — Detect / Protect / Preserve  

## Integrity notes

- Do not present **target** quality (e.g. 95%+) as measured results.
- The 94% risk score on the UI slide is a **mock verdict**, not a benchmark.
- Model choice (gradient-boosted tree vs compact network) is still open.
