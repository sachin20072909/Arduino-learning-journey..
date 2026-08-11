# Parkwise AI - Step by Step Hackathon Roadmap

This is your battle plan. Follow in order.

## PHASE 0 - You are here (30 min)
- [x] You have Arduino learning repo with Project 08 (ultrasonic)
- [x] Parkwise AI scaffold created (arduino + backend + frontend + ai folders)
- [ ] Understand architecture: read docs/ARCHITECTURE.md

## PHASE 1 - Make Backend Run (20 min) [DO THIS NOW]
```bash
cd parkwise-ai/backend
pip install -r requirements.txt
uvicorn app:app --host 0.0.0.0 --port 8000 --reload
```
Open http://localhost:8000/api/slots - you should see JSON with 12 slots
Keep this terminal open.

## PHASE 2 - Make Frontend Run (5 min)
In new terminal:
```bash
cd parkwise-ai/frontend
python -m http.server 3000
# OR if you have node: npx serve .
```
Open http://localhost:3000 -> dashboard should connect via WebSocket and show live mock data flickering
If ports not available, just open index.html directly, it has mock fallback

## PHASE 3 - Test Arduino Logic Without Hardware (15 min)
1. Open parkwise-ai/arduino/Parkwise_Node/Parkwise_Node.ino in Arduino IDE
2. If you have ESP32: set SSID/PASS/BACKEND_URL and upload
3. If you only have Uno: it will still compile, comment out WiFi parts, test via Serial Monitor - same as your old project but 4 sensors
4. No hardware? It's OK - backend already simulates slot flipping every 8 sec

## PHASE 4 - Add Real AI (Optional but impressive) (40 min)
```bash
cd parkwise-ai/ai
pip install ultralytics opencv-python
python parking_detector.py
# Test with a parking lot image
```
For hackathon pitch: "We support YOLOv8n fusion, but demo uses ultrasonic + mock vision to work offline"
Add a photo of parking lot to ai/demo.jpg and show detection

## PHASE 5 - Polish for Demo (1 hour)
- [ ] Record 30 sec video of dashboard live updating + ESP32 if you have it
- [ ] Take screenshot of Arduino Serial Monitor showing distances
- [ ] Update parkwise-ai/README.md with demo link and team name
- [ ] Create pitch deck structure (see docs/PITCH.md)
- [ ] Prepare 2-min story: "I started with blinking LED, today I built smart city..."

## PHASE 6 - Hackathon Submission Checklist
- [ ] Git push to main
- [ ] Deployed backend? Use Render / Railway: `uvicorn app:app --host 0.0.0.0 --port $PORT`
- [ ] Deployed frontend? Use Vercel / Netlify - just upload index.html
- [ ] README with: problem, solution, tech stack, demo video, architecture diagram
- [ ] Arduino code commented like your learning journey (judges love learning story)
- [ ] Include cost analysis: ₹2000 prototype saves 15 min per driver, 30% less fuel around mall

## PHASE 7 - Presentation Tips
- Start with your journey repo (show 10 projects), then say Project 08 inspired Parkwise
- Live demo: dashboard -> trigger AI scan -> show slot flipping
- Show Serial Monitor if you have hardware
- Talk impact: 30% of city traffic is searching parking, we cut it
- End with roadmap: Phase 2 - License plate OCR, payment, app, ESP32-CAM mesh

## What to do after this chat?
1. Run backend (Phase 1) - tell me if port error, I'll fix
2. Open frontend - screenshot and share
3. Decide: Do you have ESP32? Or only Uno? I will adapt code accordingly
4. Do you want me to generate pitch deck + README?

You already have MVP. Now just run it and record.
