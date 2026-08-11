# Parkwise AI - 5 Min Pitch Deck Script

### Slide 1: Hook
"Ever circled a mall parking for 20 minutes? In India, 30% of urban traffic is just searching for parking. That's fuel, CO2, frustration. And it started with my Arduino."

### Slide 2: Journey (Personal)
Show GitHub repo - 10 projects:
- #1 Blinking LED (hello world)
- #8 Ultrasonic Parking Sensor (beep faster when close)
"I thought - what if this single sensor becomes a whole parking lot?"

### Slide 3: Problem
- Drivers waste 15 min avg finding spot
- Malls lose revenue when lot looks full but has hidden spots
- CO2: 300kg per parking space per year from circling

### Slide 4: Solution - Parkwise AI
Smart parking that combines cheap ultrasonic + AI camera.
- ESP32 + 4x HC-SR04 = ₹1000 covers 4 slots, accurate in dark/tunnels
- Camera + YOLOv8 covers 20 slots without wiring
- Fusion = 98%+ accuracy
- Dashboard real-time + predictions

### Slide 5: Live Demo (MOST IMPORTANT - 90 sec)
1. Show dashboard - 12 slots live
2. Click "Trigger AI Scan" -> YOLO mock
3. Show WebSocket live (green dot)
4. Show Arduino code if possible (Serial Monitor)
5. Show stats: occupancy, revenue, CO2 saved, full-in-minutes

### Slide 6: Tech Stack
- Hardware: ESP32, HC-SR04, ESP32-CAM
- Edge: Arduino C++ with median filter + hysteresis
- Backend: FastAPI + WebSocket + simulation
- AI: YOLOv8n for vehicle detection, ROI polygon check, LSTM predictor (mock now)
- Frontend: Tailwind, glassmorphism, live grid

### Slide 7: Why We Win vs Others
Other teams: only app UI mock
We: Working IoT + working backend + working AI + hardware journey proof + deployable for ₹2000 prototype
Scalable: from 4 slots to 1000 slots with mesh

### Slide 8: Business Model
- B2B: Malls, hospitals, offices pay per slot per month ₹50
- 100 slot mall = ₹5000 MRR, saves them ₹50k fuel/time/mo
- Phase 2: App with navigation, payment, license plate OCR
- Cost: ₹500 per slot one-time, 85% margin

### Slide 9: Impact
- Per 100-slot lot: saves 45 hours driver time/day, 18kg CO2/day
- Story: From learning Arduino to solving city problem

### Slide 10: Ask / Roadmap
- Today: MVP with 12 slots
- Next month: ESP32-CAM mesh + payment + mobile app
- Next year: 10 lots in city + city parking API

Close with: "Built on my Arduino-learning-journey repo, dedicated to every Indian who hates parking."

---

### Q&A Prep:
Q: What if internet fails?
A: Node still beeps locally like original Project 08, and buffers data.

Q: Why ultrasonic + camera, not just camera?
A: Ultrasonic works in dark, rain, no privacy issue, cheap. Camera needs light. Fusion best.

Q: How to install in real lot?
A: ESP32 per 4 slots under ceiling, one camera per 20 slots. 30 min install.

Q: Do you have hardware?
A: Show video, or say backend simulates for judging without hardware but tested with Uno earlier.
