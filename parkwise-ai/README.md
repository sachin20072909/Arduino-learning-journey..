# Parkwise AI - Smart Parking Management System
> From Blinking LED to Smart City. Built on Arduino Learning Journey.

![Status](https://img.shields.io/badge/status-hackathon%20MVP-green)
![Stack](https://img.shields.io/badge/stack-ESP32%20%2B%20FastAPI%20%2B%20YOLOv8%20%2B%20WebSocket-blue)

### 🎥 Demo Video (add link)
[Link to dashboard / loom video]

### 🌐 Live Demo
- Backend: `https://your-backend.onrender.com` 
- Dashboard: Frontend runs on WebSocket at `ws://backend/ws`

### Problem
- 30% urban traffic = searching parking
- Avg 15 min wasted per driver
- No real-time visibility for mall operators

### Solution
**Parkwise AI** fuses:
- **Ultrasonic sensing** (your Project 08 improved to 4 slots) - cheap, works in dark, 99% accurate for near-range
- **AI Vision** (YOLOv8) - one camera covers 12+ slots, no per-slot wiring
- **Realtime dashboard** with occupancy, predictions, CO2 saved, revenue

### Architecture
```
[HC-SR04 x4 + ESP32] --WiFi HTTP POST--> [FastAPI] --WebSocket--> [Dashboard]
[Camera] --YOLOv8--> [AI Detector] --Fusion--> [FastAPI]
```

### Repo Structure
```
parkwise-ai/
├── arduino/Parkwise_Node/  - ESP32 code (upgraded from Project 08)
├── backend/                 - FastAPI + WebSocket + mock simulation
├── frontend/                - Glassmorphism dashboard (Tailwind)
├── ai/                      - YOLOv8 detector + predictor
└── docs/                    - Architecture, roadmap, pitch
```

### How to Run

**Backend:**
```bash
cd backend
pip install -r requirements.txt
uvicorn app:app --host 0.0.0.0 --port 8000 --reload
# Open http://localhost:8000/api/slots
```

**Frontend:**
```bash
cd frontend
python -m http.server 3000
# Open http://localhost:3000
```

**Arduino:**
1. Install ESP32 board in Arduino IDE
2. Open `arduino/Parkwise_Node/Parkwise_Node.ino`
3. Set WIFI_SSID, WIFI_PASS, BACKEND_URL
4. Upload to ESP32. Wire HC-SR04 per comments.

**Without hardware?**
Backend auto-simulates slot flips every 8 sec. Frontend has mock fallback. You can still demo full flow.

### Key Features for Hackathon Judges
- ✅ Working MVP (even offline)
- ✅ Real IoT code (not just UI)
- ✅ Real AI narrative (YOLOv8 + sensor fusion)
- ✅ Scalable economics (₹500/slot)
- ✅ Impact metrics (time, CO2, revenue)
- ✅ Journey story from learning repo

### Tech Stack
ESP32, HC-SR04, Arduino C++, FastAPI, WebSocket, YOLOv8, TailwindCSS

### Roadmap
- v1.0 (Hackathon): 12 slots, dashboard, simulation
- v1.5: ESP32-CAM mesh, PostgreSQL, mobile PWA
- v2.0: License plate OCR, UPI payment, navigation inside mall

### Team
Sachin - Arduino Learner -> Smart City Builder

### Link to Learning Journey
This repo extends [Arduino-learning-journey projects 01-10]. Project 08 (Ultrasonic Parking Sensor) was the seed for this.

---
Built for Hackathon - 2026
