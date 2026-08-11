# Parkwise AI - Architecture

## Vision
Smart parking that saves time, fuel, and CO2. From your Arduino Project 08 (Ultrasonic Parking Sensor) to city-scale.

## System Diagram

```
[HC-SR04 x4 + LEDs] -> [ESP32 Node] -> WiFi -> [FastAPI Backend] <-> [WebSocket] <-> [Dashboard]
                                     |                |
                                     |                v
                               [YOLOv8 AI Camera] -> [Sensor Fusion] -> [Predictor] -> [DB]
```

### Layers:

1. **Sensing Layer:**
   - Ultrasonic (accurate, cheap, works in dark) for Slots 1-4
   - Camera + YOLOv8 for Slots 5-12 (vision, no wiring per slot)
   - Fusion: if both say occupied -> 99% confidence, if conflict -> trust ultrasonic

2. **Edge Layer (ESP32):**
   - Median filter (5 readings)
   - Hysteresis (occupied <30cm, free >70cm)
   - HTTP POST every 5s + local LED
   - Offline fallback: still beeps like your original project

3. **Backend Layer (FastAPI):**
   - REST: /api/slots, /api/stats, /api/slots/update, /api/ai/detect
   - WebSocket: realtime broadcast to all dashboards
   - Auto simulation for demo without hardware
   - Future: add PostgreSQL + Redis + MQTT

4. **AI Layer:**
   - detector.py: defines ROI polygons, runs YOLO, checks point-in-polygon
   - predictor.py: simple time-series, in real -> LSTM trained on lot history
   - Metrics: occupancy rate, fill-time prediction, CO2 saved

5. **Presentation Layer:**
   - Tailwind dashboard, glassmorphism
   - Live grid, stats, activity log, forecast
   - Mobile-responsive for parking attendant

## Hardware Bill (for hackathon prototype)

| Item | Qty | Cost |
|------|-----|------|
| ESP32 DevKit | 1 | ₹500 |
| HC-SR04 | 4 | ₹400 |
| Breadboard + wires | 1 | ₹200 |
| LEDs + resistors | 8 | ₹100 |
| Webcam / ESP32-CAM | 1 | ₹800 |
| Total | | ~₹2000 |

## Why this wins hackathon:

- **You have journey:** Explain how you went from blinking LED to smart city. Story sells.
- **Works without hardware:** Backend has mock simulation, judges still see working demo
- **Real IoT + Real AI:** Not just UI, actual Arduino + YOLO narrative
- **Metrics that matter:** CO2, revenue, time saved - impact
