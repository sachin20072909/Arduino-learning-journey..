"""
Parkwise AI - Backend (FastAPI)
Receives data from ESP32 nodes, serves dashboard API, and broadcasts realtime via WebSocket
Run: uvicorn app:app --host 0.0.0.0 --port 8000 --reload
"""

from fastapi import FastAPI, WebSocket, WebSocketDisconnect
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import List, Dict
from datetime import datetime
import asyncio
import random
import json

app = FastAPI(title="Parkwise AI API", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ---- Data Models ----
class SlotUpdate(BaseModel):
    id: int
    distance_cm: float
    occupied: bool
    sensor: str = "ultrasonic"

class NodePayload(BaseModel):
    node_id: str
    slots: List[SlotUpdate]

class Slot(BaseModel):
    id: int
    node_id: str
    distance_cm: float
    occupied: bool
    sensor: str
    last_update: str
    confidence: float = 0.98  # for AI demo
    vehicle_type: str = "none" # car, bike, none - for AI vision demo

# In-memory DB (replace with Redis/Postgres for production)
parking_lot_state: Dict[int, Slot] = {}
stats = {
    "total_slots": 12,
    "occupied_today": 0,
    "revenue_today": 0.0
}

# Initialize with 12 slots for demo
for i in range(1, 13):
    parking_lot_state[i] = Slot(
        id=i,
        node_id="PARKWISE_NODE_01" if i <= 4 else "AI_CAMERA_01",
        distance_cm=150.0,
        occupied=random.choice([True, False]) if i > 4 else False,
        sensor="ultrasonic" if i <= 4 else "camera_ai",
        last_update=datetime.now().isoformat(),
        confidence=random.uniform(0.92, 0.99),
        vehicle_type=random.choice(["car", "none"]) 
    )

# WebSocket manager
class ConnectionManager:
    def __init__(self):
        self.active_connections: List[WebSocket] = []
    async def connect(self, ws: WebSocket):
        await ws.accept()
        self.active_connections.append(ws)
    def disconnect(self, ws: WebSocket):
        self.active_connections.remove(ws)
    async def broadcast(self, message: str):
        for connection in self.active_connections[:]:
            try:
                await connection.send_text(message)
            except:
                pass

manager = ConnectionManager()

@app.get("/")
def root():
    return {"message": "Parkwise AI Backend Running", "slots": len(parking_lot_state)}

@app.get("/api/slots")
def get_slots():
    return {
        "total": len(parking_lot_state),
        "available": len([s for s in parking_lot_state.values() if not s.occupied]),
        "occupied": len([s for s in parking_lot_state.values() if s.occupied]),
        "slots": list(parking_lot_state.values())
    }

@app.get("/api/stats")
def get_stats():
    occupied = len([s for s in parking_lot_state.values() if s.occupied])
    total = len(parking_lot_state)
    return {
        "total_slots": total,
        "occupied": occupied,
        "free": total - occupied,
        "occupancy_rate": round(occupied/total*100, 1) if total else 0,
        "predicted_full_in_minutes": random.randint(15, 120) if occupied/total < 0.8 else 0,
        "revenue_today": occupied * 20 + random.randint(100, 500),
        "co2_saved_kg": round(occupied * 0.12, 2), # marketing metric
        "peak_hours": ["9 AM - 11 AM", "5 PM - 8 PM"]
    }

@app.post("/api/slots/update")
async def update_slots(payload: NodePayload):
    """Called by ESP32 Arduino Node"""
    for slot_update in payload.slots:
        # Merge - keep same ID space for simplicity
        existing = parking_lot_state.get(slot_update.id)
        if existing:
            parking_lot_state[slot_update.id] = Slot(
                id=slot_update.id,
                node_id=payload.node_id,
                distance_cm=slot_update.distance_cm,
                occupied=slot_update.occupied,
                sensor=slot_update.sensor,
                last_update=datetime.now().isoformat(),
                confidence=0.99 if slot_update.sensor == "ultrasonic" else existing.confidence,
                vehicle_type="car" if slot_update.occupied else "none"
            )
        else:
            parking_lot_state[slot_update.id] = Slot(
                id=slot_update.id,
                node_id=payload.node_id,
                distance_cm=slot_update.distance_cm,
                occupied=slot_update.occupied,
                sensor=slot_update.sensor,
                last_update=datetime.now().isoformat()
            )
    
    # broadcast to all dashboards
    data = await get_slots_data()
    await manager.broadcast(json.dumps({"type": "update", "data": data}))
    return {"status": "ok", "updated": len(payload.slots)}

async def get_slots_data():
    occupied = len([s for s in parking_lot_state.values() if s.occupied])
    total = len(parking_lot_state)
    return {
        "total": total,
        "available": total - occupied,
        "occupied": occupied,
        "slots": [s.model_dump() for s in parking_lot_state.values()]
    }

# Mock AI camera endpoint
@app.post("/api/ai/detect")
async def ai_detect(payload: dict):
    """
    Simulates YOLOv8 vehicle detection from camera feed
    In real hackathon, this would receive image and run model
    """
    # Mock detection
    fake_detections = []
    for i in range(5, 13): # camera covers slot 5-12
        occupied = random.random() > 0.4
        fake_detections.append({
            "slot_id": i,
            "occupied": occupied,
            "confidence": round(random.uniform(0.85, 0.99), 3),
            "bbox": [random.randint(0,640), random.randint(0,480), 80, 150],
            "vehicle_type": "car" if occupied else "none"
        })
        # update state
        if i in parking_lot_state:
            parking_lot_state[i].occupied = occupied
            parking_lot_state[i].confidence = fake_detections[-1]["confidence"]
            parking_lot_state[i].last_update = datetime.now().isoformat()
            parking_lot_state[i].vehicle_type = fake_detections[-1]["vehicle_type"]

    data = await get_slots_data()
    await manager.broadcast(json.dumps({"type": "ai_update", "data": data, "detections": fake_detections}))
    
    return {"detections": fake_detections, "parking": data}

# For demo: auto-simulate changes every 8 seconds
@app.on_event("startup")
async def startup_simulation():
    async def simulate():
        while True:
            await asyncio.sleep(8)
            # flip a random slot
            sid = random.randint(1, len(parking_lot_state))
            if sid in parking_lot_state:
                # 30% chance to flip
                if random.random() < 0.3:
                    parking_lot_state[sid].occupied = not parking_lot_state[sid].occupied
                    parking_lot_state[sid].last_update = datetime.now().isoformat()
                    data = await get_slots_data()
                    try:
                        await manager.broadcast(json.dumps({"type": "sim", "data": data}))
                    except:
                        pass
    asyncio.create_task(simulate())

@app.websocket("/ws")
async def websocket_endpoint(websocket: WebSocket):
    await manager.connect(websocket)
    try:
        # send initial state
        data = await get_slots_data()
        await websocket.send_text(json.dumps({"type": "init", "data": data}))
        while True:
            await websocket.receive_text() # keep alive, ignore incoming
    except WebSocketDisconnect:
        manager.disconnect(websocket)
