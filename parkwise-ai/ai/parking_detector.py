"""
Parkwise AI - AI Vision Module
This file shows how to integrate YOLOv8 for real parking detection.
For hackathon demo without GPU, we provide mock + instructions for real model.

Setup real model:
  pip install ultralytics opencv-python
  yolo predict model=yolov8n.pt source=parking_lot.jpg

Parking detection logic:
  1. Define ROI polygons for each parking slot (from config)
  2. Run vehicle detection (car, bike, truck)
  3. Check if center of detected vehicle is inside slot polygon
  4. Return occupancy per slot
"""

import random
import json
from datetime import datetime

# Mock slot definitions - you'd calibrate these from top-down camera
PARKING_SLOTS_ROI = {
    5: [(100,100),(200,100),(200,250),(100,250)],
    6: [(220,100),(320,100),(320,250),(220,250)],
    7: [(340,100),(440,100),(440,250),(340,250)],
    8: [(460,100),(560,100),(560,250),(460,250)],
    9: [(100,270),(200,270),(200,420),(100,420)],
    10: [(220,270),(320,270),(320,420),(220,420)],
    11: [(340,270),(440,270),(440,420),(340,420)],
    12: [(460,270),(560,270),(560,420),(460,420)],
}

class ParkwiseDetector:
    def __init__(self, use_real_yolo=False):
        self.use_real = use_real_yolo
        self.model = None
        if use_real_yolo:
            try:
                from ultralytics import YOLO
                self.model = YOLO('yolov8n.pt') # nano for speed
                print("YOLOv8 loaded")
            except Exception as e:
                print(f"YOLO not available, falling back to mock: {e}")
                self.use_real = False
    
    def is_point_in_polygon(self, point, polygon):
        # ray casting
        x, y = point
        inside = False
        n = len(polygon)
        px1, py1 = polygon[0]
        for i in range(n+1):
            px2, py2 = polygon[i % n]
            if y > min(py1, py2):
                if y <= max(py1, py2):
                    if x <= max(px1, px2):
                        if py1 != py2:
                            xinters = (y-py1)*(px2-px1)/(py2-py1)+px1
                        if px1 == px2 or x <= xinters:
                            inside = not inside
            px1, py1 = px2, py2
        return inside

    def detect_mock(self, image_path=None):
        """Mock detection for hackathon without camera"""
        results = []
        for slot_id, roi in PARKING_SLOTS_ROI.items():
            occupied = random.random() > 0.45
            cx = sum([p[0] for p in roi])//4
            cy = sum([p[1] for p in roi])//4
            results.append({
                "slot_id": slot_id,
                "occupied": occupied,
                "confidence": round(random.uniform(0.88, 0.99), 3),
                "vehicle_type": random.choice(["car","car","bike"]) if occupied else "none",
                "center": (cx, cy),
                "bbox": [cx-40, cy-60, 80, 120],
                "timestamp": datetime.now().isoformat()
            })
        return results

    def detect_real(self, image_path):
        """Real YOLO detection - call this when you have camera"""
        if not self.model:
            return self.detect_mock(image_path)
        
        # Real logic pseudocode:
        # results = self.model(image_path, classes=[2,3,5,7]) # car, motorbike, bus, truck
        # ... process
        # For now return mock
        return self.detect_mock(image_path)

    def get_occupancy_summary(self, detections):
        occupied = sum(1 for d in detections if d["occupied"])
        return {
            "total": len(detections),
            "occupied": occupied,
            "free": len(detections)-occupied,
            "occupancy_rate": occupied/len(detections) if detections else 0,
            "slots": detections
        }

# --- Prediction module (simple time-series mock) ---
class OccupancyPredictor:
    def __init__(self):
        # In real project: train LSTM on historic data
        self.peak_hours = [8,9,10,17,18,19] # 9-11 AM, 5-8 PM
        
    def predict(self, current_occupancy_rate, hour_now=None):
        if hour_now is None:
            hour_now = datetime.now().hour
        
        # Simple heuristic: if near peak and occupancy < 80%, predict fast fill
        if hour_now in self.peak_hours:
            fill_rate_per_min = 2.5 # % per minute
        else:
            fill_rate_per_min = 0.6
        
        remaining = 100 - (current_occupancy_rate*100)
        minutes_to_full = int(remaining / fill_rate_per_min) if fill_rate_per_min>0 else 999
        
        # Predict availability for next 2 hours
        forecast = []
        rate = current_occupancy_rate
        for i in range(1, 9): # 15min intervals
            rate = min(1.0, rate + (fill_rate_per_min/100)*15* random.uniform(0.8,1.2))
            forecast.append({
                "in_minutes": i*15,
                "predicted_occupancy": round(rate, 2)
            })
        
        return {
            "minutes_to_full": minutes_to_full,
            "forecast": forecast,
            "recommendation": "Open overflow lot" if minutes_to_full < 30 else "Normal operations"
        }

if __name__ == "__main__":
    print("=== Parkwise AI Detection Test ===")
    detector = ParkwiseDetector(use_real_yolo=False)
    detections = detector.detect_mock()
    print(json.dumps(detector.get_occupancy_summary(detections), indent=2))

    predictor = OccupancyPredictor()
    print("\nPrediction:")
    print(json.dumps(predictor.predict(0.65, hour_now=9), indent=2))
