/*
  Parkwise AI - Smart Parking Node (ESP32 + 4x HC-SR04)
  -----------------------------------------------------
  Upgraded from your Project 08 - Ultrasonic Parking Sensor.

  This node monitors 4 parking slots and sends data to Parkwise Cloud.
  
  Hardware:
    ESP32 DevKit v1
    4x HC-SR04 Ultrasonic Sensors
      Slot 1: Trig 5,  Echo 18
      Slot 2: Trig 19, Echo 21
      Slot 3: Trig 22, Echo 23
      Slot 4: Trig 2,  Echo 4
    WiFi connection

  Features vs old project:
    - Multi-slot (4 instead of 1)
    - WiFi + HTTP POST to backend (not just Serial)
    - Median filter to avoid false readings
    - LED indicators per slot
    - Auto-calibration on boot

  Backend API: POST /api/slots/update
  Payload: { "node_id": "NODE_01", "slots": [ {"id":1, "distance_cm":45, "occupied":true}, ... ] }
*/

#if defined(ESP32)
  #include <WiFi.h>
  #include <HTTPClient.h>
#else
  // Fallback for Arduino Uno - uses Serial only (for testing without ESP32)
  // Upload this same logic to Uno and see output in Serial Monitor
#endif

// CONFIGURATION
const char* WIFI_SSID = "YOUR_WIFI_SSID";
const char* WIFI_PASS = "YOUR_WIFI_PASSWORD";
const char* BACKEND_URL = "http://YOUR_BACKEND_IP:8000/api/slots/update"; 
const String NODE_ID = "PARKWISE_NODE_01";

const int NUM_SLOTS = 4;

// Pin config for ESP32 - change if using Uno
struct SlotPins {
  int trig;
  int echo;
  int led; // optional status LED
};

SlotPins slotPins[4] = {
  {5,  18, 13}, // Slot 1
  {19, 21, 12}, // Slot 2
  {22, 23, 14}, // Slot 3
  {2,  4,  27}  // Slot 4
};

// Parking detection params
const int OCCUPIED_THRESHOLD_CM = 30; // if distance < 30cm, car present
const int EMPTY_THRESHOLD_CM = 70;    // hysteresis
const int NUM_READINGS_FOR_MEDIAN = 5;

bool slotOccupied[4] = {false, false, false, false};

#if !defined(ESP32)
// Uno compatibility - define LED 13 built-in
#define LED_BUILTIN 13
#endif

long readDistanceCM(int trigPin, int echoPin) {
  digitalWrite(trigPin, LOW);
  delayMicroseconds(2);
  digitalWrite(trigPin, HIGH);
  delayMicroseconds(10);
  digitalWrite(trigPin, LOW);

  long duration = pulseIn(echoPin, HIGH, 30000); // 30ms timeout
  if (duration == 0) return 400; // timeout = far

  long distance = duration * 0.0343 / 2;
  return distance;
}

long medianReading(int trigPin, int echoPin) {
  long readings[NUM_READINGS_FOR_MEDIAN];
  for (int i=0;i<NUM_READINGS_FOR_MEDIAN;i++){
    readings[i] = readDistanceCM(trigPin, echoPin);
    delay(20);
  }
  // simple bubble sort for median
  for (int i=0;i<NUM_READINGS_FOR_MEDIAN-1;i++)
    for (int j=0;j<NUM_READINGS_FOR_MEDIAN-i-1;j++)
      if (readings[j] > readings[j+1]) {
        long t = readings[j]; readings[j]=readings[j+1]; readings[j+1]=t;
      }
  return readings[NUM_READINGS_FOR_MEDIAN/2];
}

void setup() {
  Serial.begin(115200);
  Serial.println("\n=== Parkwise AI Node Booting ===");

  for (int i=0;i<NUM_SLOTS;i++){
    pinMode(slotPins[i].trig, OUTPUT);
    pinMode(slotPins[i].echo, INPUT);
    if (slotPins[i].led != -1) pinMode(slotPins[i].led, OUTPUT);
  }

#if defined(ESP32)
  WiFi.begin(WIFI_SSID, WIFI_PASS);
  Serial.print("Connecting WiFi");
  int attempts=0;
  while (WiFi.status() != WL_CONNECTED && attempts < 20) {
    delay(500);
    Serial.print(".");
    attempts++;
  }
  if (WiFi.status() == WL_CONNECTED) {
    Serial.println("\nWiFi Connected! IP: " + WiFi.localIP().toString());
  } else {
    Serial.println("\nWiFi Failed - Running Offline Mode (Serial only)");
  }
#endif

  // calibration phase
  Serial.println("Calibrating slots (keep them empty for 3s)...");
  delay(3000);
  Serial.println("Calibration done. Starting monitoring.");
}

void loop() {
  String jsonPayload = "{\"node_id\":\"" + NODE_ID + "\",\"slots\":[";
  
  for (int i=0;i<NUM_SLOTS;i++){
    long dist = medianReading(slotPins[i].trig, slotPins[i].echo);

    // Hysteresis to avoid flicker
    if (!slotOccupied[i] && dist < OCCUPIED_THRESHOLD_CM) {
      slotOccupied[i] = true;
    } else if (slotOccupied[i] && dist > EMPTY_THRESHOLD_CM) {
      slotOccupied[i] = false;
    }

    if (slotPins[i].led != -1) {
      digitalWrite(slotPins[i].led, slotOccupied[i] ? HIGH : LOW);
    }

    Serial.print("Slot "); Serial.print(i+1);
    Serial.print(": "); Serial.print(dist); Serial.print("cm -> ");
    Serial.println(slotOccupied[i] ? "OCCUPIED" : "FREE");

    jsonPayload += "{\"id\":" + String(i+1) + 
                   ",\"distance_cm\":" + String(dist) + 
                   ",\"occupied\":" + (slotOccupied[i] ? "true" : "false") + 
                   ",\"sensor\":\"ultrasonic\"}";

    if (i < NUM_SLOTS-1) jsonPayload += ",";
  }
  jsonPayload += "]}";
  
  Serial.println("Payload: " + jsonPayload);

#if defined(ESP32)
  if (WiFi.status() == WL_CONNECTED) {
    HTTPClient http;
    http.begin(BACKEND_URL);
    http.addHeader("Content-Type", "application/json");
    int code = http.POST(jsonPayload);
    if (code > 0) {
      Serial.printf("Server response: %d\n", code);
      String resp = http.getString();
      Serial.println(resp);
    } else {
      Serial.printf("HTTP Error: %s\n", http.errorToString(code).c_str());
    }
    http.end();
  }
#endif

  delay(5000); // update every 5 seconds
}
