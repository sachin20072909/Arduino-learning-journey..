/*
  Project 11 - Parkwise AI Node (Smart Parking Bay)
  -------------------------------------------------
  Hackathon MVP for Parkwise AI.

  One ultrasonic sensor watches a parking bay. When a car pulls in,
  the bay is marked OCCUPIED (red LED, barrier down). When it leaves,
  the bay is FREE (green LED, barrier up). A buzzer beeps faster the
  closer the car gets (reusing Project 08 logic), and a CSV line is
  printed over Serial every second so a laptop dashboard (Python /
  Node.js / web page) can read the live state.

  Wiring:
    - HC-SR04: VCC->5V, GND->GND, Trig->D10, Echo->D9
    - Green LED (+ -> 220R -> D3, - -> GND)   = bay FREE
    - Red LED   (+ -> 220R -> D2, - -> GND)   = bay OCCUPIED
    - SG90 servo: signal -> D4, VCC->5V, GND->GND  (boom barrier)
    - Piezo buzzer: + -> D8, - -> GND

  Serial output (9600 baud, CSV):
    BAY,<id>,<STATE>,<distance_cm>
  e.g.
    BAY,1,FREE,152
    BAY,1,OCCUPIED,9

  Concepts:
    - State machine with debounce (a passing cat won't flip state)
    - Non-blocking timing with millis()
    - Servo as boom barrier
    - Structured serial telemetry for a dashboard
*/

#include <Servo.h>

// ---------- Pins ----------
const int trigPin   = 10;
const int echoPin   = 9;
const int redPin    = 2;
const int greenPin  = 3;
const int servoPin  = 4;
const int buzzerPin = 8;

// ---------- Bay / detection config ----------
const int BAY_ID            = 1;
const int TRIGGER_DISTANCE  = 15;   // cm: closer than this = something in the bay
const unsigned long DEBOUNCE_MS = 2000; // wait this long before trusting a state change

// ---------- Servo positions (tune for your servo/barrier) ----------
const int BARRIER_UP   = 0;
const int BARRIER_DOWN = 90;

Servo barrier;

// ---------- State machine ----------
enum BayState { UNKNOWN, FREE, OCCUPIED };
BayState state        = UNKNOWN;
BayState lastReading  = UNKNOWN;
unsigned long readingSince = 0;       // when the current candidate reading started

// ---------- Timing ----------
unsigned long lastTelemetry = 0;
const unsigned long TELEMETRY_MS = 1000;

// ---------- Buzzer ----------
unsigned long lastBeep = 0;

// ---------------------------------------------------------------
long readDistanceCm() {
  digitalWrite(trigPin, LOW);
  delayMicroseconds(2);
  digitalWrite(trigPin, HIGH);
  delayMicroseconds(10);
  digitalWrite(trigPin, LOW);

  long duration = pulseIn(echoPin, HIGH, 30000UL); // 30ms timeout => ~5m max
  if (duration == 0) return 999;                   // out of range
  return duration * 0.0343 / 2;
}

void setBayOutputs(BayState s) {
  switch (s) {
    case FREE:
      digitalWrite(greenPin, HIGH);
      digitalWrite(redPin, LOW);
      barrier.write(BARRIER_UP);
      noTone(buzzerPin);
      break;
    case OCCUPIED:
      digitalWrite(greenPin, LOW);
      digitalWrite(redPin, HIGH);
      barrier.write(BARRIER_DOWN);
      break;
    case UNKNOWN:
    default:
      digitalWrite(greenPin, LOW);
      digitalWrite(redPin, LOW);
      noTone(buzzerPin);
      break;
  }
}

void emitTelemetry(long cm) {
  const char* stateStr = (state == FREE) ? "FREE" :
                         (state == OCCUPIED) ? "OCCUPIED" : "UNKNOWN";
  Serial.print("BAY,");
  Serial.print(BAY_ID);
  Serial.print(',');
  Serial.print(stateStr);
  Serial.print(',');
  Serial.println(cm);
}

// Proximity beeps while a car is approaching (only during UNKNOWN->transition)
void updateBuzzer(long cm, BayState candidate) {
  if (candidate == OCCUPIED) {
    unsigned long beepGap;
    int freq;
    int beepLen;
    if (cm < 10)      { beepGap = 150; freq = 1000; beepLen = 60;  }
    else if (cm < 30) { beepGap = 400; freq = 800;  beepLen = 120; }
    else              { beepGap = 800; freq = 600;  beepLen = 150; }

    unsigned long now = millis();
    if (now - lastBeep >= beepGap) {
      lastBeep = now;
      tone(buzzerPin, freq, beepLen);
    }
  } else {
    noTone(buzzerPin);
  }
}

// ---------------------------------------------------------------
void setup() {
  pinMode(trigPin, OUTPUT);
  pinMode(echoPin, INPUT);
  pinMode(redPin, OUTPUT);
  pinMode(greenPin, OUTPUT);
  pinMode(buzzerPin, OUTPUT);

  barrier.attach(servoPin);
  barrier.write(BARRIER_UP);

  Serial.begin(9600);
  while (!Serial) { ; }  // wait for serial (Leonardo/Micro only; harmless on Uno)
  Serial.println("Parkwise AI Node booted");
}

void loop() {
  unsigned long now = millis();
  long cm = readDistanceCm();

  // Candidate state based on instantaneous reading
  BayState candidate = (cm < TRIGGER_DISTANCE) ? OCCUPIED : FREE;

  // Debounce: only change state if candidate is stable for DEBOUNCE_MS
  if (candidate != lastReading) {
    lastReading  = candidate;
    readingSince = now;
  } else if (candidate != state && (now - readingSince) >= DEBOUNCE_MS) {
    state = candidate;
    setBayOutputs(state);
    Serial.print("STATE_CHANGE,BAY,");
    Serial.print(BAY_ID);
    Serial.print(',');
    Serial.println(state == FREE ? "FREE" : "OCCUPIED");
  }

  // Beep pattern only while a car is entering (candidate OCCUPIED but not yet locked)
  // plus when already OCCUPIED and very close, to mimic parking sensor.
  if (state != FREE || candidate == OCCUPIED) {
    updateBuzzer(cm, candidate);
  } else {
    noTone(buzzerPin);
  }

  // Telemetry once per second
  if (now - lastTelemetry >= TELEMETRY_MS) {
    lastTelemetry = now;
    emitTelemetry(cm);
  }

  delay(20); // small pause, doesn't break timing because we use millis()
}
