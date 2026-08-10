/*
  Project 10 - Knock Detector
  ---------------------------
  A piezo buzzer used in reverse as a vibration sensor: it produces a
  small voltage when knocked. Light an LED when someone knocks on the desk.

  Wiring:
    - Piezo "+" -> A0, piezo "-" -> GND
      (a 1 Megaohm resistor between A0 and GND improves sensitivity)
    - LED anode -> pin 12 through a 220 ohm resistor -> GND

  Concepts learned:
    - sensors can often work in reverse (piezo = speaker AND mic)
    - detecting a spike above a threshold
    - using millis() for timing instead of delay()
*/

const int piezoPin = A0;
const int ledPin   = 12;

const int knockThreshold = 100;  // voltage spike level that counts as a knock
const unsigned long ledDuration = 1500;  // how long LED stays on (ms)

unsigned long ledOnTime = 0;

void setup() {
  pinMode(ledPin, OUTPUT);
  Serial.begin(9600);
}

void loop() {
  int reading = analogRead(piezoPin);
  Serial.println(reading);   // watch values to tune the threshold

  if (reading > knockThreshold) {
    digitalWrite(ledPin, HIGH);
    ledOnTime = millis();   // remember when the knock happened
  }

  // Turn the LED off after ledDuration milliseconds — without blocking
  if (millis() - ledOnTime > ledDuration) {
    digitalWrite(ledPin, LOW);
  }
}
