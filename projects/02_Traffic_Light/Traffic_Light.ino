/*
  Project 02 - Traffic Light
  --------------------------
  Simulate a traffic signal with red, yellow and green LEDs.

  Wiring:
    - Red LED    -> pin 4 through a 220 ohm resistor -> GND
    - Yellow LED -> pin 3 through a 220 ohm resistor -> GND
    - Green LED  -> pin 2 through a 220 ohm resistor -> GND

  Concepts learned:
    - controlling multiple outputs
    - writing helper functions to keep loop() clean
*/

const int redPin    = 4;
const int yellowPin = 3;
const int greenPin  = 2;

void setup() {
  pinMode(redPin, OUTPUT);
  pinMode(yellowPin, OUTPUT);
  pinMode(greenPin, OUTPUT);
}

void loop() {
  // Green: go
  lightOnly(greenPin);
  delay(5000);

  // Yellow: get ready
  lightOnly(yellowPin);
  delay(1500);

  // Red: stop
  lightOnly(redPin);
  delay(5000);
}

// Turn one LED on and the other two off
void lightOnly(int pin) {
  digitalWrite(redPin,    (pin == redPin)    ? HIGH : LOW);
  digitalWrite(yellowPin, (pin == yellowPin) ? HIGH : LOW);
  digitalWrite(greenPin,  (pin == greenPin)  ? HIGH : LOW);
}
