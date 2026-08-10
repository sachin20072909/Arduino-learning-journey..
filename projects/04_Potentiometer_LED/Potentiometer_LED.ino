/*
  Project 04 - Potentiometer Controlled LED
  -----------------------------------------
  Read a knob (potentiometer) and use it to set the brightness of an LED.
  This is the same idea as Fading_of_LED, but with a human in control.

  Wiring:
    - Potentiometer outer pins -> 5V and GND
    - Potentiometer middle pin -> analog pin A0
    - LED anode -> pin 9 (a PWM pin) through a 220 ohm resistor -> GND

  Concepts learned:
    - analogRead() returns 0..1023
    - analogWrite() expects 0..255
    - map() to convert between ranges
*/

const int potPin = A0;   // potentiometer wiper
const int ledPin = 9;    // PWM-capable pin for the LED

void setup() {
  pinMode(ledPin, OUTPUT);
  // analog pins don't need pinMode() for reading
}

void loop() {
  int potValue   = analogRead(potPin);            // 0 .. 1023
  int brightness = map(potValue, 0, 1023, 0, 255); // 0 .. 255

  analogWrite(ledPin, brightness);
  delay(10);   // small delay to keep the reading stable
}
