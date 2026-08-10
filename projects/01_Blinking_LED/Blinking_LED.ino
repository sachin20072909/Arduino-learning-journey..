/*
  Project 01 - Blinking LED
  -------------------------
  The "Hello World" of Arduino: blink an LED on and off.

  Wiring:
    - LED anode (long leg) -> pin 13 through a 220 ohm resistor
    - LED cathode (short leg) -> GND
    - Or just use the built-in LED on pin 13 of the Arduino Uno.

  Concepts learned:
    - pinMode(), digitalWrite(), delay()
*/

const int ledPin = 13;   // the pin the LED is connected to

void setup() {
  pinMode(ledPin, OUTPUT);   // set the LED pin as an output
}

void loop() {
  digitalWrite(ledPin, HIGH);  // turn the LED on
  delay(500);                  // wait half a second
  digitalWrite(ledPin, LOW);   // turn the LED off
  delay(500);                  // wait half a second
}
