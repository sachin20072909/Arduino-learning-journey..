/*
  Project 03 - Push Button Controlled LED
  ---------------------------------------
  Turn an LED on only while a push button is held down.

  Wiring:
    - Push button: one side -> pin 7, other side -> GND
      (we use the Arduino's built-in pull-up resistor, no external
      resistor needed)
    - LED anode -> pin 8 through a 220 ohm resistor -> GND

  Concepts learned:
    - digitalRead()
    - INPUT_PULLUP (the internal pull-up resistor)
    - button logic: with INPUT_PULLUP the pin reads LOW when pressed
*/

const int buttonPin = 7;
const int ledPin    = 8;

void setup() {
  pinMode(ledPin, OUTPUT);
  pinMode(buttonPin, INPUT_PULLUP);  // enable internal pull-up resistor
}

void loop() {
  // With INPUT_PULLUP the pin is HIGH when released, LOW when pressed
  if (digitalRead(buttonPin) == LOW) {
    digitalWrite(ledPin, HIGH);   // button pressed -> LED on
  } else {
    digitalWrite(ledPin, LOW);    // button released -> LED off
  }
}
