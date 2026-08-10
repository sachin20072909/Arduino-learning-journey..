/*
  Project 05 - RGB LED Colour Mixer
  ---------------------------------
  Cycle an RGB LED through different colours by mixing red, green and blue.

  Wiring (common-cathode RGB LED):
    - Longest pin (cathode) -> GND
    - Red pin    -> pin 11 through a 220 ohm resistor
    - Green pin  -> pin 10 through a 220 ohm resistor
    - Blue pin   -> pin 9  through a 220 ohm resistor

  Concepts learned:
    - PWM mixing: three analogWrite() values make a colour
    - a reusable setColor() function
*/

const int redPin   = 11;
const int greenPin = 10;
const int bluePin  = 9;

void setup() {
  pinMode(redPin, OUTPUT);
  pinMode(greenPin, OUTPUT);
  pinMode(bluePin, OUTPUT);
}

void loop() {
  setColor(255, 0, 0);    // red
  delay(1000);
  setColor(0, 255, 0);    // green
  delay(1000);
  setColor(0, 0, 255);    // blue
  delay(1000);
  setColor(255, 255, 0);  // yellow (red + green)
  delay(1000);
  setColor(0, 255, 255);  // cyan (green + blue)
  delay(1000);
  setColor(255, 0, 255);  // magenta (red + blue)
  delay(1000);
  setColor(255, 255, 255);// white (all on)
  delay(1000);
}

void setColor(int r, int g, int b) {
  analogWrite(redPin, r);
  analogWrite(greenPin, g);
  analogWrite(bluePin, b);
}
