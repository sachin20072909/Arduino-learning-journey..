/*
  Project 0 - Fading of LED
  -------------------------
  Smoothly fades an LED in and out using PWM (Pulse Width Modulation).

  Wiring:
    - LED anode (long leg) -> pin 9 (a PWM pin) through a 220 ohm resistor -> GND
    - LED cathode (short leg) -> GND

  Concepts learned:
    - analogWrite() expects values from 0 to 255
    - PWM (Pulse Width Modulation) for controlling LED brightness
*/

const int led = 9;           // the PWM pin the LED is attached to

void setup() {
  pinMode(led, OUTPUT);      // declare pin 9 to be an output
}

void loop() {
  // fade in from min to max in increments of 5 points:
  for (int f = 0; f <= 255; f += 5) {
    analogWrite(led, f);
    delay(30);
  }
  
  // fade out from max to min in increments of 5 points:
  for (int f = 255; f >= 0; f -= 5) {
    analogWrite(led, f);
    delay(30);
  }
}
