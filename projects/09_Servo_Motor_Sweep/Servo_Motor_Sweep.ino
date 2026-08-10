/*
  Project 09 - Servo Motor Sweep
  ------------------------------
  Sweep a servo motor back and forth from 0 to 180 degrees.

  Wiring (SG90 micro servo):
    - Brown wire  -> GND
    - Red wire    -> 5V
    - Orange wire -> pin 5

  Concepts learned:
    - using a library with #include
    - servo.write(angle) moves the servo to a position
    - a for-loop to create smooth motion
*/

#include <Servo.h>   // built into the Arduino IDE

Servo myServo;       // create a servo object

const int servoPin = 5;

void setup() {
  myServo.attach(servoPin);   // tell the library which pin to use
}

void loop() {
  // Sweep from 0 to 180 degrees
  for (int angle = 0; angle <= 180; angle++) {
    myServo.write(angle);
    delay(15);                 // wait for the servo to reach the position
  }

  // Sweep back from 180 to 0 degrees
  for (int angle = 180; angle >= 0; angle--) {
    myServo.write(angle);
    delay(15);
  }
}
