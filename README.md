# Arduino-learning-journey..

It's my journey of learning in my initial stages in electronics.

This repository collects small Arduino projects I build while learning
electronics. Each project lives in its own folder under `projects/`, with
comments in the code explaining the wiring and the concepts learned.

## Projects

| # | Project | What it does | Key concepts |
|---|---------|--------------|--------------|
| 0 | [Fading of LED](Fading_of_LED.ino) | Smoothly fades an LED in and out | `analogWrite()`, PWM |
| 01 | [Blinking LED](projects/01_Blinking_LED) | The classic "Hello World" of Arduino | `digitalWrite()`, `delay()` |
| 02 | [Traffic Light](projects/02_Traffic_Light) | Red, yellow, green LEDs cycle like a signal | multiple outputs, helper functions |
| 03 | [Push Button LED](projects/03_Push_Button_LED) | LED lights while a button is held | `digitalRead()`, `INPUT_PULLUP` |
| 04 | [Potentiometer LED](projects/04_Potentiometer_LED) | Knob controls LED brightness | `analogRead()`, `map()` |
| 05 | [RGB LED Colour Mixer](projects/05_RGB_LED_Colour_Mixer) | Cycles an RGB LED through colours | PWM colour mixing |
| 06 | [Night Light (LDR)](projects/06_Night_Light_LDR) | LED turns on automatically in the dark | voltage dividers, thresholds, Serial Monitor |
| 07 | [Buzzer Melody](projects/07_Buzzer_Melody) | Plays a tune on a piezo buzzer | `tone()`, arrays |
| 08 | [Ultrasonic Parking Sensor](projects/08_Ultrasonic_Parking_Sensor) | Measures distance and beeps faster when closer | HC-SR04, `pulseIn()`, timing |
| 09 | [Servo Motor Sweep](projects/09_Servo_Motor_Sweep) | Sweeps a servo 0°–180° | libraries, `Servo.h` |
| 10 | [Knock Detector](projects/10_Knock_Detector) | LED lights when the desk is knocked | piezo as sensor, `millis()` |

## Parts used across the projects

- Arduino Uno (or compatible board)
- Breadboard + jumper wires
- LEDs: red, yellow, green, one RGB common-cathode LED
- Resistors: 220 ohm and 10k
- Push button
- 10k potentiometer
- LDR (light dependent resistor)
- Piezo buzzers x2 (one as sound, one as vibration sensor)
- HC-SR04 ultrasonic sensor
- SG90 micro servo

## How to run a project

1. Open the `.ino` file in the [Arduino IDE](https://www.arduino.cc/en/software)
   (or Arduino Web Editor).
2. Build the circuit exactly as described in the comment block at the top
   of the sketch.
3. Choose your board and port under **Tools**, then click **Upload**.

More projects coming as the journey continues...
