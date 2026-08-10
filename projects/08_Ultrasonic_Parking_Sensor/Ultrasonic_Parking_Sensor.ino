/*
  Project 08 - Ultrasonic Parking Sensor
  --------------------------------------
  Measure distance with an HC-SR04 ultrasonic sensor. A buzzer beeps
  faster the closer an object gets — like a car's parking sensor.

  Wiring:
    - HC-SR04: VCC -> 5V, GND -> GND, Trig -> pin 10, Echo -> pin 9
    - Buzzer "+" -> pin 8, buzzer "-" -> GND

  Concepts learned:
    - how the sensor works: send a pulse, time the echo
    - distance = time x speed of sound
    - mapping distance to behaviour
*/

const int trigPin   = 10;
const int echoPin   = 9;
const int buzzerPin = 8;

void setup() {
  pinMode(trigPin, OUTPUT);
  pinMode(echoPin, INPUT);
  pinMode(buzzerPin, OUTPUT);
  Serial.begin(9600);
}

void loop() {
  // 1. Send a 10 microsecond pulse from Trig
  digitalWrite(trigPin, LOW);
  delayMicroseconds(2);
  digitalWrite(trigPin, HIGH);
  delayMicroseconds(10);
  digitalWrite(trigPin, LOW);

  // 2. Time how long the Echo stays high (in microseconds)
  long duration = pulseIn(echoPin, HIGH);

  // 3. Convert to cm: sound travels ~0.0343 cm/us, out and back (so /2)
  long distance = duration * 0.0343 / 2;

  Serial.print("Distance: ");
  Serial.print(distance);
  Serial.println(" cm");

  // 4. Beep based on distance
  if (distance < 10) {
    tone(buzzerPin, 1000, 60);     // very close -> rapid beeping
    delay(100);
  } else if (distance < 30) {
    tone(buzzerPin, 800, 120);     // getting close
    delay(400);
  } else if (distance < 60) {
    tone(buzzerPin, 600, 150);     // far-ish
    delay(800);
  } else {
    noTone(buzzerPin);             // far enough -> silent
    delay(500);
  }
}
