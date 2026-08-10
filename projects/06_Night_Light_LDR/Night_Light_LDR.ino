/*
  Project 06 - Light Sensing Night Light
  --------------------------------------
  An LED turns on automatically when the room gets dark, using an LDR
  (light dependent resistor). Like a real street light!

  Wiring:
    - LDR: one leg -> 5V, other leg -> A1 AND a 10k resistor to GND
      (the LDR + resistor form a voltage divider)
    - LED anode -> pin 6 through a 220 ohm resistor -> GND

  Concepts learned:
    - voltage dividers with sensors
    - thresholds: deciding "dark" vs "light"
    - printing to the Serial Monitor to debug sensor values
*/

const int ldrPin = A1;
const int ledPin = 6;
const int darkThreshold = 400;   // adjust based on your room lighting

void setup() {
  pinMode(ledPin, OUTPUT);
  Serial.begin(9600);            // start serial so we can see readings
}

void loop() {
  int lightLevel = analogRead(ldrPin);
  Serial.print("Light level: ");
  Serial.println(lightLevel);

  if (lightLevel < darkThreshold) {
    digitalWrite(ledPin, HIGH);  // it's dark -> turn the light on
  } else {
    digitalWrite(ledPin, LOW);   // it's bright -> turn it off
  }

  delay(200);
}
