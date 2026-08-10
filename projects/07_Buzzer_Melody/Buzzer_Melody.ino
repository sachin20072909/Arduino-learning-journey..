/*
  Project 07 - Buzzer Melody Player
  ---------------------------------
  Play a simple tune on a piezo buzzer using tone().

  Wiring:
    - Buzzer "+" -> pin 3 through a 220 ohm resistor
    - Buzzer "-" -> GND

  Concepts learned:
    - tone(pin, frequency) produces a square wave at a pitch
    - arrays hold the melody notes and their durations
    - frequency = how fast the pin toggles = pitch we hear
*/

const int buzzerPin = 3;

// Notes of a short melody (frequencies in Hz)
int melody[] = {
  262, 294, 330, 349, 392, 440, 494, 523   // C D E F G A B C
};

// How long each note plays (in ms)
int durations[] = {
  300, 300, 300, 300, 300, 300, 300, 600
};

int noteCount = 8;

void setup() {
  for (int i = 0; i < noteCount; i++) {
    tone(buzzerPin, melody[i], durations[i]);  // play note, auto-stop
    delay(durations[i] + 80);                  // wait + small gap
  }
  noTone(buzzerPin);   // make sure the buzzer is silent at the end
}

void loop() {
  // Melody plays once at power-up; nothing to do here.
}
