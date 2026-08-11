"""
Parkwise AI - Serial Dashboard
------------------------------
Reads CSV telemetry from the Arduino over USB and displays a live
bay-status dashboard in the terminal.

Setup:
    pip install pyserial
Run:
    python dashboard.py <COM_PORT>        # e.g. COM3 on Windows
    python dashboard.py /dev/ttyACM0      # Linux
    python dashboard.py /dev/tty.usbmodemXXXX  # macOS

What you'll see:
    - A colored status box per bay (green = FREE, red = OCCUPIED)
    - Live distance reading
    - Timestamped state-change log

This is the cheapest possible "AI dashboard" for a hackathon demo.
Replace prints with a Flask/FastAPI + websocket page later if you want
a browser UI.
"""

import sys
import time
from datetime import datetime

try:
    import serial
except ImportError:
    sys.exit("pyserial not installed. Run: pip install pyserial")


def color(text, code):
    return f"\033[{code}m{text}\033[0m"


RED_BG    = "41"
GREEN_BG  = "42"
YELLOW    = "33"


class BayView:
    def __init__(self, bay_id: int):
        self.bay_id = bay_id
        self.state  = "UNKNOWN"
        self.dist   = -1

    def update(self, state: str, dist: int):
        self.state = state
        self.dist  = dist

    def render(self) -> str:
        if self.state == "FREE":
            label = color(" FREE ", f"1;37;{GREEN_BG}")
        elif self.state == "OCCUPIED":
            label = color(" OCCUPIED ", f"1;37;{RED_BG}")
        else:
            label = color(" UNKNOWN ", f"1;30;{YELLOW}")
        return f"Bay {self.bay_id}: {label}   distance: {self.dist:>4} cm"


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    port = sys.argv[1]
    baud = 9600

    print(f"Connecting to {port} @ {baud}...")
    ser = serial.Serial(port, baud, timeout=1)
    time.sleep(2)  # Uno resets on serial open

    bays: dict[int, BayView] = {}

    # Clear screen
    print("\033[2J\033[H", end="")
    print(color("=== Parkwise AI - Live Dashboard ===", "1;36"))
    print("Press Ctrl+C to stop.\n")

    try:
        while True:
            raw = ser.readline().decode("utf-8", errors="replace").strip()
            if not raw:
                continue
            parts = raw.split(",")

            if parts[0] == "BAY" and len(parts) >= 4:
                bay_id = int(parts[1])
                state  = parts[2]
                dist   = int(parts[3])
                bays.setdefault(bay_id, BayView(bay_id)).update(state, dist)

            elif parts[0] == "STATE_CHANGE" and len(parts) >= 4:
                ts = datetime.now().strftime("%H:%M:%S")
                print(color(f"[{ts}] Bay {parts[2]} -> {parts[3]}", YELLOW))
                continue

            # Redraw bay status block (simple terminal UI)
            print("\033[H", end="")           # cursor home
            print(color("=== Parkwise AI - Live Dashboard ===", "1;36"))
            print(f"Last message: {raw[:60]:<60}")
            print("-" * 52)
            for bid in sorted(bays):
                print(bays[bid].render())
            print("-" * 52)

    except KeyboardInterrupt:
        print("\nDisconnecting...")
    finally:
        ser.close()


if __name__ == "__main__":
    main()
