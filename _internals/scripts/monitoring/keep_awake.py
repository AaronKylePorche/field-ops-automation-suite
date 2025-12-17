# keep_awake.py
# Keeps Windows from sleeping (and keeps the display awake) while this runs.
# Usage:
#   python keep_awake.py
# Press Ctrl+C to stop (sleep/display holds are released).

import time
try:
    from ctypes import windll
except Exception:
    windll = None

ES_CONTINUOUS       = 0x80000000
ES_SYSTEM_REQUIRED  = 0x00000001  # prevent system sleep
ES_DISPLAY_REQUIRED = 0x00000002  # keep display on
ES_AWAYMODE_REQUIRED= 0x00000040  # optional; ignored on some editions

def _prevent_sleep_and_display():
    if windll:
        try:
            # Hold system + display awake continuously
            windll.kernel32.SetThreadExecutionState(
                ES_CONTINUOUS | ES_SYSTEM_REQUIRED | ES_DISPLAY_REQUIRED | ES_AWAYMODE_REQUIRED
            )
        except Exception:
            pass

def _allow_sleep():
    if windll:
        try:
            windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS)
        except Exception:
            pass

def main():
    print("Keep-awake: ON (system + display held awake)")
    _prevent_sleep_and_display()
    try:
        while True:
            # Refresh every 30s to be extra safe
            time.sleep(30)
            _prevent_sleep_and_display()
    except KeyboardInterrupt:
        pass
    finally:
        _allow_sleep()
        print("Keep-awake: OFF")

if __name__ == "__main__":
    main()
