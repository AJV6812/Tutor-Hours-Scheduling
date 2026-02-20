# This is the main execution point for the program. It basically just executes generate_ical and ical_to_gcal_sync and displays how many more error, debug and info messages there are in the log as a result.

import os
from pathlib import Path
import generate_ical
import ical_to_gcal_sync

config = {}
config_path = os.environ.get("CONFIG_PATH", "config.py")
exec(Path(config_path).read_text(), config)

current_log = ""
if os.path.exists(config["LOGFILE"]):
    with open(config["LOGFILE"], "r", encoding="utf-8") as f:
        current_log = f.read()

info, error, debug = (
    current_log.count("[INFO]"),
    current_log.count("[ERROR]"),
    current_log.count("[DEBUG]"),
)

generate_ical.main()
ical_to_gcal_sync.main()

with open(config["LOGFILE"], "r", encoding="utf-8") as f:
    current_log = f.read()

new_info, new_error, new_debug = (
    current_log.count("[INFO]"),
    current_log.count("[ERROR]"),
    current_log.count("[DEBUG]"),
)

print(
    f"Generated {new_error-error} error message(s), {new_debug-debug} debug messages and {new_info-info} info messages."
)
print(f"See {config['LOGFILE']} for more information.")
