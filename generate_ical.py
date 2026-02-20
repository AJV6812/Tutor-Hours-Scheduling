import hashlib
import uuid
import shutil
import logging
import os
from pathlib import Path
import dateutil
import pytz
import sys

import googleapiclient
from icalendar import Calendar, Event, vText
from auth import auth_with_sheets_api


def main():
    # Logging and configuration initialisation
    config = {}
    config_path = os.environ.get("CONFIG_PATH", "config.py")
    exec(Path(config_path).read_text(), config)

    logger = logging.getLogger(__name__)
    logger.setLevel(logging.DEBUG)
    if config.get("LOGFILE", None):
        handler = logging.FileHandler(filename=config["LOGFILE"], mode="a")
    else:
        handler = logging.StreamHandler(sys.stderr)
    handler.setFormatter(logging.Formatter("%(asctime)s|[%(levelname)s] %(message)s"))
    logger.addHandler(handler)

    service = auth_with_sheets_api(config)
    sheet = service.spreadsheets().values()
    logger.info("> Gained access to Google API")

    # Going through the name of each sheet and getting a 2D list of values corresponding to the values in that sheet
    worksheets = {}
    for worksheet_name in config["WORKSHEET_NAMES"]:
        try:
            worksheets[worksheet_name] = (
                sheet.get(spreadsheetId=config["SPREADSHEET_ID"], range=worksheet_name)
                .execute()
                .get("values", [])
            )
        except googleapiclient.errors.HttpError as e:
            logger.error(f"Error in accessing sheet {worksheet_name}. {e}")
            continue

    logger.debug(f"> Loaded {len(worksheets)} sheets")

    # Initialising calendars for each sheet and one master calendar
    individual_calendars = {
        worksheet_name: Calendar() for worksheet_name in config["WORKSHEET_NAMES"]
    }
    master_calendar = Calendar()

    for worksheet_name, values in worksheets.items():
        header_row = values[0]
        try:
            # Look for which column each header is in for this particular spreadsheet
            date_index = header_row.index(config["DATE_HEADER"])
            start_time_index = header_row.index(config["START_TIME_HEADER"])
            end_time_index = header_row.index(config["END_TIME_HEADER"])
            location_index = header_row.index(config["LOCATION_HEADER"])
            type_index = header_row.index(config["TYPE_HEADER"])
            name_index = header_row.index(config["NAME_HEADER"])
        except ValueError as e:
            logger.error(f"> Could not read headers in {worksheet_name} as {e}")
            continue

        for row in values[1:]:
            event: Event = Event()

            name = f"{worksheet_name}'s {row[type_index]}" # For regular tutor hours
            if row[name_index] != "":
                name = f"{row[name_index]} ({name})" # For seminar naming
            
            if not row[start_time_index] or not row[date_index] or not row[end_time_index] or not type_index:
                logger.error(f"> Could not interpret event: {name} as some critical fields were empty.")
                continue

            event.add("summary", name) # Summary corresponds to the title shown in Google Calendar.

            # Using dateutil.parser.parse to read in the time and date. If this does not work for a specific case, you can switch to a rigid parser, or set fuzzy to True.
            try:
                start_dt = dateutil.parser.parse(
                    f"{row[date_index]} {row[start_time_index]}"
                )
                end_dt = dateutil.parser.parse(
                    f"{row[date_index]} {row[end_time_index]}"
                )
            except dateutil.parser._parser.ParserError as e:
                logger.error(f"> Could not process time in {name} ({e})")
                continue

            # Changes the datetime object to set it to the Sydney timezone
            start_dt = pytz.timezone("Australia/Sydney").localize(start_dt)
            end_dt = pytz.timezone("Australia/Sydney").localize(end_dt)

            event.add("dtstart", start_dt)
            event.add("dtend", end_dt)
            event["location"] = vText(row[location_index])

            logger.info(f"> Processed {name} on {row[date_index]}")

            # Here we set a uid based on the hash of the event string.
            # This means that if the event is exactly the same, the uids will match and it will not be deleted and readded by Google Calendar.
            # If they are different at all, then ical_to_gcal will notice the different uids, and update the calendar.
            seed = event.to_ical()
            m = hashlib.md5()
            m.update(seed)
            new_uuid = uuid.UUID(m.hexdigest())
            event.add("uid", new_uuid)

            # Add the event to both the individual calendar, and the master calendar
            individual_calendars[worksheet_name].add_component(event)
            master_calendar.add_component(event)

    # Save each calendar in a separate folder (to make ical_to_gcal_sync work properly)
    shutil.rmtree("./icals")
    for name, calendar in individual_calendars.items():
        os.makedirs(f"./icals/{name}")
        with open(f"./icals/{name}/cal.ics", "wb") as f:
            f.write(calendar.to_ical())

    os.makedirs("./icals/master")
    with open("./icals/master/cal.ics", "wb") as f:
        f.write(master_calendar.to_ical())
    logger.info("> Converted spreadsheet into icals")
