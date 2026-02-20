# New College Tutoring Auto-Calendar Script

This is the script used to transfer tutor hours from a Google Spreadsheet into a Google Calendar.

It might look like a lot but it should be doable to get this set up in about an hour, and after which never have to look at it for the rest of the year!

It's strongly recommended that the tutor who sets this up is someone who has done COMP1531 (or otherwise has experience with git and Python). Even better if there's anyone who already knows how OAuth and Google API works, but if not this guide should still be clear enough for the code to be functional.

If something in the guide is unclear, you may be able to find additional help here:
- https://github.com/andrewramsay/ical_to_gcal_sync?tab=readme-ov-file
- https://developers.google.com/workspace/calendar/api/quickstart/python
- https://developers.google.com/workspace/sheets/api/quickstart/python
And you can see an example of how a tutor hours spreadsheet could be laid out here:
- https://docs.google.com/spreadsheets/d/1mK9tzrYF8xmd1Ncdhn9qOw81cKFGS-qQVlkeK0hRkSo/edit?usp=sharing


This guide assumes you have a modern version of Python set up already (including pip) and are using VSCode.

Firstly, setting we'll set up the code base:

1. Download all the source files into a directory. Open VS Code in this directory.
2. Windows + Shift + P (or otherwise open the VS Code command palate), and select "Python: Create Environment"
3. Select venv, pick a 3.12.3 python version (other versions will probably still work)
4. Leave .venv as the name
5. Select install project dependencies, and use requirements.txt as your depedencies.

Basically all this has done is created a separate space in which Python runs in, and installed all the required modules into that space. At this point the python code itself is ready to run, but we need to set up the Google sheets file and Google Calendar first.

From this point onwards, I'll call a Google Sheets file a workbook, and an individual sheet inside it a sheet.

6. Create a workbook (call it whatever you like), and inside that sheet make individual sheets for each tutor (Ensure the name of each sheet is only the name of the tutor, the name of the sheet is how each tutor is distinguished). Feel free to make additional sheets, we will later configure the program so only sheets with tutor names are recognised by the program (for example you may want to make an extra sheet with helpful information about term dates to help with scheduling).
7. Inside each tutor's sheet, create headings (**Important**: the headings should be in row 1, not further down) for the fields you want to keep track of. Currently, the program will track Date, Start Time, End Time, Location and Type, though it's fairly easy to add or remove headings in the code if required. Similarly, if you add extra headings (say for week or day of week), without modifying the code, the program will simply ignore those headings. The Name header is intended for the names of special events, such as seminars. If the Name is blank, the calendar will show [Tutor Name]'s [Event Type] as the name.

The spreadsheet is now ready to be read by the program. Next we will get Google API configured so that the program has permission to read the Google Sheet and modify the Calendar. I'm basically following these two guides, to set up the workstation:
- https://developers.google.com/workspace/sheets/api/quickstart/python 
- https://developers.google.com/workspace/calendar/api/quickstart/python.

The two guides are practically identical, the only real difference is the link to which API to enable.

8. Create a [Google Cloud Project](https://developers.google.com/workspace/guides/create-project). The project Name and ID I used are Tutor Hours Scheduling and tutor-hours-scheduling respectively.
9. We enable the [Google Sheets API](https://console.cloud.google.com/flows/enableapi?apiid=calendar-json.googleapis.com) and [Google Calendar API](https://console.cloud.google.com/flows/enableapi?apiid=sheets.googleapis.com) (ensure the correct project is selected in the top left if you have multiple).
10. Follow the instructions of the above guides for Configuring OAuth consent. **Important**: If it asks you to create a client, ensure that client is a Desktop Application not Webapp or any other.
11. Once you've created the client, the next pop-up will ask whether you want to download the credentials, make sure to download them as `account_info.json` and put them in your working directory.
12. Under the [audience page](https://console.cloud.google.com/auth/audience), ensure to add your own email as a test user.

Note that in this configuration, you will have to log in through the app every 7 days. If this doesn't bother you, then leave it, but if you want to automate the system completely, you should go to "https://console.cloud.google.com/auth/" and put the application into production.

Now we can set up Google Calendar and the configuration file for the program.

13. In Google Calendar, create calendar's for each tutor, and one master tutoring calendar (name of the calendar doesn't really matter). Go into settings for each calendar, you will likely want to make each one public. Take note of the google calendar ID found under "Integrate calendar" in the calendar settings. It will look something like this: `8a14f6e116f203318fcf5ff3e815c7743d8762fea64eaec0e6983e5d698949c0@group.calendar.google.com.`
14. In `config.py`, ICAL_FIELDS should represent a list of .ics files that the program will load into Google Calendar. Ensure for each calendar created in the previous step, there is a separate dictionary for each one - each with three fields, 'source' (which is `./icals/master` for the master tutor calendar and `./icals/TutorName` for each of the tutor calendars), 'destination' (which is the Google Calendar ID noted in the previous step) and 'files' (which should be True).
15. Ensure `APPLICATION_NAME` matches the application ID for this project (if you used the same as mine it'll be tutor-hours-scheduling).
16. Set `SPREADSHEET_ID` to ID found in the URL of the spreadsheet. For example in https://docs.google.com/spreadsheets/d/1CtLZRSk1SFdRVYnkbzy2fmLWEUF2P9S2utjD52YkW2Q/edit?gid=1712255416#gid=1712255416, `1CtLZRSk1SFdRVYnkbzy2fmLWEUF2P9S2utjD52YkW2Q` is the ID.
17. Set the worksheet names to the names of each sheet you want to read events from (most likely ["Tutor1", "Tutor2", ...]). If you used different header names in each sheet of the workbook, change the header strings to match accordingly (be careful as they're case sensitive). For example if in the spreadsheet the header you used is "Starting Time" instead of "Start Time", change START\_TIME\_HEADER to "Starting Time".

We're finally ready to run the program.

18. Run main.py, which should open a website, then sign into your Google Account. (There may be some warnings, but we can ignore those). Once you've given permission, we can return to VSCode.
19. You can watch the progress of the program by opening tutor_calendar.log.
20. In future, run main.py periodically to update the calendar, or set up a cronjob or something similar to have it run once a day in the background.
21. You can share the Google Calendars with others by going into the calendar settings, integrate calendar and sharing either the URL or the ICAL link.

Additional documentation that may be helpful if you wish to modify the code for a specific purpose:
- https://icalendar.readthedocs.io/en/stable/reference/api/icalendar.cal.event.html
- https://larrybolt.github.io/online-ics-feed-viewer/https://larrybolt.github.io/online-ics-feed-viewer/
- https://github.com/andrewramsay/ical_to_gcal_sync?tab=readme-ov-file
- https://icalendar.readthedocs.io/en/stable/how-to/usage.html