#!/bin/python

import sys
import xlrd
from datetime import date as datetime
import httplib2
import os
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage
from apiclient import discovery

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/calendar-python-quickstart.json
SCOPES = "https://www.googleapis.com/auth/calendar"
CLIENT_SECRET_FILE = "client_secret.json"
APPLICATION_NAME = "Google Calendar API Python Quickstart"
MEG_CALENDAR_ID = "sp2kd7vp3rnrst975s0j443kh4@group.calendar.google.com"


class Shift:
    def __init__(self, date, start_time, end_time):
        self.date = date

        self.start_hour = start_time["hour"]
        self.start_minute = start_time["minute"] if start_time["minute"] else "00"

        self.end_minute = end_time["minute"] if end_time["minute"] else "00"
        self.end_hour = end_time["hour"]

        self.is_off = self.start_hour == 0 and self.start_minute == 15

    def start_shift(self):
        return "{0.date}T{0.start_hour}:{0.start_minute}".format(self)

    def end_shift(self):
        return "{0.date}T{0.end_hour}:{0.end_minute}".format(self)

    def __repr__(self):
        return "{0.date}  {0.start_hour}:{0.start_minute}  {0.end_hour}:{0.end_minute}".format(self)


def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser("~")
    credential_dir = os.path.join(home_dir, ".credentials")
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir, "calendar-python-quickstart.json")

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        credentials = tools.run(flow, store)
        print("Storing credentials to " + credential_path)
    return credentials


def xls_to_list(xls_path):
    """
    Converts the xls file at the given path to a list of rows

    :param xls_path: the input xls file path
    :return a list each containing a row (list of cells) from the xls sheet
        special cells are:
            date rows formatted <year>-<month>-<day>
            time rows formatted <hour>:<minutes>
    """
    workbook = xlrd.open_workbook(xls_path)
    worksheet = workbook.sheet_by_index(1)
    sheet = []
    for rownum in xrange(worksheet.nrows):
        row = []
        for cell in worksheet.row_values(rownum):
            if isinstance(cell, str):
                row.append(cell.encode("utf-8"))
            elif isinstance(cell, float) and cell > 0:
                (year, month, day, hour, minute, second) = xlrd.xldate_as_tuple(cell, datemode=1)
                if year == 0:
                    # It's a time
                    row.append({"hour": hour, "minute": minute})
                else:
                    # It's a date
                    row.append("{0}-{1}-{2}".format(datetime.today().year, month, day - 1))
            else:
                row.append(cell)
        sheet.append(row)
    return sheet


def sheet_to_shifts(sheet):
    """
    Converts a list of rows from a schedule sheet to a list of Shift objects

    Rows are formatted:
      0  1   2     3       4     5     6       7     8  ...
     [ ,    , , Sunday,         , ,  Monday,        , , ...]
     [ ,    , ,   date,         , ,    date,        , , ...]
      ...
     [ , Meg, , shift 1, shift 2, , shift 1, shift 2, , ...] (start times)
     [ ,    , , shift 1, shift 2, , shift 2, shift 2, , ...] (end times)

    :param sheet: a list of rows from the wylie wagg schedule sheet
    :return: a list of Shift objects
    """
    shifts = []
    dates = []
    for i, row in enumerate(sheet):
        if row[3] == "Sunday":
            # This is the line above a date line
            dates = sheet[i + 1]
        if row[1] == "Meg":
            # This line contains the shift start times
            # The next line contains the shift end times
            start_time = sheet[i]
            end_time = sheet[i + 1]
            prev_time = None
            for date_index, date in enumerate(dates):
                shift_cell = date or prev_time
                if start_time[date_index] and shift_cell:
                    shift = Shift(date if date else dates[date_index - 1],
                                  start_time[date_index], end_time[date_index])
                    if not shift.is_off:
                        shifts.append(shift)
                prev_time = date if date != "" else None

    return shifts


def add_shifts_to_calendar(shifts, dry_run=False):
    """
    Adds the shifts to the google calendar via the google calendar API

    :param shifts: a list of shift tuples (year, start time, end time)
    :param dry_run: False if events should be added using the google calendar API, True if a dry run with just prints
    """
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build("calendar", "v3", http=http)

    for shift in shifts:
        event = {
            "summary": "Meg Working",
            "start": {"dateTime": "{}:00-05".format(shift.start_shift())},
            "end": {"dateTime": "{}:00-05".format(shift.end_shift())}
        }

        if not dry_run:
            created_event = service.events().insert(calendarId=MEG_CALENDAR_ID, body=event).execute()
            print("Event created: {} {}".format(event, created_event.get("htmlLink")))
        else:
            print("Dry Run Event: {}".format(event))


if __name__ == "__main__":
    # TODO: Use argparse or some shit
    xls_path = sys.argv[1]
    dry_run = len(sys.argv) < 3 or not sys.argv[2] == '--create'

    sheet = xls_to_list(xls_path)
    shifts = sheet_to_shifts(sheet)
    add_shifts_to_calendar(shifts, dry_run)
