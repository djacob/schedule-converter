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
        (self.start_hour, self.start_minute) = start_time.split(":")
        (self.end_hour, self.end_minute) = end_time.split(":")

        self.start_minute = self.start_minute if self.start_minute != "0" else "00"
        self.end_minute = self.end_minute if self.end_minute != "0" else "00"

        self.is_off = start_time == "0:15"

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
    Converts the xls file at the given path and writes out a csv file

    :param xls_path: the input xls file path
    :return a list each containing a row from the xls sheet
        date rows are formatted <year>-<month>-<day>
        time rows are formatted <hour>:<minutes>
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
                    row.append("{0}:{1}".format(hour, minute))
                else:
                    row.append("{0}-{1}-{2}".format(datetime.today().year, month, day - 1))
            else:
                row.append(cell)
        sheet.append(row)
    return sheet


def sheet_to_shifts(sheet):
    """
    Converts a list of rows from a schedule sheet to a list of shifts

    :param sheet: a list of rows from the wylie wagg schedule sheet
    :return: a list of shift tuples in the form (<year>, <start time>, <end time>)
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
                empty_cell = not date and not prev_time
                if start_time[date_index] != "" and not empty_cell:
                    shift = Shift(date if date != "" else dates[date_index - 1],
                                  start_time[date_index], end_time[date_index])
                    if not shift.is_off:
                        shifts.append(shift)
                prev_time = date if date != "" else None

    return shifts


def add_shifts_to_calendar(shifts):
    """
    Adds the shifts to the google calendar via the google calendar API

    :param shifts: a list of shift tuples (year, start time, end time)
    """
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build("calendar", "v3", http=http)

    for shift in shifts:
        event = {
            "summary": "Meg Working",
            "start": {
                "dateTime": "{}-05:00".format(shift.start_shift())
            },
            "end": {
                "dateTime": "{}-05:00".format(shift.end_shift())
            }
        }

        created_event = service.events().insert(calendarId=MEG_CALENDAR_ID, body=event).execute()
        print("Event created: {} {}".format(event, created_event.get("htmlLink")))


if __name__ == "__main__":
    xls_path = sys.argv[1]
    sheet = xls_to_list(xls_path=xls_path)
    shifts = sheet_to_shifts(sheet=sheet)
    add_shifts_to_calendar(shifts)
