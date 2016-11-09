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
SCOPES = 'https://www.googleapis.com/auth/calendar'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Google Calendar API Python Quickstart'


def xls_to_list(xls_path):
    """
    Converts the xls file at the given path and writes out a csv file
    :param xls_path the input xls file path
    """
    workbook = xlrd.open_workbook(xls_path)
    worksheet = workbook.sheet_by_index(1)
    sheet = []
    for rownum in xrange(worksheet.nrows):
        row = []
        for cell in worksheet.row_values(rownum):
            if type(cell) == type(u''):
                row.append(cell.encode('utf-8'))
            elif type(cell) == type(0.1) and cell > 0:
                date = xlrd.xldate_as_tuple(cell, datemode=1)
                if date[0] == 0:
                    row.append("{0}:{1}".format(date[3], date[4]))
                else:
                    row.append("{0}-{1}-{2}".format(datetime.today().year, date[1], date[2] - 1))
            else:
                row.append(cell)
        sheet.append(row)

    return sheet


def sheet_to_shifts(sheet):
    shifts = []
    dates = []
    for i, row in enumerate(sheet):
        if row[3] == 'Sunday':
            dates = sheet[i + 1]
        if row[1] == 'Meg':
            time_in = sheet[i]
            time_out = sheet[i + 1]
            prev_time = None
            for i, date in enumerate(dates):
                shift_num = 'first' if date else ('second' if prev_time else None)
                if time_in != '' and shift_num:
                    shift = (date if date != '' else dates[i - 1], time_in[i], time_out[i])
                    if shift[1] != '' and shift[1] != '0:15':
                        shifts.append(shift)
                prev_time = date if date != '' else None

    return shifts


def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'calendar-python-quickstart.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        # if flags:
        #     credentials = tools.run_flow(flow, store, flags)
        # else: # Needed only for compatibility with Python 2.6
        credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials


def add_shifts_to_calendar(shifts):
    # Refer to the Python quickstart on how to setup the environment:
    # https://developers.google.com/google-apps/calendar/quickstart/python
    # Change the scope to 'https://www.googleapis.com/auth/calendar' and delete any
    # stored credentials.

    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('calendar', 'v3', http=http)

    for shift in shifts:
        print('shift\n{0}'.format(shift))
        start_hour = shift[1].split(':')[0]
        start_min = shift[1].split(':')[1]
        end_hour = shift[2].split(':')[0]
        end_min = shift[2].split(':')[1]
        event = {
          'summary': 'Meg Working',
          'start': {
            'dateTime': '{0}T{1}:{2}:00-05:00'.format(shift[0], start_hour, start_min if start_min != '0' else '00')
          },
          'end': {
            'dateTime': '{0}T{1}:{2}:00-05:00'.format(shift[0], end_hour, end_min if end_min != '0' else '00')
          }
        }

        print('EVENT {0}'.format(event))

        meg_calendar_id = 'sp2kd7vp3rnrst975s0j443kh4@group.calendar.google.com'
        event = service.events().insert(calendarId=meg_calendar_id, body=event).execute()
        print('Event created: %s' % (event.get('htmlLink')))


if __name__ == "__main__":
    xls_path = sys.argv[1]
    sheet = xls_to_list(xls_path=xls_path)
    shifts = sheet_to_shifts(sheet=sheet)
    add_shifts_to_calendar(shifts)
