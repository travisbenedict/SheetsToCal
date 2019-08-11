from __future__ import print_function
import httplib2
import os
import re
import json

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage
import dateutil.parser as dparser

import gspread
from oauth2client.service_account import ServiceAccountCredentials

import datetime

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/calendar-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/calendar'
CLIENT_SECRET_FILE = 'calendar_secret.json'
APPLICATION_NAME = 'Google Calendar API Python Quickstart'


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
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials


def process_sheet():
    # use creds to create a client to interact with the Google Drive API
    scope = ['https://spreadsheets.google.com/feeds',  # Getting the scope right is incredibly important
         'https://www.googleapis.com/auth/drive']      # Not totally sure what this means but it works
    creds = ServiceAccountCredentials.from_json_keyfile_name('gspread_secret.json', scope)
    client = gspread.authorize(creds)

    # Open the space requests sheet
    sheet = client.open("Hanszen Space Request Form (2018-2019) (Responses)").sheet1
    # Spreadsheet starts at row 3
    i = 3
    end = False
    while not end:

        GMT_OFF = '-06:00'

        # Make sure that the entry exists
        timestamp = re.search("\'.*\'",str(sheet.cell(i, 1))).group(0)

        if timestamp != "\'\'":
            # Get event information
            Space = re.search("\'.*\'", str(sheet.cell(i, 5))).group(0).replace("\'", "")
            Event = re.search("\'.*\'", str(sheet.cell(i, 6))).group(0).replace("\'", "")
            Date = re.search("\'.*\'", str(sheet.cell(i, 7))).group(0).replace("\'", "")
            Start = re.search("\'.*\'", str(sheet.cell(i, 8))).group(0).replace("\'", "").replace(" ", "")
            End = re.search("\'.*\'", str(sheet.cell(i, 9))).group(0).replace("\'", "").replace(" ", "")
            Recurring = re.search("\'.*\'", str(sheet.cell(i, 10))).group(0).replace("\'", "").replace(" ", "")

            month, day, year = Date.split("/")

            # Add leading zeroes to month and day
            if len(month) == 1:
                month = "0" + month
            if len(day) == 1:
                day = "0" + day

            Date = "{0}-{1}-{2}".format(year, month, day)
            Start = str(dparser.parse(Start)).split(" ")[1]
            End = str(dparser.parse(End)).split(" ")[1]

            summary = "{0}, {1}".format(Event, Space)
            dt_start = "{0}T{1}{2}".format(Date, Start, GMT_OFF)
            dt_end = "{0}T{1}{2}".format(Date, End, GMT_OFF)

            event_json = {
                'summary': summary,  # Title of the event.
                'start': {
                    'dateTime': dt_start,
                },
                # "location": "A String", # Geographic location of the event as free-form text. Optional.
                'end': {
                    'dateTime': dt_end,
                    # "dateTime": "2018-03-24T14:00:00%s" % GMT_OFF
                },
            }
            # Write the event to the calendar
            write_event(event_json)
            i += 1

        else:
            end = True


def write_event(event_json):
    """
    Write an event to the desired Google Calendar
    Input:
    event_json --- information for the event in json format
   """
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('calendar', 'v3', http=http)

    # Make sure event does not already exist
    is_new = True

    # THIS WAS HARDCODED IF SHEET IS CHANGED TIME CAN BE UPDATED
    # Year month day hour minute + 'Z' is needed not sure why
    minTime = datetime.datetime(2018, 2, 28, 12, 30).isoformat() + 'Z'

    eventsResult = service.events().list(
        calendarId="rice.edu_u951gedp8edgqg1ok7u3k8jid8@group.calendar.google.com", timeMin=minTime,
        maxResults=1000, singleEvents=True, orderBy='startTime').execute()
    events = eventsResult.get('items', [])
    for event in events:
        start_i = event['start'].get('dateTime', event['start'].get('date'))
        end_i = event['end'].get('dateTime', event['end'].get('date'))
        summ_i = event['summary']
        summ_f = event_json['summary']
        start_f = event_json['start'].get('dateTime', event_json['start'].get('date'))
        end_f = event_json['end'].get('dateTime', event_json['end'].get('date'))
        start_dif = dparser.parse(start_i) - dparser.parse(start_f)
        end_dif = dparser.parse(end_i) - dparser.parse(end_f)

        if summ_i == summ_f and start_dif.seconds == 0 and end_dif.seconds == 0:
            is_new = False

    # Add event if it's new
    if is_new:
        # Get the calendarId from the setting in Google Calendar
        service.events().insert(calendarId="rice.edu_u951gedp8edgqg1ok7u3k8jid8@group.calendar.google.com",
                                body=event_json).execute()


if __name__ == '__main__':

    process_sheet()