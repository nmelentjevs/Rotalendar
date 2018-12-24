from __future__ import print_function
from googleapiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools
from xlsrotaparse import *

# If modifying these scopes, delete the file token.json.
try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

SCOPES = 'https://www.googleapis.com/auth/calendar'
store = file.Storage('credentials.json')
creds = store.get()
if not creds or creds.invalid:
    flow = client.flow_from_clientsecrets('credentials.json', SCOPES)
    creds = tools.run_flow(flow, store, flags) \
    if flags else tools.run(flow, store)

CAL = build('calendar', 'v3', http=creds.authorize(Http()))


GMT_OFF = '-0:00'
def createrota():
    for i, k in zip(calendardatesstart, calendardatesfinish):

        event = {
          'summary': 'Work',
          'location': 'Hippodrome Casino Work',
          'description': 'Work work work work work',
          'start': {
            'dateTime': i,
            'timeZone': 'Greenwich Mean Time',
          },
          'end': {
            'dateTime': k,
            'timeZone': 'Greenwich Mean Time',
          },
          'reminders': {
            'useDefault': False,
            'overrides': [
              {'method': 'email', 'minutes': 24 * 60},
              {'method': 'popup', 'minutes': 10},
            ],
          },
        }


        e = CAL.events().insert(calendarId='primary', body=event).execute()

        print('Event created: Workday %s' % i)

createrota()