from __future__ import print_function
import datetime
import dateutil.relativedelta as REL
import pickle
import os.path
import time
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import traceback

# If modifying these scopes, delete the file token.pickle.
SCOPES = [
            'https://www.googleapis.com/auth/calendar',
            'https://www.googleapis.com/auth/tasks'
         ]

# Code from Google API Documentation
def googleAuth():
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    return creds

def main():
    print("\n\nCalendar Copy by Rohan Prabhu")
    print("NOTE: Have steps 1 & 2 completed from https://developers.google.com/calendar/quickstart/python")
    print("This app will be able to access/edit your Google Calendar and Tasks.")
    check = str(input("Continue? (y/n) "))
    if(check == 'n'):
        exit(0)

    operation = input("\n\nOPTIONS: \n 1. Create Tasks from All-Day events on a different calendar \n 2. Duplicate Hourly Events from a different Calendar \n 3. 1+2 \nENTER: ")
    if(operation not in ["1", "2", "3"]):
        print("Invalid option provided")
        exit(1)

    pot_start_date = datetime.datetime.utcnow() + REL.relativedelta(weekday=REL.SU)
    pot_end_date = (pot_start_date + REL.relativedelta(weekday=REL.SA))
    print(pot_start_date, pot_end_date)

    date_option = input("\n\nDATE OPTIONS: \n 1. Next Week (Sun - Sat) [%s] \n 2. Entered Dates \nENTER: " % (str(pot_start_date.strftime('%m-%d-%Y'))
                                                                                                                        + " - " + str(pot_end_date.strftime('%m-%d-%Y'))))
    if(date_option == "1"):
        start_date = pot_start_date.isoformat()+"Z"
        end_date = pot_end_date.isoformat()+"Z"
        print(start_date, end_date)
    elif(date_option == "2"):
        try:
            start_date_input = str(input("Start Date [format: mm/dd/yy]: "))
            end_date_input = str(input("End Date [format: mm/dd/yy]: "))
            start_date = datetime.datetime.strptime(start_date_input, '%m/%d/%y').isoformat()+"Z"
            end_date = datetime.datetime.strptime(end_date_input, '%m/%d/%y').isoformat()+"Z"
        except Exception as e:
            print("Invalid dates provided")
            exit(1)
    else:
        print("Invalid option provided")
        exit(1)

    try:
        creds = googleAuth()
    except Exception as err:
        traceback.print_exc()
        exit(1)

    calendar_service = build('calendar', 'v3', credentials=creds)
    tasks_service = build('tasks', 'v1', credentials=creds)
    existingTasks = tasks_service.tasks().list(tasklist='@default', dueMax = end_date).execute()

    print("\n\nEnter the calendar Title / search term to copy from (from your personal Google Calendar List)\nNOTE - if the search term matches > 1 calendars, it will copy from all.")
    searchTerm = input("ENTER: ")
    # searchTerm = "MSA"
    page_token = None
    calendar_ids = []

    # Lots of help from https://qxf2.com/blog/google-calendar-python/
    # Currently set up to search for case of multiple matching calendars - WILL work for one
    while True:
        calendar_list = calendar_service.calendarList().list(pageToken=page_token).execute()
        for calendar_list_entry in calendar_list['items']:
            if searchTerm in calendar_list_entry['summary']:
                calendar_ids.append(calendar_list_entry['id'])
            elif 'summaryOverride' in calendar_list_entry and searchTerm in calendar_list_entry['summaryOverride']:
                calendar_ids.append(calendar_list_entry['id'])
        page_token = calendar_list.get('nextPageToken')
        if not page_token:
            break

    for calendar_id in calendar_ids:
        eventsResult = calendar_service.events().list(
                calendarId=calendar_id,
                timeMin=start_date,
                timeMax=end_date,
                singleEvents=True,
                orderBy='startTime').execute()

        events = eventsResult.get('items', [])
        for event in events:
            if("optional" not in event['summary'] and "Optional" not in event['summary']): # MSA specific concern
                if(operation == "1" or operation == "3"):
                    if('start' in event and 'date' in event['start']):
                        # Characteristic of an all-day event
                        task = {
                          'title': event['summary'],
                          'due': event['start']['date']+'T12:00:00.000Z'
                        }
                        try:
                            result = tasks_service.tasks().insert(tasklist='@default', body=task).execute()
                            print('Task created from all-day event: %s' % (event['summary']))
                        except Exception as e:
                            print("Not created - all-day event -> task name: " + event['summary'])
                            traceback.print_exc()
                            continue
                if(operation == "2" or operation == "3"):
                    if('start' in event and 'date' not in event['start']):
                        # Normal event
                        try:
                            event = calendar_service.events().insert(calendarId='primary', body=event).execute()
                            print('Event created: %s' % (event.get('summary')))
                        except Exception as e:
                            print("Not created - hourly event name: " + event['summary'] )
                            traceback.print_exc()
                            continue

    print("Complete!")
    exit(0)


if __name__ == '__main__':
    main()
