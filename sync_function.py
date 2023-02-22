import os

import boto3
import dotenv
import pandas as pd
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build


def main():
    dotenv.load_dotenv()

    # Loading google service account credentials
    creds = Credentials.from_service_account_file('./credentials.json')

    calendar = build('calendar', 'v3', credentials=creds)
    sheets = build('sheets', 'v4', credentials=creds)

    sheet = sheets.spreadsheets().get(spreadsheetId=os.getenv('SHEET_ID')).execute()
    sheet = sheet['sheets'][0]['properties']['title']
    sheet = sheets.spreadsheets().values().get(spreadsheetId=os.getenv(
        'SHEET_ID'), range=sheet).execute()  # Getting the sheet data

    sheet_df = pd.DataFrame(sheet['values'][1:], columns=sheet['values'][0])
    # Selecting the columns we need
    sheet_df = sheet_df[['Mitä?', 'Alkaa:',
                         'Loppuu:', 'Kestää koko päivän?', 'id']]

    # Converting the dates to datetime objects
    sheet_df['Alkaa:'] = pd.to_datetime(sheet_df['Alkaa:'], format='%d.%m.%Y klo %H.%M', errors='coerce').fillna(
        pd.to_datetime(sheet_df['Alkaa:'], format='%d.%m.%Y', errors='coerce'))
    sheet_df['Loppuu:'] = pd.to_datetime(sheet_df['Loppuu:'], format='%d.%m.%Y klo %H.%M', errors='coerce').fillna(
        pd.to_datetime(sheet_df['Loppuu:'], format='%d.%m.%Y', errors='coerce'))

    # Getting all events from the calendar we want to sync with
    pageToken = None
    eventDf = pd.DataFrame(['id', 'summary', 'start', 'end'])
    while True:
        events = calendar.events().list(calendarId=os.getenv(
            'CALENDAR_ID'), pageToken=pageToken).execute()
        for event in events['items']:
            eventDf = pd.concat([eventDf, pd.DataFrame(
                [[event['id'], event['summary'], event['start'], event['end']]], columns=['id', 'summary', 'start', 'end'])])
        pageToken = events.get('nextPageToken')
        if not pageToken:
            break

    # Uncomment these lines to export the data to excel files
    # sheet_df.to_excel('sheet.xlsx', index=False)
    # eventDf.to_excel('events.xlsx', index=False)

    # Iterating through the rows in the sheet
    for index, row in sheet_df.iterrows():
        if (row['Kestää koko päivän?'] == 'Kyllä'):
            allDay = True
        else:
            allDay = False

        # Checking if the event is already in the calendar
        if row['id'] == '1':
            continue
        elif (row['id'] in eventDf['id'].values):
            compare = eventDf[eventDf['id'] == row['id']]

            # Checking if the event has changed
            if allDay:
                if compare['summary'].values[0] == row['Mitä?'] and compare['start'].values[0]['date'] == row['Alkaa:'].strftime('%Y-%m-%d') and compare['end'].values[0]['date'] == row['Loppuu:'].strftime('%Y-%m-%d'):
                    continue
            else:
                if compare['summary'].values[0] == row['Mitä?'] and compare['start'].values[0]['dateTime'] == row['Alkaa:'].strftime('%Y-%m-%dT%H:%M:%S') and compare['end'].values[0]['dateTime'] == row['Loppuu:'].strftime('%Y-%m-%dT%H:%M:%S'):
                    continue

            # Updating the event
            event = calendar.events().get(calendarId=os.getenv(
                'CALENDAR_ID'), eventId=row['id']).execute()
            event['summary'] = row['Mitä?']
            if allDay:
                event['start']['date'] = row['Alkaa:'].strftime('%Y-%m-%d')
                event['end']['date'] = row['Loppuu:'].strftime('%Y-%m-%d')
            else:
                event['start']['dateTime'] = row['Alkaa:'].strftime(
                    '%Y-%m-%dT%H:%M:%S')
                event['end']['dateTime'] = row['Loppuu:'].strftime(
                    '%Y-%m-%dT%H:%M:%S')
            event = calendar.events().update(calendarId=os.getenv('CALENDAR_ID'),
                                             eventId=event['id'], body=event).execute()

            print(f'Event updated: {event.get("htmlLink")}')
            sendNotificationEmail(
                event['htmlLink'], 'Tapahtuma päivitetty kalenterissa')

        elif row['id'] == '' or row['id'] == ' ' or row['id'] == None:
            # Creating a new event
            event = {
                'summary': row['Mitä?'],
                'start': {
                    'timeZone': 'Europe/Helsinki',
                },
                'end': {
                    'timeZone': 'Europe/Helsinki',
                },
            }

            if allDay:
                event['start']['date'] = row['Alkaa:'].strftime('%Y-%m-%d')
                # Google Calendar doesn't include the last day of an all-day event
                event['end']['date'] = (
                    row['Loppuu:'] + pd.Timedelta(days=1)).strftime('%Y-%m-%d')
                print(event['start']['date'])
                print(event['end']['date'])
            else:
                event['start']['dateTime'] = row['Alkaa:'].strftime(
                    '%Y-%m-%dT%H:%M:%S')
                event['end']['dateTime'] = row['Loppuu:'].strftime(
                    '%Y-%m-%dT%H:%M:%S')
                print(event['start']['dateTime'])
                print(event['end']['dateTime'])

            event = calendar.events().insert(
                calendarId=os.getenv('CALENDAR_ID'), body=event).execute()
            sheets.spreadsheets().values().update(spreadsheetId=os.getenv('SHEET_ID'),
                                                  range=f'L{index+2}', valueInputOption='USER_ENTERED', body={'values': [[event['id']]]}).execute()
            print('Event created: %s' % (event.get('htmlLink')))
            sendNotificationEmail(
                event['htmlLink'], 'Uusi tapahtuma lisätty kalenteriin')


def sendNotificationEmail(message, subject):
    client = boto3.client('sns', region_name='eu-north-1', aws_access_key_id=os.getenv(
        'AWS_ACCESS_KEY_ID'), aws_secret_access_key=os.getenv('AWS_SECRET_ACCESS_KEY'))
    response = client.publish(
        TopicArn=os.getenv('SNS_INFO_TOPIC_ARN',),
        Message=message,
        Subject=subject
    )
    return response


def sendErrorEmail(message, subject):
    client = boto3.client('sns', region_name='eu-north-1', aws_access_key_id=os.getenv(
        'AWS_ACCESS_KEY_ID'), aws_secret_access_key=os.getenv('AWS_SECRET_ACCESS_KEY'))
    response = client.publish(
        TopicArn=os.getenv('SNS_ERROR_TOPIC_ARN'),
        Message=message,
        Subject=subject
    )
    return response


def function_handler(event, context):
    try:
        main()
    except Exception as e:
        sendErrorEmail(
            str(e), 'Virhe google-tapahtumakalenterin päivittämisessä')
        raise e


if __name__ == '__main__':
    function_handler(None, None)
