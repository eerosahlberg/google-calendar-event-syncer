# Google Calendar event syncer

This is a simple script that syncs events defined in a Google sheet to a Google calendar. It is intended to be run as a cron job on a server or on a serverless platform like Google Cloud functions.

## Setup

### Python libraries

The script uses the following Python libraries:

-   `google-api-python-client` - Google APIs client library for Python
-   `google-auth` - Google Authentication Library: OAuth 2.0 client
-   `python-dotenv` - Reads the key-value pair from .env file and can set them as environment variables.
-   `boto3` - AWS SDK for Python
-   `pandas` - Python Data Analysis Library

These can be installed with `pip install -r requirements.txt`.

### Google Sheet

The Google sheet should have the following columns:

-   `Mitä:` - The title of the event
-   `Alkaa:` - (dd.mm.yyyy / dd.mm.yyyy klo hh.mm) The start date/datetime of the event
-   `Loppuu:` - (dd.mm.yyyy / dd.mm.yyyy klo hh.mm) The end date/datetime of the event
-   `Kestää koko päivän:` - (Kyllä/Ei) Whether the event should be marked as an all-day event
-   `id:` - The id of the event in the calendar. If this is not set, the script will create a new event. If this is set, the script will update the event with the given id. If the id is set to 1, the script will ignore the event.

Names of the columns can be customized in the script to work with sheets constructed in other ways.

### Google Calendar

The script updates to a calendar with the id defined in the `.env`-file. The calendar should be shared with the service account email address.

### Google service account

The script uses a Google service account to authenticate to the Google APIs. The service account should have the following permissions:

-   `Calendar API` - `Calendar API` - `View and edit events on all your calendars`
-   `Google Sheets API` - `Google Sheets API` - `View and edit spreadsheets that this application has been installed in`

The service account should be shared with the calendar and the sheet.

The service account credentials loaded from the Google cloud console need to be saved as `credentials.json` in the root of the project.

### Environment variables

The script uses the following environment variables:

-   `SHEET_ID` - The id of the Google sheet
-   `CALENDAR_ID` - The id of the Google calendar

### AWS SNS notifications

The script can send notifications to an AWS SNS topic when an event is created or updated. The following environment variables are used for this:

-   `AWS_ACCESS_KEY_ID` - The AWS access key id
-   `AWS_SECRET_ACCESS_KEY` - The AWS secret access key
-   `AWS_ERROR_TOPIC_ARN` - The ARN of the SNS topic to send error notifications to
-   `AWS_INFO_TOPIC_ARN` - The ARN of the SNS topic to send info notifications to

The script is configured to work in AWS region `eu-north-1` by default. This can be changed in the script.
