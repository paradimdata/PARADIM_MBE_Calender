import datetime
import os.path
from pandas import DataFrame

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/calendar.readonly"]


def main():
  """Shows basic usage of the Google Calendar API.
  Prints the start and name of the next 10 events on the user's calendar.
  """
  creds = None
  # The file token.json stores the user's access and refresh tokens, and is
  # created automatically when the authorization flow completes for the first
  # time.
  if os.path.exists("token.json"):
    creds = Credentials.from_authorized_user_file("token.json", SCOPES)
  # If there are no (valid) credentials available, let the user log in.
  if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
      creds.refresh(Request())
    else:
      flow = InstalledAppFlow.from_client_secrets_file(
          "credentials.json", SCOPES
      )
      creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open("token.json", "w") as token:
      token.write(creds.to_json())

  try:
    service = build("calendar", "v3", credentials=creds)

    # Call the Calendar API
    now = datetime.datetime.utcnow().isoformat() + "Z"  # 'Z' indicates UTC time
    print("Getting the upcoming 10 events")
    events_result = (
        service.events()
        .list(
            calendarId="okvnvt57fklmv6ctnsbrd32f1o@group.calendar.google.com",
            timeMin=now,
            singleEvents=True,
            orderBy="startTime",
        )
        .execute()
    )
    events = events_result.get("items", [])
    props = [[],[],[]]

    if not events:
      print("No upcoming events found.")
      return

    # Prints the start and name of the next 10 events
    for event in events:
      start = event["start"].get("dateTime", event["start"].get("date"))
      end = event["end"].get("dateTime", event["end"].get("date"))
      summary = event["summary"]
      props[0] = props[0] + [start]
      props[1] = props[1] + [end]
      props[2] = props[2] + [summary]
      print(start, event["summary"])
      print(end, event["summary"])
      
    index = 0
    prop_numbers = []  
    for index in range(len(props[2]) - 1):
      if '#' in props[2][index]:
        prop_numbers = prop_numbers + [int(i) for i in test_string.split() if i.isdigit()]


    df = DataFrame({'Prop Number': props[2], 'Start Date': props[0], 'End Date': props[1]})
    df.to_excel('Calendar.xlsx', sheet_name='sheet1', index=False)

  except HttpError as error:
    print(f"An error occurred: {error}")


if __name__ == "__main__":
  main()