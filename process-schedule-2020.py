#!/usr/bin/env python3

# [START sheets_quickstart]
from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

import csv
import re
from datetime import date, datetime

# Main parameters
cutoff = date.fromisoformat('2020-02-04')
roles = ['TM', 'SP', 'TT', 'GE', 'E', 'GR', 'IT', 'CJ', '1M']

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# The ID and range of a sample spreadsheet.
# SAMPLE_SPREADSHEET_ID = '1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms'
# RANGE_FORMAT = 'Class Data!A2:E'
SPREADSHEET_ID = '1xD1dnISQGhi2xDriT4kG-ABvcZMO-8vCaRclGn3LjsY'
RANGE_FORMAT = 'Schedule 2020!A1:AZ50'
OUTFILE = 'roledates2020.csv'

def main():
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
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
    
    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                range=RANGE_FORMAT.format(1)).execute()
    values = result.get('values', [])

    dates = list(filter(lambda d: d <= cutoff, [datetime.strptime(d, '%d-%b-%Y').date() for d in values[0][1:]]))
    with open(OUTFILE, 'w') as wf:
        writer = csv.DictWriter(wf, ["Name"]+roles)
        writer.writeheader()

        for row in values[1:]:
            if row[0] == "Open Roles":
                break;
            
            # Extract (role, date) iterable -- ignoring empty or N/A rode
            rd = filter(lambda p: p[0] and not re.match(r'(?i)N/A', p[0]),
                [(row[i+1], dates[i]) for i in range(0, len(dates)-1)])

            # Expand multiple roles
            rd = [(list(filter(lambda r: re.match(r'\w+', r), re.split(r'\W+', p[0]))), p[1]) for p in rd]
            
            # filter out unwanted roles
            dd = dict(filter(lambda rp: rp[0] in roles, [(r, p[1]) for p in rd for r in p[0]]))

            roledates = { "Name": row[0] }
            roledates.update()
            for role in roles:
                roledates[role] = dd[role] if role in dd else None

            writer.writerow(roledates)
            
if __name__ == '__main__':
    main()
# [END sheets_quickstart]