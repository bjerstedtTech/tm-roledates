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
cutoff = date.fromisoformat('2020-03-04')
roles = ['TM', 'SP', 'TT', 'GE', 'E', 'GR', 'IT', 'CJ', '1M']

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# The ID and range of a sample spreadsheet.
# SAMPLE_SPREADSHEET_ID = '1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms'
# RANGE_FORMAT = 'Class Data!A2:E'
SPREADSHEET_ID = '1xD1dnISQGhi2xDriT4kG-ABvcZMO-8vCaRclGn3LjsY'
# RANGE_FORMAT = 'Schedule 2020!A1:AZ50'
RANGE_FORMAT = 'Schedule {}!A1:AZ50'
OUTFILE = 'roledates2020.csv'

def getSheetService():
    """Opens processor for Google Sheets API"""
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
    
    return build('sheets', 'v4', credentials=creds)


def readYear(sheetService, year, cutoff = None):
    # Read the sheet and extract the values
    sheet = sheetService.values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=RANGE_FORMAT.format(year)
        ).execute()
    values = sheet.get('values', [])

    # Get the list of dates from the first line
    dates = list(filter(lambda d: d <= cutoff, [datetime.strptime(d, '%d-%b-%Y').date() for d in values[0][1:]]))
    indices = range(1, len(dates)+1)

    result = []
    # Process each line on sheet until we find one containing "Open Roles" in first column
    for row in values[1:]:
        if re.match(r'(?i)^open roles', row[0]): break

        # get array of role-list date pairs for current participant
        rd = filter(lambda p: p[0] and not re.match(r'(?i)N/A', p[0]),
            [(row[i] if i < len(row) else None, dates[i-1]) for i in indices])

        # break up multiple roles into arrays
        rd = [(list(filter(lambda r: re.match(r'\w+', r), re.split(r'\W+', p[0]))), p[1]) for p in rd]

        # filter out unwanted roles and build dictionary of roles + name
        #   Assumption is original row has dates in ascending order so only most recent
        #   date for each role will be kept.
        dd = dict(filter(lambda rp: rp[0] in roles, [(r, p[1]) for p in rd for r in p[0]]))
        dd['Name'] = re.split(r'[-,]', row[0])[0]

        result.append(dd)

    return result   

def main():
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """

    # Call the Sheets API
    sheetService = getSheetService().spreadsheets()

    pastYear = readYear(sheetService, 2019, cutoff)
    currentYear = readYear(sheetService, 2020, cutoff)

    with open(OUTFILE, 'w') as wf:
        writer = csv.DictWriter(wf, ["Name"]+roles)
        writer.writeheader()

        for currRoles in currentYear:
            pastRoles = [r for r in pastYear if r['Name'] == currRoles['Name']]
            pastRoles = pastRoles[0] if len(pastRoles) == 1 else None
            if pastRoles:
                for role in roles:
                    if not role in currRoles and role in pastRoles:
                        currRoles[role] = pastRoles[role]

            writer.writerow(currRoles)
            
if __name__ == '__main__':
    main()
# [END sheets_quickstart]