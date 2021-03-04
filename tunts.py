from __future__ import print_function
import pickle
import os.path
from math import floor, ceil
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# The ID and range of a sample spreadsheet.
SPREADSHEET_ID = '1A8DqosDjCkTO_GgahWVxPbkwNXxl8N2pC1Y7ByTEjY0'
RANGE_NAME = 'engenharia_de_software!A2:F'
POST_RANGE_NAME = 'engenharia_de_software!G4:H'

ABSENCE_TOLERANCE = 0.25
ABSENCE_NUMBER_INDEX = 2
P1_INDEX = 3
P3_INDEX = 6

def setCredentials():
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

    return service

def getValues(service):

    # Call the Sheets API
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                range=RANGE_NAME).execute()
    values = result.get('values', [])

    return values

def postValues(service, result):
    body = {
        'values': result
    }
    result = service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID, range=POST_RANGE_NAME,
        valueInputOption='RAW', body=body).execute()

def isFlunkedByAbsence(total_classes, number_absences):
    return number_absences > total_classes*ABSENCE_TOLERANCE

def calculateMeanTests(tests):
    acc = 0
    for test in tests:
        acc += int(test)
    return acc/len(tests)

def getStudentStatus(total_classes, number_absences, tests):
    if isFlunkedByAbsence(total_classes, number_absences):
        return ['Reprovado por falta', 0]
    else:
        m = calculateMeanTests(tests)
        if m < 50:
            return ['Reprovado por Nota', 0]
        else:
            if m < 70:
                naf = ceil(100-m)
                return ['Exame Final', naf]
            else:
                return ['Aprovado', 0]

def main():
    service = setCredentials()
    values = getValues(service)
    total_classes = values[0][0].split()[-1]
    result = []
    del values[:2]
    for row in values:
        status = getStudentStatus(int(total_classes), int(row[ABSENCE_NUMBER_INDEX]), row[P1_INDEX:P3_INDEX])
        result.append(status)

    postValues(service, result)

if __name__ == '__main__':
    main()