import argparse
import httplib2
import json
import os
import random

from apiclient import discovery
import oauth2client
from oauth2client import client
from oauth2client import tools

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'GoogleSheetMutator'


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
                                   'sheets-python-quickstart.json')

    store = oauth2client.file.Storage(credential_path)
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


def setup_service():
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
    service = discovery.build('sheets', 'v4', http=http,
                              discoveryServiceUrl=discoveryUrl)

    return service


def read_from_spreadsheet(service, spreadsheet_id, range_name):
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id, range=range_name).execute()
    values = result.get('values', [])
    print('Read %d rows' % len(values))
    return values


def append_to_spreadsheet(service, spreadsheet_id, range_name, row):
    service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            valueInputOption='RAW',
            range=range_name,
            body=row).execute()
    print('Row written to spreadsheet at range ' + range_name + '\n' +
            ', '.join(row['values'][0]))


def get_new_row(values, range_name):
    '''Get a new row.
    The row is selected random between the already existing row.
    '''
    range_name = range_name.split("!")[0] + "!A" + str(len(values) + 2) + ":F"
    row_values = values[random.randint(0, len(values) - 1)][:]
    row = {
        'values': [row_values,],
        'range': range_name,
        'majorDimension': 'ROWS',
    }

    return row, range_name


def parse_arguments():
    parser = argparse.ArgumentParser()
    parser.add_argument("--spreadsheet", metavar="spreadsheet",
            required=True, help="The ID of the spreadsheet to mutate")
    parser.add_argument("--range", metavar="range",
            help="The range of columns to read", required=True)
    parsed_args = parser.parse_args()

    return parsed_args.spreadsheet, parsed_args.range


def main():
    """Shows basic usage of the Sheets API.

    Creates a Sheets API service object and prints the names and majors of
    students in a sample spreadsheet:
    https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
    """
    spreadsheet_id, range_name = parse_arguments()

    service = setup_service()

    values = read_from_spreadsheet(
            service, spreadsheet_id, range_name=range_name)
    if not values:
        print('No data found.')

    new_row, range_name = get_new_row(values, range_name)
    append_to_spreadsheet(
        service, spreadsheet_id, range_name=range_name, row=new_row)


if __name__ == '__main__':
    main()

