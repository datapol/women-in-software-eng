#!/usr/bin/python

"""Much thanks to this blog post:
http://www.mattcutts.com/blog/write-google-spreadsheet-from-python/
"""

import argparse

from gdata import gauth
from gdata.spreadsheet import service as ss_service
from gdata.service import RequestError
from oauth2client.client import flow_from_clientsecrets
from oauth2client.file import Storage
from oauth2client import tools
from read_data import read_data

CLIENT_SECRETS_FILENAME = 'client_secrets.json'
SS_KEY = '0AlZH8QBl60oodEJTdFA5TlZOcDJCMU02RkZoSHF5SHc'
WORKSHEET_ID = 'od6'

ss_client = None

def init_ss_client(client_secrets_filename, flags):
    global ss_client
    if not ss_client:
        storage = Storage('creds.dat')
        credentials = storage.get()
        if credentials is None or credentials.invalid:
            credentials = tools.run_flow(
                flow_from_clientsecrets(
                    client_secrets_filename,
                    scope=['https://spreadsheets.google.com/feeds']),
                storage,
                flags)

        ss_client = ss_service.SpreadsheetsService(
            additional_headers={'Authorization': 'Bearer %s' % credentials.access_token})
        ss_client.auth_token = gauth.OAuth2TokenFromCredentials(credentials)
        ss_client.source = 'Update Script'
        ss_client.ProgrammaticLogin()

def clear_ss_data(ss_key, worksheet_id):
    list_feed = ss_client.GetListFeed(ss_key, worksheet_id)
    if list_feed:
        for row in list_feed.entry:
            ss_client.DeleteRow(row)

def update_ss_from_file(ss_key, worksheet_id, data_filename):
    if not ss_client:
        print('Oops! SpreadsheetsService client not initialized.')
        return

    clear_ss_data(ss_key, worksheet_id)

    rows_data = read_data(data_filename)

    for row_data in rows_data:
        row_data = dict((key.replace('_', ''), str(value))
                        for key, value in row_data.items())
        inserted = False
        while not inserted:
            try:
                ss_client.InsertRow(row_data, ss_key, worksheet_id)
                inserted = True
            except RequestError as e:
                print("Request error: {0}".format(e))
            except:
                raise


if __name__ == '__main__':
    parser = argparse.ArgumentParser('Update spreadsheet.', parents=[tools.argparser])
    parser.add_argument('-c', '--client-secrets-filename')
    parser.add_argument('-d', '--data-filename')
    parser.add_argument('-s', '--spreadsheet-key')

    flags = parser.parse_args()
    client_secrets_filename = flags.client_secrets_filename or CLIENT_SECRETS_FILENAME
    data_filename = flags.data_filename
    ss_key = flags.spreadsheet_key

    init_ss_client(client_secrets_filename, flags)
    update_ss_from_file(ss_key, WORKSHEET_ID, data_filename)
