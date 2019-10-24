""""
Google Sheets Script for new sheets documents
Description: This is a script to automatically create a Google Sheet and populate
it with the right data.

Note: most of code is from following Google's poorly documented documentation
* https://developers.google.com/drive
* https://developers.google.com/sheets/api
"""

# imports
from __future__ import print_function
import datetime
import pickle
import os.path
import os, winshell
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from apiclient import errors
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/drive.metadata.readonly','https://www.googleapis.com/auth/drive','https://www.googleapis.com/auth/spreadsheets.readonly','https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']

def main():
    """ Warning: spaghetti code
    """

    # inits dates
    # put: + datetime.timedelta(days='Days') # replace 'Days' # behind date_now to change the when to create sheets
    date_now = datetime.datetime.now() + datetime.timedelta(days=1)
    date_end = date_now + datetime.timedelta(days=6)

    # builds service interaction
    service = pickler()

    # folder name test
    if date_now.month == date_end.month:
        month = str(date_now.month)
    else:
        month = str(date_now.month) + '-' + str(date_end.month)
    dates = str(date_now.day) + '-' + str(date_end.day)
    file_name = month + '/' + dates
    folder_name = date_now.strftime("%B") + " " + date_now.strftime("%y")

    # # Call the Drive v3 API
    file_id = copy_file(service, '''TODO: Insert id of google sheets template''' ,file_name).get('id')
    folder_names = print_top_files(service)

    # test in folder exists
    flag = False
    for folder, id in folder_names:
        if folder_name == folder:
            flag = True
            folder_id = id
            break

    if flag != True:
        folder_id  = make_folder(service, folder_name)

    move_into_folder(service, folder_id, file_id)

    # call sheets API
    sheets(file_id, SCOPES, date_now)
    windows_change(file_id)

def windows_change(file_id):
    ''' changes windows shell startup '''

    url = 'https://docs.google.com/spreadsheets/d/' + str(file_id)
    path = os.path.join("C:/Users/zeyul/AppData/Roaming/Microsoft/Windows/Start Menu/Programs/Startup", "important.url")
    if os.path.exists(path):
        os.remove(path)
    with open(path, 'w') as fp:
        fp.write(f'[InternetShortcut]\n')
        fp.write('URL=%s' % url)


def sheets(file_id, SCOPES, date_now):
    ''' changes the sheet time and dates for every sheets '''

    credentials = ServiceAccountCredentials.from_json_keyfile_name('sheets.json', SCOPES)
    gc = gspread.authorize(credentials)
    sh = gc.open_by_key(file_id)
    worksheet_list = sh.worksheets()
    for index, worksheet in enumerate(worksheet_list):
        day_ = str((date_now + datetime.timedelta(days=index)).day)
        month_ = str((date_now + datetime.timedelta(days=index)).month)
        worksheet = sh.get_worksheet(index)
        worksheet.update_title(day_)
        worksheet.update_acell('A1', f'{month_}/{day_}')

def move_into_folder(service, folder_id, file_id):
    ''' moves doc to folder '''
    file = service.files().get(fileId=file_id,
                               fields='parents').execute()
    previous_parents = ",".join([parent["id"] for parent in file.get('parents')])
    # Move the file to the new folder
    file = service.files().update(fileId=file_id,
                                        addParents=folder_id,
                                        removeParents=previous_parents,
                                        fields='id, parents').execute()

def make_folder(service, folder_name):
    ''' makes folder if not there '''
    file_metadata = {
        'title': folder_name,
        'mimeType': 'application/vnd.google-apps.folder',
        'parents': [{'id': '''TODO: id of folder'''}]
    }
    file = service.files().insert(body=file_metadata,
                                  fields='id').execute()
    return file.get('id')

def print_top_files(service):
    result = []
    folder_names = []
    page_token = None
    while True:
        try:
            param = {}
            if page_token:
                param['pageToken'] = page_token
            files = service.files().list(q="mimeType = 'application/vnd.google-apps.folder' and  '''TODO: id of folder''' in parents",
                                         spaces='drive',
                                         fields='nextPageToken, items(id, title)',**param).execute()

            for file in files.get('items', []):
                # get title into array
                folder_names.append({file.get('title'), file.get('id')})
            result.extend(files['items'])
            page_token = files.get('nextPageToken')
            if not page_token:
                break
        except errors.HttpError as error:
            print('An error occurred: %s' % error)
        break
    return folder_names


def copy_file(service, origin_file_id, copy_title):
    """Copy an existing file.

    Args:
        service: Drive API service instance.
        origin_file_id: ID of the origin file to copy.
        copy_title: Title of the copy.

    Returns:
        The copied file if successful, None otherwise.
    """
    copied_file = {'title': copy_title}
    try:
        return service.files().copy(
            fileId=origin_file_id, body=copied_file).execute()
    except errors.HttpError as error:
        print('An error occurred: %s' % error)
    return None

def pickler():
    ''' google's pickle '''
    creds = None
    """The file token.pickle stores the user's access and refresh tokens, and is
    created automatically when the authorization flow completes for the first
    time.
    """
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

    return build('drive', 'v2', credentials=creds)

if __name__ == '__main__':
    main()
