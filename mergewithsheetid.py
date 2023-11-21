import os.path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# google cloud에서 api key https://console.cloud.google.com/cloud-resource-manager  :  
# 1. APIs & Services - Enable API : Google Sheet API, Google Drive API
# 2. Credentials OAuth2.0 Client ID 생성 : download file name - client_secrets.json

def search_file(creds, search_filenm):
  """Search file in drive location

  Load pre-authorized user credentials from the environment.
  TODO(developer) - See https://developers.google.com/identity
  for guides on implementing OAuth2 for the application.
  """
#   creds, _ = google.auth.default()

  try:
    # create drive api client
    service = build("drive", "v3", credentials=creds)
    files = []
    page_token = None
    while True:
      # pylint: disable=maybe-no-member
      response = (
          service.files()
          .list(
              q="mimeType='application/vnd.google-apps.spreadsheet'", # https://developers.google.com/drive/api/guides/mime-types?hl=ko
              spaces="drive",
              fields="nextPageToken, files(id, name)",
              pageToken=page_token,
          )
          .execute()
      )
      for file in response.get("files", []):
        # Process change
        print(f'Found file: {file.get("name")}, {file.get("id")}')
        if search_filenm == file.get("name") :
          return file
      page_token = response.get("nextPageToken", None)
      if page_token is None:
        break

  except HttpError as error:
    print(f"An error occurred: {error}")
    files = None

  return files

def add_sheets(service, gsheet_id, sheet_name):
  """Add merge sheet 
     https://learndataanalysis.org/add-new-worksheets-to-existing-google-sheets-file-with-google-sheets-api/
  """
  try:
      request_body = {
          'requests': [{
              'addSheet': {
                  'properties': {
                      'title': sheet_name,
                      'tabColor': {
                          'red': 0.44,
                          'green': 0.99,
                          'blue': 0.50
                      }
                  }
              }
          }]
      }
      response = service.spreadsheets().batchUpdate(
          spreadsheetId=gsheet_id,
          body=request_body
      ).execute()
      return response
  except Exception as e:
      print(e)
      return None

def clear_sheets(service, gsheet_id, sheet_name):
  """Clear merge sheet 
     https://developers.google.com/sheets/api/samples/sheet?hl=ko
  """
  try:
      request_body = {
         'ranges': [sheet_name]
      }
      response = ( 
          service.spreadsheets()
          .values()
          .batchClear(spreadsheetId=gsheet_id, body=request_body)
          .execute()
      )
      print('clear sheet : ', response)
      return response
  except Exception as e:
      print(e)
  
def batch_update_values(
    service, spreadsheet_id, range_name, value_input_option, _values
):
  """
  Creates the batch_update the user has access to.
  Load pre-authorized user credentials from the environment.
  TODO(developer) - See https://developers.google.com/identity
  for guides on implementing OAuth2 for the application.
  """
#   creds, _ = google.auth.default()
  # pylint: disable=maybe-no-member
  try:
    # service = build("sheets", "v4", credentials=creds)

    if add_sheets(service, spreadsheet_id, range_name) == None : # Merge sheet 생성하기
      clear_sheets(service, spreadsheet_id, range_name)

    values = _values
    data = [
        {"range": range_name, "values": values},
        # Additional ranges to update ...
    ]
    body = {"valueInputOption": value_input_option, "data": data
           }
    result = (
        service.spreadsheets()
        .values()
        .batchUpdate(spreadsheetId=spreadsheet_id, body=body)
        .execute()
    )
    print(f"{(result.get('totalUpdatedCells'))} cells updated.")
    return result
  except HttpError as error:
    print(f"An error occurred: {error}")
    return error
  
def main():
  """Shows basic usage of the Sheets API.
  Prints values from a sample spreadsheet.
  """

  creds = None
  # If modifying these scopes, delete the file token.json. 
  # https://developers.google.com/identity/protocols/oauth2/scopes?hl=ko
  SCOPES = ["https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"]
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
          "client_secrets.json", SCOPES
      )
      creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open("token.json", "w") as token:
      token.write(creds.to_json())

  try:
    # search_fname = '파일작업예시'  # 찾을 sheet 이름
    # fileinfo = search_file(creds, search_fname)  # search from google drive 
    # print('fileinfo .. : %s' % fileinfo)

    # The ID and range of a sample spreadsheet.
    SCAN_SPREADSHEET_ID = "1O5atw1GsYukg-fFsZFiR47fQC8y4FYF_dNKnFozR8BY" 
    SCAN_SHEET_RANGE_NAME = ["통합테스트_Check List!A:Z"]

    service = build("sheets", "v4", credentials=creds) # scan google sheet

    # Call the Sheets API
    sheet = service.spreadsheets()
    result = (
        sheet.values()
        .batchGet(spreadsheetId=SCAN_SPREADSHEET_ID, ranges=SCAN_SHEET_RANGE_NAME)
        .execute()
    )
    ranges = result.get("valueRanges", [])
    print(f"{len(ranges)} ranges retrieved")
    print('ranges : ', ranges)
    allcells = []
    for idx in range(len(ranges)) :
      values = ranges[idx].get("values", [])
      for row in values:
        # Print columns A and E, which correspond to indices 0 and 4.
        # print(f"{row[0]}, {row[1]}, {row[2]}, {row[3]}")
        print('row : ', row)
        allcells.append(row)
    print("allcells : ", allcells)

    # write to google sheet 
    batch_update_values(service, SCAN_SPREADSHEET_ID, "통합테스트_MergeList", "USER_ENTERED", allcells) 

  except HttpError as err:
    print(err)


if __name__ == "__main__":
  main()