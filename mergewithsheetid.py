import os.path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from datetime import date, datetime

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
      clear_sheets(service, spreadsheet_id, f'{range_name}!A5:U')

    values = _values
    data = [
        {"range": f'{range_name}!A5:U', "values": values},
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

    # 취합 작업 시간 Logging
    mergetime = datetime.now().strftime('%Y-%m-%d %H:%M')
    logData = [["취합시간", mergetime]]
    data = [
        {"range": f'{range_name}!E1:F1', "values": logData},
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

def different_head_sheets(creds):
  try:
    # search_fname = '파일작업예시'  # 찾을 sheet 이름
    # fileinfo = search_file(creds, search_fname)  # search from google drive 
    # print('fileinfo .. : %s' % fileinfo)

    # The ID and range of a sample spreadsheet.
    SCAN_SPREADSHEET_ID = "your spreadsheet id for scanning"  
    # SCAN_SHEET_RANGE_NAME = ["TC_회원!A6:AE"]
    SCAN_SHEET_RANGE_NAME = ["★TC_업무공통!A6:AA","TC_업무공통(SKT)!A6:AA","TC_업무공통(통합알림)!A6:AA","TC_상품!A6:Z","TC_할인!A6:AA",
                             "TC_쿠폰!A6:AA","TC_구독!A5:AD","TC_결제!A5:AA","TC_요금계산!A5:AA","TC_배송관리!A6:AA","TC_배송정책!A6:AA",
                             "TC_채널!A5:AA","TC_제휴입점!A6:AA","TC_제휴사연동!A5:AA","TC_정산!A5:AA","TC_전시!A6:AA","TC_추천마케팅!A6:AA",
                             "TC_회원!A6:AE","TC_사용자권한!A6:AA","★TC_고객상담!A6:AA"]
    service = build("sheets", "v4", credentials=creds) # scan google sheet

    # Call the Sheets API
    sheet = service.spreadsheets()

    # Loop Sheet Range
    allcells = []
              # 0  1  2  3  4  5  6  7  8  9 10 11 12 13 14 15 16 17 18 19 20
    std_row = ['','','','','','','','','','','','','','','','','','','','','']
    # 0-10: 모듈,순번,시나리오ID,시나리오명,케이스ID,케이스명,FO(APP),크롬,엣지,스윙,스윙모바일(Android),
    # 11-20: 스윙모바일(iOS),통테1차,Tester,Test계획일,Test date,등록자,등록일,수정자,수정일,비고  
    for scan_sheet_rg_name in SCAN_SHEET_RANGE_NAME :
      result = (
          sheet.values()
          .batchGet(spreadsheetId=SCAN_SPREADSHEET_ID, ranges=scan_sheet_rg_name)
          .execute()
      )
      ranges = result.get("valueRanges", [])
      scan_sheet_name = scan_sheet_rg_name.split('!')[0]
      # print(f'{type(ranges[0])} - ranges : {ranges[0]}')
      sheet_values = ranges[0]
      values = sheet_values['values']
      print(f"{scan_sheet_rg_name} -{scan_sheet_name} : {len(values)}건 retrieved")
      # print(f'{len(values)}-{values}')
      valid_data_cnt = 0
      for idx, row in enumerate(values):
        # if scan_sheet_name in ["TC_제휴입점"] :
        #   print(f'<{idx}> {row}')

        # 의미있는 값을 포함하는 항목이 3개 미만이면 skip 
        if len(row) < 3 : 
          if scan_sheet_name in ["★TC_업무공통","TC_업무공통(SKT)","TC_업무공통(통합알림)","TC_상품","TC_할인","TC_쿠폰","TC_배송관리","TC_배송정책",
                                 "TC_제휴입점","TC_전시","TC_추천마케팅","TC_회원","TC_사용자권한","★TC_고객상담"]:
            print(f'< row 번호 {idx + 6} : nodata skip >') # A6 부터 Data 시작하는 경우
          elif scan_sheet_name in ["TC_구독","TC_결제","TC_요금계산","TC_채널","TC_제휴사연동","TC_정산"]:
            print(f'< row 번호 {idx + 5} : nodata skip >') # A5 부터 Data 시작하는 경우
          else :
            print(f'< undefined row 번호 {idx} : nodata skip >') # A0 부터 Data 시작하는 경우
          continue
        tmp_row = std_row[:]  # 변수 초기화
        for j in range(0,len(row)) :
          if j <= 4 :
            if j == 0 :
              tmp_row[0] = scan_sheet_name  # 모듈 sheet 명
            tmp_row[j+1] = row[j]

          if scan_sheet_name in ["★TC_업무공통","TC_업무공통(SKT)","TC_업무공통(통합알림)","TC_할인","TC_쿠폰","TC_결제","TC_요금계산",
                                 "TC_배송관리","TC_배송정책","TC_채널","TC_제휴입점","TC_제휴사연동","TC_정산","TC_전시","TC_추천마케팅",
                                 "TC_사용자권한","★TC_고객상담"]:
            if j >= 12: # M
              tmp_row[j-6] = row[j]
          elif scan_sheet_name in ["TC_상품"]:
            if j >= 11: # L
              tmp_row[j-5] = row[j]
          elif scan_sheet_name in ["TC_구독"]:
            if j >= 15: # P
              tmp_row[j-9] = row[j]
          elif scan_sheet_name in ["TC_회원"]:
            if j >= 16: # Q
              tmp_row[j-10] = row[j]

        # if scan_sheet_name in ["TC_제휴입점"] :
        #   print(f"{tmp_row}")
        allcells.append(tmp_row) 
        valid_data_cnt = valid_data_cnt + 1
      print(f'              ==>  merged valid data : {valid_data_cnt} 건')
    print("allcells rowcount : ", len(allcells))

    # write to google sheet 
    batch_update_values(service, SCAN_SPREADSHEET_ID, "@TC_통계(전체)", "USER_ENTERED", allcells) 

  except HttpError as err:
    print(err)

def same_head_sheets(creds) :
  try:
    # search_fname = '파일작업예시'  # 찾을 sheet 이름
    # fileinfo = search_file(creds, search_fname)  # search from google drive 
    # print('fileinfo .. : %s' % fileinfo)
    print("same_head_sheets")
    return None
    # The ID and range of a sample spreadsheet.
        
    SCAN_SPREADSHEET_ID = "your spreadsheet id for scanning"  
    SCAN_SHEET_RANGE_NAME = ["01.상품!A2:Z","02.할인!A2:Z","03.쿠폰!A2:Z","04.구독!A2:Z","05.결제!A2:Z",
                             "06.요금계산!A2:Z","07.배송제고!A2:Z","08.채널!A2:Z","09.제휴입점!A2:Z",
                             "10.제휴사연동!A2:Z","11.정산!A2:Z","12.전시!A2:Z","13.추천마케팅!A2:Z",
                             "14.회원!A2:Z","15.고객상담!A2:Z","16.업무공통!A2:Z",
                             "16.업무공통!A2:Z"]

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
    # print('ranges : ', ranges)
    allcells = []
    for idx in range(len(ranges)) :
      values = ranges[idx].get("values", [])
      for row in values:
        # Print columns A and E, which correspond to indices 0 and 4.
        # print(f"{row[0]}, {row[1]}, {row[2]}, {row[3]}")
        print('row : ', row)
        allcells.append(row)
    # print("allcells : ", allcells)

    # write to google sheet 
    batch_update_values(service, SCAN_SPREADSHEET_ID, "17.관리용(전체취합)", "USER_ENTERED", allcells) 

  except HttpError as err:
    print(err)


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
  
  sheet_type = {"DIFFERENT_HEAD_SHEET": True, "SAME_HEAD_SHEET" : False}

  if sheet_type["DIFFERENT_HEAD_SHEET"] :
    different_head_sheets(creds)
  else :
    same_head_sheets(creds)

if __name__ == "__main__":
  main()
