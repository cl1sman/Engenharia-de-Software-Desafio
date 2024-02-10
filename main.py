# Dependencies:> pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib
# Run:> python3 main.py

import os.path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# The ID and range of a sample spreadsheet.
SPREADSHEET_ID = "Spreadsheet_id"
RANGE_NAME = "engenharia_de_software!A4:H27"


def main():
  # Logging
  creds = None
  
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
    service = build("sheets", "v4", credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()
    result = (
        sheet.values()
        .get(spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME)
        .execute()
    )

    values = result['values']
    temp_table = []
    for row in values:
        # Average
        avg = ((int(row[3]) + int(row[4]) + int(row[5])) / 3) / 10

        # nf > 25%
        if int(row[2]) > (60 * 0.25):
            situation = 'Reprovado por Falta'
            row.append(situation)
            row.append(0)
            temp_table.append(row)
        
        elif avg < 5:
            situation = 'Reprovado por Nota'
            row.append(situation)
            row.append(0)
            temp_table.append(row)
            
        elif 5 <= avg < 7:
            situation = 'Exame Final'
            naf = round(10 - avg)
            row.append(situation)
            row.append(naf)
            temp_table.append(row)
            
        else:
            situation = 'Aprovado'
            row.append(situation)
            row.append(0)
            temp_table.append(row)
            
    for i in temp_table:
        print(i)
        
    # Update Sheets API
    
    result = (
        sheet.values()
        .update(spreadsheetId=SPREADSHEET_ID, range='A4:H27', valueInputOption='USER_ENTERED',
               body={'values': temp_table})
        .execute()
    )
  except HttpError as err:
    print(err)


if __name__ == "__main__":
  main()
