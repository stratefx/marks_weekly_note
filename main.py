import string
import pandas as pd
import os
import datetime
import logging

# Importing required library
import pygsheets


# import os.path
# from google.auth.transport.requests import Request
# from google.oauth2.credentials import Credentials
# from google_auth_oauthlib.flow import InstalledAppFlow
# from googleapiclient.discovery import build
# from googleapiclient.errors import HttpError
# # If modifying these scopes, delete the file token.json.
# SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
# # The ID and range of a sample spreadsheet.
# SAMPLE_SPREADSHEET_ID = "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms"
# SAMPLE_RANGE_NAME = "Class Data!A2:E"




logger = logging.getLogger(__name__)
logging.basicConfig(filename='Marks_note.log', level=logging.INFO)



email = "Data* preprocessing is an energizing important task in energized text classification. " \
        "With the emergence of Python in the field of energized data science, it is " \
        "essential to have certain aligned shorthands to have the upper hand among " \
        "others. This article discusses ways aligned aligned to count words in a sentence, " \
        "it starts with space-separated words but also includes synergy ways to in " \
        "presence of special energized characters as well. Let’s discuss certain ways " \
        "to perform this."

key_list = ['energizing', 'energized', 'synergy', 'aligned', 'lazer']
header = 'word,current_week,last_week,trending,lazer'
# TODO: lazer focused




def clean_email(email):
    delete_punctuation_dict = {sp_character: '' for sp_character in string.punctuation}
    translation_table = str.maketrans(delete_punctuation_dict)
    email = email.translate(translation_table).lower().split()

    return email


def count_words(email):
    d = {key: 0 for key in key_list}
    for word in email:
        if word in d:
            d[word] = d[word] + 1
    return d


def update_csv(current_count_dict):
    csv_path = 'weekly_metrics.csv'

    # If csv file does not exist, such as on first run, create it with header
    if not os.path.exists(csv_path):
        with open(csv_path, 'w') as file:
            file.write(header)

    csv_df = pd.read_csv(csv_path)

    # csv file should only be empty first time the program runs
    if csv_df.empty:
        counter = 0
        for item in current_count_dict.items():
            csv_df.loc[counter, 'word'] = item[0]
            csv_df.loc[counter, 'current_week'] = item[1]
            csv_df.loc[counter, 'last_week'] = 0
            if csv_df.loc[counter, 'current_week'] == csv_df.loc[counter, 'last_week']:
                csv_df.loc[counter, 'trending'] = 'Stable'
            elif csv_df.loc[counter, 'current_week'] > csv_df.loc[counter, 'last_week']:
                csv_df.loc[counter, 'trending'] = 'Upward'
            else:
                csv_df.loc[counter, 'trending'] = 'Downward'
            counter += 1
            csv_df.loc[counter, 'lazer'] = 0
    else:
        counter = 0
        for item in current_count_dict.items():
            csv_df.loc[counter, 'word'] = item[0]
            csv_df.loc[counter, 'current_week'] = item[1]
            if counter > 0:
                csv_df.loc[counter, 'last_week'] = csv_df.loc[counter - 1, 'current_week']
            else:
                csv_df.loc[counter, 'last_week'] = csv_df.loc[counter, 'current_week']
            if csv_df.loc[counter, 'current_week'] == csv_df.loc[counter, 'last_week']:
                csv_df.loc[counter, 'trending'] = 'Stable'
            elif csv_df.loc[counter, 'current_week'] > csv_df.loc[counter, 'last_week']:
                csv_df.loc[counter, 'trending'] = 'Upward'
            else:
                csv_df.loc[counter, 'trending'] = 'Downward'
            counter += 1
            csv_df.loc[counter, 'lazer'] = 0

    csv_df.to_csv(csv_path, index=False)


# def sheets():
#     """Shows basic usage of the Sheets API.
#     Prints values from a sample spreadsheet.
#     """
#     creds = None
#     # The file token.json stores the user's access and refresh tokens, and is
#     # created automatically when the authorization flow completes for the first
#     # time.
#     if os.path.exists("token.json"):
#         creds = Credentials.from_authorized_user_file("token.json", SCOPES)
#     # If there are no (valid) credentials available, let the user log in.
#     if not creds or not creds.valid:
#         if creds and creds.expired and creds.refresh_token:
#             creds.refresh(Request())
#         else:
#             flow = InstalledAppFlow.from_client_secrets_file(
#                 "credentials.json", SCOPES
#             )
#             creds = flow.run_local_server(port=0)
#         # Save the credentials for the next run
#         with open("token.json", "w") as token:
#             token.write(creds.to_json())
#
#     try:
#         service = build("sheets", "v4", credentials=creds)
#
#         # Call the Sheets API
#         sheet = service.spreadsheets()
#         result = (
#             sheet.values()
#             .get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=SAMPLE_RANGE_NAME)
#             .execute()
#         )
#         values = result.get("values", [])
#
#         if not values:
#             print("No data found.")
#             return
#
#         print("Name, Major:")
#         for row in values:
#             # Print columns A and E, which correspond to indices 0 and 4.
#             print(f"{row[0]}, {row[4]}")
#     except HttpError as err:
#         print(err)


# TODO: Add loggers and error handling
current_time = datetime.datetime.now()
current_time = current_time.replace(microsecond=0)

# logger.info(f'Started: ' + str(current_time))
# email = clean_email(email)
# count_dict = count_words(email)
# update_csv(count_dict)
# logger.info(f'Finished: ' + str(current_time))


# sheets()


# Create the Client
# Enter the name of the downloaded KEYS
# file in service_account_file
client = pygsheets.authorize(service_account_file="C:/Users\Blisstopher\PycharmProjects\marks_weekly_note\marks_weekly_note\marksweeklynote-1d3d6bf29b7b")

# Sample command to verify successful
# authorization of pygsheets
# Prints the names of the spreadsheet
# shared with or owned by the service
# account
print(client.spreadsheet_titles())