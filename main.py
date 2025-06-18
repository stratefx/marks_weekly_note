import string
import random
import pandas as pd
import os
import datetime
import logging
import os.path
from config import KEY_LIST, HEADER, MONTH_DICT, chat_GPT_magic_8_ball

logger = logging.getLogger(__name__)
logging.basicConfig(filename='Marks_note.log', level=logging.INFO)



email = " --EMAIL GOES HERE-- " # TODO: Paste email here, ensuring not to forget the date in the upper right



"""
Removes punctuation and special characters from email and segregates words into a list for later counting
Returns a list named email
"""
def clean_email(email):
    try:
        delete_punctuation_dict = {sp_character: '' for sp_character in string.punctuation}
        translation_table = str.maketrans(delete_punctuation_dict)
        email = email.translate(translation_table).lower().split()

        m = email[0]
        d = email[1]
        y = email[2]
        date_string = m + ' ' + d + ' ' + y
        date_obj = datetime.datetime.strptime(date_string, '%B %d %Y')

        return email, date_obj
    except:
        logger.error(f'Error: ', exc_info=True)
        exit(1)



"""
Iterates through the list of words returned by clean_email, searches for matches to KEY_LIST of
buzz words most often used in Mark's emails, then saves the count for each word in a dictionary
"""
def count_words(email):
    try:
        d = {key: 0 for key in KEY_LIST}
        prev = ''
        for word in email:
            if word == 'focused' and prev == 'lazer':
                if d['lazer focused'] > 0:
                    d['lazer focused'] = d['lazer focused'] + 1
                else:
                    d['lazer focused'] = 0
                if d['lazer'] > 0:
                    d['lazer'] = d['lazer'] - 1
                else:
                    d['lazer'] = 0
            elif word in d:
                d[word] = d[word] + 1
            prev = word

        return d
    except:
        logger.error(f'Error: ', exc_info=True)
        exit(1)



"""
Utilizes 2 csv files as a datasource. One stored to the desktop named weekly_metrics.csv, which windows OS will convert 
to a spreadsheet, and the second locally as a csv in the program files named metrics_csv.csv. First checks OS 
if csv file exists and creates a new one if not. Then updates csv with current metrics for each word by iterating 
through list of buzz words (KEY_LIST)
"""
def update_csv(current_count_dict, date_obj):
    try:
        # create dynamic path names that change and create new file every Jan 1 for historical data
        year = str(date_obj.year)
        CSV_PATH = os.getenv('CSV_PATH')
        CSV_PATH = CSV_PATH.replace('.',f'_{year}.')

        CSV_LOCAL_PATH = f'metrics_csv_{year}_local.csv'

        # If csv file does not exist, such as on first run, create it with HEADER
        if not os.path.exists(CSV_PATH):
            logger.info(f'CSV file being created. Saved to desktop at location: ', CSV_PATH)
            with open(CSV_PATH, 'w') as file:
                file.write(HEADER)

            logger.info(f'CSV file being created. Saved to location: metrics_csv_{year}.csv in project hierarchy tree.')
            with open(CSV_LOCAL_PATH, 'w') as file:
                file.write(HEADER)

        csv_df = pd.read_csv(CSV_PATH)

        # csv file should only be empty first time the program runs
        if csv_df.empty:
            logger.info('First run. Historical data pending. Rocket thrusters engaged.')
            counter = 0

            for item in current_count_dict.items():
                csv_df.loc[counter, 'word'] = item[0]
                csv_df.loc[counter, 'date'] = date_obj.strftime('%B-%d-%Y')
                csv_df.loc[counter, 'current_count'] = item[1]
                csv_df.loc[counter, 'last_week_count'] = 0
                csv_df.loc[counter, 'trending'] = 'No trend data available at this time'
                csv_df.loc[counter, 'month'] = item[1]
                csv_df.loc[counter, 'quarter'] = item[1]
                csv_df.loc[counter, 'year'] = item[1]

                # Loops months
                for i in MONTH_DICT:
                    if MONTH_DICT[i] == date_obj.strftime("%B"):
                        csv_df.loc[counter, MONTH_DICT[i]] = item[1]
                    else:
                        csv_df.loc[counter, MONTH_DICT[i]] = 0

                csv_df.loc[counter, 'month'] = date_obj.month
                csv_df.loc[counter, 'quarter'] = (date_obj.month - 1) // 3 + 1
                csv_df.loc[counter, 'year'] = date_obj.year
                csv_df.loc[counter, 'month_name'] = date_obj.strftime("%B")

                csv_df.loc[counter, 'ChatGPT'] = chat_GPT_magic_8_ball.get(random.randint(1, 13))

                counter += 1
        else:
            counter = 0

            for item in current_count_dict.items():
                csv_df.loc[counter, 'word'] = item[0]
                csv_df.loc[counter, 'date'] = date_obj.strftime('%B-%d-%Y')
                csv_df.loc[counter, 'current_count'] = item[1]

                if counter > 0:
                    csv_df.loc[counter, 'last_week_count'] = csv_df.loc[counter - 1, 'current_count']
                else:
                    csv_df.loc[counter, 'last_week_count'] = csv_df.loc[counter, 'current_count']
                if csv_df.loc[counter, 'current_count'] == csv_df.loc[counter, 'last_week_count']:
                    csv_df.loc[counter, 'trending'] = 'Stable'
                elif csv_df.loc[counter, 'current_count'] > csv_df.loc[counter, 'last_week_count']:
                    csv_df.loc[counter, 'trending'] = 'Upward'
                else:
                    csv_df.loc[counter, 'trending'] = 'Downward'

                # Loops months
                for i in MONTH_DICT:
                    if MONTH_DICT[i] == date_obj.strftime("%B"):
                        csv_df.loc[counter, MONTH_DICT[i]] = item[1] + csv_df.loc[counter, MONTH_DICT[i]]

                csv_df.loc[counter, 'month'] = date_obj.month
                csv_df.loc[counter, 'quarter'] = (date_obj.month - 1) // 3 + 1
                csv_df.loc[counter, 'year'] = date_obj.year
                csv_df.loc[counter, 'month_name'] = date_obj.strftime("%B")

                csv_df.loc[counter, 'ChatGPT'] = chat_GPT_magic_8_ball.get(random.randint(1, 13))

                counter += 1

        csv_df.to_csv(CSV_PATH, index=False) # TODO
        csv_df.to_csv(CSV_LOCAL_PATH, index=False)
    except:
        logger.error(f'Error: ', exc_info=True)
        exit(1)



current_time = datetime.datetime.now()
current_time = current_time.replace(microsecond=0)

if __name__ == '__main__':
    logger.info(f'Started: ' + str(current_time))
    email, date_obj = clean_email(email)
    count_dict = count_words(email)
    update_csv(count_dict, date_obj)
    logger.info(f'Finished: ' + str(current_time))