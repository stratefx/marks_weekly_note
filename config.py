# Key_list are words expected to be found often in email from Mark every week
KEY_LIST = ['energizing', 'energized', 'synergy', 'aligned', 'lazer', 'lazer focused']

# Header consists of metrics to track and display in dashboard
HEADER = ('word,date,current_count,last_week_count,trending,month,month_name,quarter,year,'
          'January,February,March,April,May,June,July,August,'
          'September,October,November,December,ChatGPT')

MONTH_DICT = {
    1: 'January',
    2: 'February',
    3: 'March',
    4: 'April',
    5: 'May',
    6: 'June',
    7: 'July',
    8: 'August',
    9: 'September',
    10: 'October',
    11: 'November',
    12: 'December'
}

chat_GPT_magic_8_ball = {
    1: 'Highly likely',
    2: 'Most likely',
    3: 'Probable',
    4: 'Plausible',
    5: '10 out of 10',
    6: '11 out of 10',
    7: '100%',
    8: 'Likely',
    9: 'All signs lead to yes',
    10: 'Yes',
    11: 'Without a doubt',
    12: 'Undoubtedly',
    13: 'Absolutely'
}