import os
"""
CONFIG file for the OUTLOOK workflow 

Sunny ☀️   🌡️+69°F (feels +69°F, 61%) 🌬️→4mph 🌔&m Thu Jun  1 11:35:35 2023
W22Q2 – 152 ➡️ 212 – 21 ❇️ 344

"""



MYSELF = os.getenv('MYSELF')
BEEMINDER = str(os.getenv('BEEMINDER'))
BEEUSER = os.getenv('BEEUSER')
BEETOKEN = os.getenv('BEETOKEN')
BEEGOAL = os.getenv('BEEGOAL')
SNOOZE_FOLDER = os.getenv('SNOOZE_FILE_LOCATION')
WEED_TX = os.getenv('WEED_TEXT').split(",")
EXCL_FOLDERS = os.getenv('EXCLUDED_FOLDERS').split(",")

MY_HOME = os.getenv('HOME')

OUTLOOK_MSG_FOLDER = f'{MY_HOME}/Library/Group Containers/UBF8T346G9.Office/Outlook/Outlook 15 Profiles/Main Profile/Data/'

OUTLOOK_DB_FILE = f'{MY_HOME}/Library/Group Containers/UBF8T346G9.Office/Outlook/Outlook 15 Profiles/Main Profile/Data/Outlook.sqlite'
#OUTLOOK_DB_FILE = '/Users/giovanni/Desktop/Main Profile/Data/Outlook.sqlite'

WF_DATA_FOLDER = os.getenv('alfred_workflow_data')
OUTLOOK_FOLDER_KEY_FILE  = f"{WF_DATA_FOLDER}/keyfolder.json"

if SNOOZE_FOLDER:
    OUTLOOK_SNOOZER_FILE  = f"{SNOOZE_FOLDER}/snoozer.json"
else:
    OUTLOOK_SNOOZER_FILE  = f"{WF_DATA_FOLDER}/snoozer.json"
REFRESH_RATE = int(os.getenv('FOLDER_REFRESH'))

OUTLOOK_LOG_FILE  = f"{WF_DATA_FOLDER}/log.txt"


if not os.path.exists(WF_DATA_FOLDER):
    os.makedirs(WF_DATA_FOLDER)

if not os.path.exists(OUTLOOK_SNOOZER_FILE):
    with open(OUTLOOK_SNOOZER_FILE, "w") as file:
        file.write("{}")
        pass  # Creates an empty file


if not os.path.exists(OUTLOOK_LOG_FILE):
    with open(OUTLOOK_LOG_FILE, "w") as file:
        pass  # Creates an empty file