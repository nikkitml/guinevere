# +-----------+
# | LIBRARIES |
# +-----------+
import os
import pandas as pd
from functions import *

# +---------------------+
# | CLASSES & FUNCTIONS |
# +---------------------+
class Automation:

    def __init__(self):

        self.automation_name = ""
        self.code_file_path = ""
        self.working_folder_path = ""

    def getAutomationName(self): return self.automation_name
    def getCodeFilepath(self): return self.code_file_path
    def getWorkingFolderpath(self): return self.working_folder_path

def getAutomationsData(path):

    df = pd.read_excel(path)

# +------+
# | MAIN |
# +------+

# CONSTANTS
automations_file_path = f"{os.environ['USERPROFILE']}\\OneDrive - Singtel\\Digital Team\\Automations\\- Shared\\GUInevere\\data\\automations.xlsx"
permissions_file_path = f"{os.environ['USERPROFILE']}\\OneDrive - Singtel\\Digital Team\\Automations\\- Shared\\GUInevere\\data\\permissions.xlsx"

automation = Automation()

# FILE / FOLDER PATHS

# READ DATA

# WINDOW

# IMAGES

# NOTEBOOK

# PYTHON PAGE

# APPS PAGE

# DASHBOARDS PAGE
