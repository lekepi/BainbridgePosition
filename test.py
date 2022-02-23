import win32com.client
import os
from datetime import datetime, timedelta
import sys
# https://www.codeforests.com/2021/05/16/python-reading-email-from-outlook-2/


outlook = win32com.client.Dispatch('outlook.application')
mapi = outlook.GetNamespace('MAPI')

for idx, folder in enumerate(mapi.Folders(1).Folders(2).Folders(6)):
    print(idx+1, folder)