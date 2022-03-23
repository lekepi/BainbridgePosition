import win32com.client
import os
from datetime import datetime, timedelta
from models import config_class
from utils import task_checker_db
from BainbridgePosition import check_path
import logging


logging.basicConfig(format='%(asctime)s-%(levelname)s-%(message)s', level=logging.INFO, filename='app.log')
BAINBRIDGE_DIR = config_class.BAINBRIDGE_PATH


def get_nav_report_from_email(my_date):
    my_date_str = my_date.strftime("%Y-%m-%d")
    local_path = check_path(my_date)
    outlook = win32com.client.Dispatch('outlook.application')
    mapi = outlook.GetNamespace('MAPI')

    found_message = None
    bb_num = 0
    for index, folder in enumerate(mapi.Folders(1).Folders(2).folders):
        if folder.name == 'BainBridge Daily File':
            bb_num = index + 1
            break

    messages = mapi.Folders(1).Folders(2).Items
    messages.Sort("[ReceivedTime]", Descending=True)

    for message in messages:
        if message.subject =='Files to NAV // Auto Generated Msg':
            found_message = message
            break

    if not found_message:
        messages = mapi.Folders(1).Folders(2).Folders(bb_num).Items
        messages.Sort("[ReceivedTime]", Descending=True)
        for message in messages:
            subject = message.subject
            if message.subject == 'Files to NAV // Auto Generated Msg':
                found_message = message
                break

    if not found_message:
        task_checker_db(status='Fail', task_details=f'Bainb Nav - {my_date_str}',
                        comment=f"Nav File not found in email",
                        task_name='Bainb Nav', task_type='Task Scheduler')
        return

    for att in found_message.Attachments:
        destination = os.path.join(local_path, att.FileName)
        att.SaveAsFile(destination)

    task_checker_db(status='Success', task_details=f'Bainb Nav',
                    comment=f"Bainbridge Nav file saved",
                    task_name='Bainb Nav', task_type='Task Scheduler')


if __name__ == '__main__':
    logging.info("BainbidgePosition started", exc_info=True)
    my_date = datetime.today()
    day = timedelta(days=1)
    day3 = timedelta(days=3)
    if my_date.weekday() == 0:
        previous_date = my_date - day3
    else:
        previous_date = my_date - day

    get_nav_report_from_email(previous_date)
    