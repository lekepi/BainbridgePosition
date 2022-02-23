import win32com.client
import os
from datetime import datetime
from models import config_class, session, PositionBainb, ParentFund, Product, TaskChecker
from utils import task_checker_db, add_log_db
import pandas as pd
import logging

# https://www.codeforests.com/2021/05/16/python-reading-email-from-outlook-2/

# TODO Run at 7am for the 3 files, store in position_bain

logging.basicConfig(format='%(asctime)s-%(levelname)s-%(message)s', level=logging.INFO, filename='app.log')
BAINBRIDGE_DIR = config_class.BAINBRIDGE_PATH


def clean_ticker(ticker):
    temp = ticker
    temp = temp.replace(" Equity", "")
    temp = temp.replace(" GR", " GY")
    temp = temp.replace(" CT", " CN")
    temp = temp.replace(" CT", " CN")
    temp = temp.replace("SXO1 Index", "SXO1 EUX")
    temp = temp.replace("ES1 Index", "ES1 CME")
    temp = temp.replace("GC1 Comdty", "GC1 CMX")
    return temp


def check_path(my_date):
    year = my_date.strftime("%Y")
    month = my_date.strftime("%B")
    day = my_date.strftime("%d")
    local_path = os.path.join(BAINBRIDGE_DIR, f'Download\\{year}\\{month}\\{day}')
    if not os.path.exists(local_path):
        os.makedirs(local_path)
    return local_path


def bainbridge_all_pos_email(my_date):
    my_date_str = my_date.strftime("%Y-%m-%d")
    task_checker_list = session.query(TaskChecker).filter(TaskChecker.task_details.startswith('Bainb Position')).\
        filter(TaskChecker.task_details.endswith(my_date_str)).filter(TaskChecker.status == 'Success').all()

    filename_list = ['Boothbay', 'ALTRA AMN', 'Ananda Market Neutral', 'ALTO']
    parent_fund_list = ['Boothbay', 'Bainbridge', 'Neutral', 'Alto']
    for index, filename in enumerate(filename_list):
        search_file = f"{filename} - Morning Positions - {my_date_str}"
        fund_name = parent_fund_list[index]
        task_success = [task for task in task_checker_list if fund_name in task.task_details]
        if not task_success:
            bainbridge_pos_email(my_date, search_file, fund_name)


def bainbridge_pos_email(my_date, search_file, fund_name):

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

    start_time = my_date.replace(hour=0, minute=0, second=0).strftime('%Y-%m-%d %H:%M')
    end_time = my_date.replace(hour=23, minute=59, second=59).strftime('%Y-%m-%d %H:%M')

    messages = mapi.Folders(1).Folders(2).Items
    messages = messages.Restrict("[ReceivedTime] >= '" + start_time + "' And [ReceivedTime] <= '" + end_time + "'")
    messages.Sort("[ReceivedTime]", Descending=True)

    for message in messages:
        if message.subject.startswith(search_file):
            found_message = message

    if not found_message:
        messages = mapi.Folders(1).Folders(2).Folders(bb_num).Items
        messages = messages.Restrict("[ReceivedTime] >= '" + start_time + "' And [ReceivedTime] <= '" + end_time + "'")
        messages.Sort("[ReceivedTime]", Descending=True)
        for message in messages:
            if message.subject.startswith(search_file):
                found_message = message
                break

    if not found_message:
        task_checker_db(status='Fail', task_details=f'Bainb Position {fund_name} - {my_date_str}',
                        comment=f"File '{search_file}' not found in email",
                        task_name='Bainb Position', task_type='Task Scheduler')
        return

    for att in found_message.Attachments:
        destination = os.path.join(local_path, att.FileName)
        att.SaveAsFile(destination)
        df = pd.read_csv(destination, header=5)
        break

    # Delete position_bainb for that fund and day
    parent_fund_db = session.query(ParentFund).all()
    product_db = session.query(Product).all()
    fund_id = [fund.id for fund in parent_fund_db if fund.name == fund_name][0]
    session.query(PositionBainb).filter(PositionBainb.parent_fund_id == fund_id).\
        filter(PositionBainb.entry_date == my_date).delete()
    session.commit()

    pos_bainb_list = []
    for index, row in df.iterrows():
        ticker = row['Ticker']
        quantity = row['Position']
        notional_usd = row['Dollar Exposure']
        trade_qty = row['Previous Day Trades']
        trade_notional_usd = row[' Trades Value (USD)']

        modif_ticker = clean_ticker(ticker)
        product_matches = [prod for prod in product_db if prod.ticker == modif_ticker]
        if not product_matches:
            des = f"Position_bainb issue: the product {ticker} is unknown, the position is not added"
            add_log_db("Bainb Position", f'Bainb Position {fund_name} - {my_date_str}', "Unknown Product", des, 'Error')
        else:
            product=product_matches[0]

            new_pos_bainb = PositionBainb(entry_date=my_date,
                                          parent_fund_id=fund_id,
                                          ticker=ticker,
                                          product_id=product.id,
                                          quantity=quantity,
                                          notional_usd=notional_usd,
                                          trade_qty=trade_qty,
                                          trade_notional_usd=trade_notional_usd)
            pos_bainb_list.append(new_pos_bainb)
    if pos_bainb_list:
        try:
            session.add_all((pos_bainb_list))
            session.commit()
            task_checker_db(status='Success', task_details=f'Bainb Position {fund_name} - {my_date_str}',
                            comment=f"Position_bainb {fund_name} added",
                            task_name='Bainb Position', task_type='Task Scheduler')
        except Exception as e:
            session.rollback()
            print(e)
            logging.error("Exception occurred", exc_info=True)

            task_checker_db(status='Fail', task_details=f'Bainb Position {fund_name} - {my_date_str}',
                            comment=f"Problem to add position_bainb {fund_name} in the DB, check the log",
                            task_name='Bainb Position', task_type='Task Scheduler')
            print(f"{fund_name} - {my_date}")


if __name__ == '__main__':
    my_date = datetime.today()
    bainbridge_all_pos_email(my_date)


