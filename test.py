import logging

logging.basicConfig(format='%(asctime)s-%(levelname)s-%(message)s', level=logging.INFO, filename='app.log')

if __name__ == '__main__':
    account_dict = {'Alto': 'ALTO_FUT - GSFO',
                    'Bainbridge': 'ANANMN - GSFO',
                    'Neutral': 'NEUTRAL - UBSD',
                    'Boothbay': 'BB_INTL - GSFO',
                    'Gold': 'ALTO_GLD - GSFO'}

    if 'Alto' in account_dict:
        print(account_dict.get('Alto'))

