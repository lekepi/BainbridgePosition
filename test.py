import logging

logging.basicConfig(format='%(asctime)s-%(levelname)s-%(message)s', level=logging.INFO, filename='app.log')

if __name__ == '__main__':
    print("HELLO")
    logging.info("test program started", exc_info=True)
