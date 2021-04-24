from datetime import datetime
import os

ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
FILE = 'PNs Consolidados ' + \
    datetime.now().strftime(r"%d-%m-%Y_%Hh%Mm%Ss") + '.xlsx'
NAME_FILE = os.path.join(ROOT_DIR, FILE)
