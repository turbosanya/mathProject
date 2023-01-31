import os

from dotenv import load_dotenv

dotenv_path = os.path.join(os.path.dirname(__file__), '.env')
load_dotenv(dotenv_path)

gen_data = os.environ.get('GENERATE')
filename = os.environ.get('FILENAME')
sheet_name = os.environ.get('SHEET_NAME')
sheet_num = int(os.environ.get('SHEET_NUMBER'))
e = float(os.environ.get('E'))
d = int(os.environ.get('D'))
i = int(os.environ.get('I'))
j = int(os.environ.get('J'))


