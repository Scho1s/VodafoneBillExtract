"""
    Author         - Valerij Jurkin
    Date Created   - 20/03/2023
    Updated        - 19/12/2023
    Purpose        - script for extracting mobile data usage from vodafone invoices.
    The script iterates through pdf files in a folder, finds appropriate figures,
    and saves it into the nested dictionary. The final dictionary is converted to
    DataFrame and exported into excel file.
"""

from pypdf import PdfReader
import os
import re
from time import time
from logger import Logger
import pandas as pd

logger = Logger(__name__)

# Constants
TRIGGER = "Usage summary"
EXPORTED_FILE = 'result.xlsx'

# Regex patterns
uk_calls = r"(?<=UK calls ).*?(?=\smin)"
abroad_calls = r"(?<=Calling abroad from the UK ).*?(?=\smin)"
non_geo_calls = r"(?<=Calling non-geographic numbers ).*?(?=\smin)"
mobile_data = r"(?<=UK mobile data ).*?(?=\sMB)"
name = r"\b\w+\b\s\b[\w-]+\n07\d{9}"


def _match(pattern):
    try:
        return pattern[0]
    except Exception as e:
        logger.info(f"No matching patterns found. {e}")


def export(dict_):
    df = pd.DataFrame.from_dict(dict_, orient='index') .fillna(0)
    df = df.astype({'Date': 'datetime64[ns]',
                    'UK (Minutes)': 'int32',
                    'International (Minutes)': 'int32',
                    'Non-geographic (Minutes)': 'int32',
                    'Data (MB)': 'float',
                    })
    df.to_excel(os.path.join(root_path, EXPORTED_FILE), sheet_name='Data')


if __name__ == "__main__":
    try:
        start = time()
        home_dir = os.path.join("C:", os.environ["HOMEPATH"], "Desktop", "Vodafone")
        table_dict = {}
        df = pd.DataFrame()
        files = [_ for _ in list(*os.walk(home_dir))[2] if _.endswith(".pdf")]
        root_path = list(*os.walk(home_dir))[0]

        for file in files:
            bill_date = file[file.find("_", 0)+1:-4]
            reader = PdfReader(os.path.join(root_path, file))
            logger.info(f"Reading file {bill_date}")
            for page in reader.pages:
                text = page.extract_text()
                if TRIGGER in text:
                    _name, _number = re.search(name, text)[0].split("\n")
                    table_dict[len(table_dict)] = {'User': _name,
                                                   'Phone Number': _number,
                                                   'Date': bill_date,
                                                   'UK (Minutes)': _match(re.findall(uk_calls, text)),
                                                   'International (Minutes)': _match(re.findall(abroad_calls, text)),
                                                   'Non-geographic (Minutes)': _match(re.findall(non_geo_calls, text)),
                                                   'Data (MB)': _match(re.findall(mobile_data, text)),
                                                   }

        logger.info("Saving into excel...")
        export(table_dict)
        logger.info(f"Time lapsed: {time() - start} seconds")
    except Exception as e:
        logger.info(e)

