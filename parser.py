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
import pandas as pd

start = time()


def _match(pattern):
    if pattern:
        return pattern[0]
    return 0


home_dir = os.path.join("C:", os.environ["HOMEPATH"], "Desktop", "Vodafone")
table_dict = {}
df = pd.DataFrame()

# Constants
TRIGGER = "Usage summary"

# Regex patterns
uk_calls = r"(?<=UK calls ).*?(?=\smin)"
abroad_calls = r"(?<=Calling abroad from the UK ).*?(?=\smin)"
non_geo_calls = r"(?<=Calling non-geographic numbers ).*?(?=\smin)"
mobile_data = r"(?<=UK mobile data ).*?(?=\sMB)"
name = r"\b\w+\b\s\b[\w-]+\n07\d{9}"

files = [_ for _ in list(*os.walk(home_dir))[2] if _.endswith(".pdf")]
root_path = list(*os.walk(home_dir))[0]

for file in files:
    bill_date = file[file.find("_", 0)+1:-4]
    reader = PdfReader(os.path.join(root_path, file))
    print(f"Reading file {bill_date}")
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

print("Saving into excel...")
df = pd.DataFrame.from_dict(table_dict, orient='index')
df = df.astype({'Date': 'datetime64[ns]',
                'UK (Minutes)': 'int32',
                'International (Minutes)': 'int32',
                'Non-geographic (Minutes)': 'int32',
                'Data (MB)': 'float',
                })
df.to_excel(os.path.join(root_path, "result.xlsx"), sheet_name='Data')
print("Done!")
print(f"Time lapsed: {time() - start} seconds")
