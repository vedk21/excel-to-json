# imports
import xlrd
from collections import OrderedDict
import simplejson as json

# Open the workbook and select the first worksheet
wb = xlrd.open_workbook('./assets/En_to_Ja_Translation.xlsx')
sh = wb.sheet_by_index(0)

# List to hold dictionaries
language_en = OrderedDict()
language_ja = OrderedDict()

# Iterate through each row in worksheet and fetch values into dict
for rownum in range(1, sh.nrows):
    row_values = sh.row_values(rownum)
    key = row_values[0].upper().replace(" ", "_")
    language_en[key] = row_values[0]
    language_ja[key] = row_values[1]

# Serialize the list of dicts to JSON
json_en = json.dumps(language_en)
json_ja = json.dumps(language_ja)

# Write to file
with open('./assets/en.json', 'w') as f:
    f.write(json_en)

with open('./assets/ja.json', 'w') as f:
    f.write(json_ja)
