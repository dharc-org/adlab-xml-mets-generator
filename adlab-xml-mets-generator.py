import json
import sys
from openpyxl import load_workbook

# check arguments
if len(sys.argv) < 3:
    print("No parameter has been included")
    print(" 1) path of json formatted config file i.e. config.json")
    print(" 1) path of xlsx file with metadata")
    print("i.e. $ python adlab-xml-mets-generator.py config.json metadata.xlsx")
    sys.exit()

# get var from arguments
config = sys.argv[1]
xslxfile = sys.argv[2]

# get json config file and parse data
with open(config, "r") as f:
    configVars = json.load(f)



# open xlsx in read mode
wb = load_workbook(filename=xslxfile, read_only=True)
ws = wb['Sheet1']

for row in ws.rows:
    for cell in row:
        print(cell.value)

# Close the workbook after reading
wb.close()
