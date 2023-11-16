# adlab-xml-mets-generator
Create xml with the mets format from a xlsx file with metadata

# dependency
Python openpyxl: https://pypi.org/project/openpyxl/
```sh
pip install openpyxl
```

# mets xml
Info: https://icar.cultura.gov.it/standard/standard-internazionali/mets

# example
The scripts need 3 arguments:
1) Custom config.json with personalized data
2) Excel file with metadata
```sh
$ python adlab-xml-mets-generator.py config.json metadata.xlsx
```
