# xml-parser-volby
Simple single-use script to convert czech election statistics data from XML to Excel for Jan.

How to use the script
- Input files must be in XML format according to https://www.volby.cz/pls/ps2021/vysledky_okres?nuts=CZ0722.
- Skript takes all XML files in current folder and append them to single .XSLX (Microsoft Excel) file - each "okres" to single sheet named "okres"+its ID.

Needs Python with Pandas, LXML and OpenPyXL:
- Install Python from: https://www.python.org/downloads/
- Run install using pip:
 ```shell
 pip install pandas
 pip install openpyxl
 pip install lxml
```

I give no guarantee, just as-is code.
