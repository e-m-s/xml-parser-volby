import pandas # table management
import xml.etree.ElementTree as ET # XML parsing
import os # find files

# Settings:
OUTPUT_PATH = 'vystup.xlsx'
ELEMENTS_PREFIX = '{http://www.volby.cz/ps/}'

# All XML files in INPUT_DIRECTORY will be parsed
INPUT_DIRECTORY = '.'


# Parse XML file formatted like: 'https://www.volby.cz/pls/ps2021/vysledky_okres?nuts=CZ0722'
def parsefile(path, writer):
  # Load XML elements tree
  tree = ET.parse(path)
  root = tree.getroot()

  # Create table
  resulttable = pandas.DataFrame()
  okresnumber = "nenastaveno"
  for child in root:
    # DEBUG: print(element.tag, ":")
    if child.tag == ELEMENTS_PREFIX+'OBEC':
      resulttable = parseobec(resulttable, okresnumber, child)
    # OKRES element should be only once.
    elif child.tag == ELEMENTS_PREFIX+'OKRES':
      okresnumber = child.get('NUTS_OKRES')

  # Print overview and export to Excel
  print(resulttable)
  resulttable.to_excel(writer, sheet_name="Okres"+okresnumber, index=False)

def parseobec(table, okresnumber, elemobec):
    # Read UCAST subelement (should be only one)
    elemucast = elemobec.find(ELEMENTS_PREFIX+'UCAST')
    if elemucast != None:
      validvotes = elemucast.get('PLATNE_HLASY')
    else:
      validvotes = 'neznámý'
    # Read elections results:
    obecnumber = elemobec.get('CIS_OBEC')
    obecname = elemobec.get('NAZ_OBEC')
    for elemvotes in elemobec.findall(ELEMENTS_PREFIX+'HLASY_STRANA'):
        table = table.append({
                'Okres číslo': okresnumber,
                'Číslo okrsku': obecnumber,
                'Název obce': obecname,
                'Strana': elemvotes.get('KSTRANA'),
                'Hlasy': elemvotes.get('HLASY'),
                'Počet platných hlasů v obci': validvotes,
                'Procenta': elemvotes.get('PROC_HLASU')
            }, ignore_index=True)
          
    return table


# iterate over XML files in input directory
writer = pandas.ExcelWriter('vystup.xlsx', engine='openpyxl')
for filename in os.listdir(INPUT_DIRECTORY):
    path = os.path.join(INPUT_DIRECTORY, filename)
    # checking if it is a file with .XML extension
    if os.path.isfile(path) and path.endswith(".xml"):
        print("Parsing: "+path+"...")
        parsefile(path, writer)
writer.save()
