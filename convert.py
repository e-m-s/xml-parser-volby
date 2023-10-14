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

  # Create list to store data rows (sublists) in:
  data = []
  # Parse data and store each line to list data
  okresnumber = "nenastaveno"
  for child in root:
    # DEBUG: print(element.tag, ":")
    if child.tag == ELEMENTS_PREFIX+'OBEC':
      parseobec(data, okresnumber, child)
    # OKRES element should be only once.
    elif child.tag == ELEMENTS_PREFIX+'OKRES':
      okresnumber = child.get('NUTS_OKRES')

  # Convert array to dataframe:
  resulttable = pandas.DataFrame(
                        data=data, 
                        columns=[
                              'Okres číslo', 'Číslo okrsku', 'Název obce', 
                              'Strana', 'Hlasy','Počet platných hlasů v obci', 'Procenta'
                              ])


  # Print overview and export to Excel
  filename = os.path.basename(path).split('.')[-2] # get filename without extension
  print("Content of "+filename+" exported to sheet "+filename+":")
  print(resulttable)
  resulttable.to_excel(writer, sheet_name=filename, index=False)

def parseobec(data, okresnumber, elemobec):
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
        data.append([
                okresnumber,
                obecnumber,
                obecname,
                elemvotes.get('KSTRANA'),
                elemvotes.get('HLASY'),
                validvotes,
                elemvotes.get('PROC_HLASU')
            ])     
    return data


# iterate over XML files in input directory
writer = pandas.ExcelWriter('vystup.xlsx', engine='openpyxl')
for filename in os.listdir(INPUT_DIRECTORY):
    path = os.path.join(INPUT_DIRECTORY, filename)
    # checking if it is a file with .XML extension
    if os.path.isfile(path) and path.endswith(".xml"):
        print("Parsing: "+path+"...")
        parsefile(path, writer)
writer.close()
