import pandas as pd
import xml.etree.ElementTree as ET
import os

nazevSouboru = 'vysledky_okres0722.xml'
oddelovac = "\t"

def parseFile(filename,writer):
  # Načtení dat
  #tree = ET.parse('https://www.volby.cz/pls/ps2021/vysledky_okres?nuts=CZ0722')
  tree = ET.parse(filename)
  root = tree.getroot()
  #data = pd.read_xml('https://www.volby.cz/pls/ps2021/vysledky_okres?nuts=CZ0722')
  #data = pd.read_xml('vysledky_okres0722.xml')
  #print(root)
  #exit()
  #cisloOkresu = "nedefinováno"
  #cisloOkresu = root[0].get('OKRES').get('NUTS_OKRES')
  #print("Okres: ", cisloOkresu)

  # Vytvoření tabulky
  tabulka = pd.DataFrame()
  #print("Obec: ", data['VYSLEDKY_OKRES']['OBEC']['CIS_OBEC'])
  #print("Hlasy: ", data['VYSLEDKY_OKRES']['OBEC']['HLASY_STRANA']['HLASY'])
  cisloOkresu = "nenastaveno"
  for element in root:
    #print(element.tag, ":")
    if element.tag.endswith('OKRES'):
      cisloOkresu = element.get('NUTS_OKRES')
      #print("\tČíslo okresu:", cisloOkresu)
    else:
      cisloObce = element.get('CIS_OBEC')
      nazevObce = element.get('NAZ_OBEC')
      #print("\t\tObec:", nazevObce, ":")
      for ucastHlasy in element:
        #print("\t\t\t", ucastHlasy.tag, ucastHlasy.attrib)
        if ucastHlasy.tag.endswith('UCAST'):
          #print("\t\t\t\t- jsem účast")
          platneHlasy = ucastHlasy.get('PLATNE_HLASY')
        elif ucastHlasy.tag.endswith('HLASY_STRANA'):
          #print("\t\t\t\t- jsem hlasy")
          hlasy = ucastHlasy
          #print(
          #    cisloOkresu, oddelovac, 
          #    cisloObce, oddelovac, 
          #    nazevObce, oddelovac,
          #    hlasy.get('KSTRANA'), oddelovac,
          #    )
          tabulka = tabulka.append({
              'Okres číslo': cisloOkresu,
              'Číslo okrsku': cisloObce,
              'Název obce': nazevObce,
              'Strana': hlasy.get('KSTRANA'),
              'Hlasy': hlasy.get('HLASY'),
              'Počet platných hlasů v obci': platneHlasy,
              'Procenta': hlasy.get('PROC_HLASU')
            }, ignore_index=True)

  # Výstup
  print(tabulka)
  vystupniSoubor = "vystup.xlsx"
  tabulka.to_excel(writer,sheet_name="Okres"+cisloOkresu, index=False)
  #with open(vystupniSoubor, 'a') as file:
  #    tabulkaAsString = tabulka.to_string(header=True, index=False)
  #    file.write(tabulkaAsString)


directory = '.'
 
# iterate over files in
# that directory
writer = pd.ExcelWriter('vystup.xlsx', engine='openpyxl')
for filename in os.listdir(directory):
    f = os.path.join(directory, filename)
    # checking if it is a file
    if os.path.isfile(f) and f.endswith(".xml"):
        print("Parsing: "+f+"...")
        parseFile(f, writer)
writer.save()
