import pandas as pd
import glob
import d6tstack.combine_csv as d6tc
import d6tstack
import d6tstack.convert_xls
from pathlib import Path
import numpy as np
from d6tstack.convert_xls import XLSSniffer
from d6tstack.utils import PrintLogger
import numpy as np
import xlwt
from xlwt import Workbook

#csv's opzoeken in de file

cfg_fnames = list(glob.glob(r'C:\Users\klaas.de.proost\OneDrive - Erasmushogeschool Brussel\1. Graduaten\Graduaten Administratie\Personeel\Formatie\*.csv'))
print(cfg_fnames)

#csv's combineren en opslaan in nieuwe csv

c = d6tstack.combine_csv.CombinerCSV(cfg_fnames,sep=';')
col_sniff = c.sniff_columns()

print('all columns equal?', c.is_all_equal())
print('')
print('which columns are present in which files?')
print('')
print(c.is_column_present())
print('')
print('in what order do columns appear in the files?')
print('')
print(col_sniff['df_columns_order'].reset_index(drop=True))


print('all columns equal?', c.is_all_equal())
print('')
print('which columns are unique?', col_sniff['columns_unique'])
print('')
print('which files have unique columns?')
print('')
print(c.is_column_present_unique())

#samengevoegde panda dataframe
merged = d6tc.CombinerCSV(cfg_fnames,sep=';', columns_select_common=True).to_pandas()


def convert_totaal(val):
    """
    Convert the string number value to a float
     - Remove spaces
     - Remove commas
     - Convert to float type
    """
    new_val = val.replace(',','.').replace(' ','')
    return float(new_val)

merged['Totaal']= merged['Totaal'].astype('str').apply(convert_totaal)
lijstdocenten = merged.Docent.values
lijst= lijstdocenten.tolist()
lijst_uniek = set(lijst)
print(lijst_uniek)


# tabel maken
def tabel_excel(naamdocent):
    path_raw = Path('C:/Users/klaas.de.proost/OneDrive - Erasmushogeschool Brussel/1. Graduaten/Graduaten Administratie/Personeel/Formatie/')
    filename = str(naamdocent)+"."+"xlsx"
    path= path_raw / filename
    docent = merged[merged['Docent']== naamdocent]
    docent ['Totaal'].fillna(0)
    docent.loc['Totaal %'] = docent.sum(numeric_only=True, axis=0)
    docent.to_excel(path, index=False)
    return ()


def genereer_alles():
    for naam in lijst_uniek:
        tabel_excel(naam)
        print ('Excel gemaakt voor' +' '+ str(naam))
    return()

genereer_alles()