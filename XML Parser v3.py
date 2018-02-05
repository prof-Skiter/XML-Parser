# -*- coding: utf-8 -*-
"""
Created on Wed Apr 26 10:22:59 2017

@author: skiter
"""

import xml.dom.minidom as md
import pandas as pd
import xlsxwriter

cfg = md.parse("cfg.xml");
print('XML file opened')
MOs = cfg.getElementsByTagName("managedObject")
print("%d MOs in file" % MOs.length)

object_list = list(set([MO.getAttribute("class") for MO in MOs]))
print(object_list)

workbook = xlsxwriter.Workbook('XML Parser.xlsx', {'nan_inf_to_errors': True})

def writetoexcel(panda, name):
    worksheet = workbook.add_worksheet(name)
    for c in range(len(panda.index)):
        worksheet.write_row(0, 0, list(panda.columns))
        worksheet.write_row(c+1, 0, list(panda.iloc[c]))
    
for obj in object_list:
    table = pd.DataFrame(columns=('object','dn'))
    i = 0
    for MO in MOs:
        if MO.getAttribute("class") == obj:
            table.loc[i, 'object'] = MO.getAttribute("class")
            table.loc[i, 'dn'] = MO.getAttribute("distName")
            for parameter in MO.getElementsByTagName("p"):
                try: 
                    table.loc[i, parameter.getAttribute("name")] = float(parameter.firstChild.data)
                except:
                    table.loc[i, parameter.getAttribute("name")] = parameter.firstChild.data
            i += 1
 
    writetoexcel(table, obj)
    print(obj,'complete')
    del table

workbook.close()