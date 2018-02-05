# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

import xml.dom.minidom as md
import pandas as pd

cfg = md.parse("cfg.xml");
MOs = cfg.getElementsByTagName("managedObject")

print("%d MOs: " % MOs.length)

panda = pd.DataFrame(columns=('object','dn'))

i = 0
for MO in MOs:
    data = [MO.getAttribute("class"),MO.getAttribute("distName")] + [parameter.firstChild.data for parameter in MO.getElementsByTagName("p")]
    columns = ['object','dn'] + [parameter.getAttribute("name") for parameter in MO.getElementsByTagName("p")]
    temp = pd.DataFrame(columns = columns)
    temp.loc[0] = data
#    try:
    panda = panda.append(temp, ignore_index = True)
#    except:
#        pass
    i += 1
    print(i)
#    if i == 100: break
    
        
print(panda.head())