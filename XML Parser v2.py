# -*- coding: utf-8 -*-
"""
Created on Tue Apr 25 16:14:59 2017

@author: skiter
"""

import xml.dom.minidom as md

cfg = md.parse("cfg.xml");
print('XML file opened')
MOs = cfg.getElementsByTagName("managedObject")
print("%d MOs in file" % MOs.length)

i = 0
data = {}
for MO in MOs:
    temp = {}
    temp['object'] = MO.getAttribute("class")
    for parameter in MO.getElementsByTagName("p"):
        temp[parameter.getAttribute("name")] = parameter.firstChild.data
    data[MO.getAttribute("distName")] = temp
    
    i += 1
    if i == 50: break

object_list = list(set([data[sample]['object'] for sample in sorted(data)]))
print(object_list)