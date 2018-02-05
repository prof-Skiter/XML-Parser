# -*- coding: utf-8 -*-
"""
Created on Fri Apr 28 16:46:41 2017

@author: skiter
"""
import xml.etree.ElementTree as ET
import pandas as pd

def fixtag(namespace, tag):
    return '{' + namespace + '}' + tag

def strtofloat(string):
    try:
        return float(string)
    except:
        return string

eNBs = ['150566', '152909', '152908', '158287', '158288', '158306', '62559', '63375', '64404', '66532', '66533', '66528', '73809', '72906', '73810', '79017', '76848', '77102', '79049', '79033', '77403', '79862']
files =['20170919_AU_LNBTS_UI.xml', '20170926_AU_LNBTS_UI.xml']
columns = ['dn', 'parameter', files[0], files[1]]
df = pd.DataFrame(columns=columns)
df = df.set_index(['dn', 'parameter'])

writer = pd.ExcelWriter('output.xlsx')

for eNB in eNBs:
    print(eNB)    
    for file in files:
        print(file)
        for event, elem in ET.iterparse(file, events=('start', 'end', 'start-ns', 'end-ns')):
            if event == 'start-ns':
                ns, url = elem
            if event == 'end' and elem.tag == fixtag(url, 'managedObject') and elem.attrib["distName"].find(eNB) != -1:
                if elem.attrib["class"] == "LNADJ" or elem.attrib["class"] == "LNADJL" or elem.attrib["class"] == "LNREL": break
                for p in elem.findall(fixtag(url, 'p')):
                    df.loc[(elem.attrib["distName"], elem.attrib["class"] + ':' + p.attrib["name"]), file] = strtofloat(p.text)
        
                elem.clear()
    df.to_excel(writer, eNB)
    df.drop(df.index, inplace=True)
                        
writer.save()