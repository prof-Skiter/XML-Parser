# -*- coding: utf-8 -*-
"""
Created on Thu Apr 27 10:24:07 2017

@author: skiter
"""

import xml.dom.minidom as md
import xlsxwriter

cfg = md.parse("20170419_AU_All_UI.xml")
print('XML file opened')
MOs = cfg.getElementsByTagName("managedObject")
print("%d MOs in file" % MOs.length)

object_list = list(set([MO.getAttribute("class") for MO in MOs]))
print(object_list)

def strtofloat(string):
    try:
        return float(string)
    except:
        return string

result = {}

for MO in MOs:
    cur_class = MO.getAttribute("class")
    if cur_class not in result.keys():
        result[cur_class] = {}
        result[cur_class]['row_count'] = 0
        result[cur_class]['dn'] = []

    result[cur_class]['dn'].append(MO.getAttribute("distName"))

    for p in [e for e in MO.childNodes if (e.nodeType == e.ELEMENT_NODE) and e.tagName == 'p']:
        try: 
            result[cur_class][p.getAttribute("name")].append(strtofloat(p.firstChild.data))
        except KeyError:
            result[cur_class][p.getAttribute("name")] = []
            result[cur_class][p.getAttribute("name")].extend(['blank' for i in range(result[cur_class]['row_count'])])
            result[cur_class][p.getAttribute("name")].append(strtofloat(p.firstChild.data))
 
    temp = {}
    for list_item in MO.getElementsByTagName("list"):
        if list_item.getElementsByTagName("item") == []:
            temp[list_item.getAttribute("name")] = str([p.firstChild.data for p in list_item.getElementsByTagName("p")])
        else:
            try:
                result[cur_class][list_item.getAttribute("name")].append('List')
            except KeyError:
                result[cur_class][list_item.getAttribute("name")] = []
                result[cur_class][list_item.getAttribute("name")].extend(['blank' for i in range(result[cur_class]['row_count'])])
                result[cur_class][list_item.getAttribute("name")].append('List')

            for item in list_item.getElementsByTagName("item"):
                for p in item.getElementsByTagName("p"):
                    param_name = 'Item-' + list_item.getAttribute("name") + '-' + p.getAttribute("name")
                    try:
                        temp[param_name].append(p.firstChild.data)
                    except KeyError:
                        temp[param_name] = []
                        temp[param_name].append(p.firstChild.data)

    for key in temp.keys():
        try:
            result[cur_class][key].append(str(temp[key]))
        except KeyError:
            result[cur_class][key] = []
            result[cur_class][key].extend(['blank' for i in range(result[cur_class]['row_count'])])
            result[cur_class][key].append(str(temp[key]))
    
    result[cur_class]['row_count'] += 1
           
    for check in result[cur_class].keys():
        if check != 'row_count':
            col_len = len(result[cur_class][check])
            if col_len < result[cur_class]['row_count']:
                result[cur_class][check].append('blank')
            if (result[cur_class]['row_count'] - len(result[cur_class][check])) > 1: print('ERROR!!! Delta is', (result[cur_class]['row_count'] - len(result[cur_class][check])), 'Object param:', cur_class, check)
#        except TypeError:
#            pass
        
for obj in object_list:
    for col_name in result[obj].keys():
        if col_name != 'row_count':
            if len(result[obj][col_name]) != len(result[obj]['dn']):
                print('ERROR: columns len doesnt match:', obj, col_name)
 
    
workbook = xlsxwriter.Workbook('XML Parser.xlsx', {'nan_inf_to_errors': True})
format = workbook.add_format()
format.set_rotation(90)
format.set_bold()
format.set_bg_color('#FFFF99')
format.set_align('center')

for obj in object_list:
    worksheet = workbook.add_worksheet(obj)
    col = 0
    for col_name in result[obj].keys():
        if col_name != 'row_count':
            worksheet.write(0, col, col_name, format)
            worksheet.write_column(1, col, result[obj][col_name])
            col += 1
    worksheet.autofilter(0, 0, len(result[obj]['dn']), len(result[obj].keys()))
    worksheet.freeze_panes(1, 1)
  
workbook.close()