# -*- coding: utf-8 -*-
"""
Created on Fri Apr 28 16:46:41 2017

@author: skiter
"""
import xml.etree.ElementTree as ET
import xlsxwriter
import time

def fixtag(namespace, tag):
    return '{' + namespace + '}' + tag

def strtofloat(string):
    try:
        return float(string)
    except:
        return string

result = {}
c = 0
timer = []
for event, elem in ET.iterparse("cfg.xml", events=('start', 'end', 'start-ns', 'end-ns')):
    t1 = time.time()
    if event == 'start-ns':
        ns, url = elem
 
    if event == 'end' and elem.tag == fixtag(url, 'managedObject'):
        MO = elem
        cur_class = MO.attrib["class"]
        if cur_class not in result.keys():
            result[cur_class] = {}
            result[cur_class]['row_count'] = 0
            result[cur_class]['dn'] = []
        
        result[cur_class]['dn'].append(MO.attrib["distName"])
        for p in MO.findall(fixtag(url, 'p')):

#            if p.attrib["name"] not in result[cur_class].keys():
#                result[cur_class][p.attrib["name"]] = []
#                result[cur_class][p.attrib["name"]].extend(['blank']*result[cur_class]['row_count'])
            
            result[cur_class][p.attrib["name"]] = result[cur_class].get(p.attrib["name"], ['blank']*result[cur_class]['row_count']) + [strtofloat(p.text)]
      
        temp = {}
        for list_item in MO.findall(fixtag(url, 'list')):
            if list_item.findall(fixtag(url, 'item')) == []:
                temp[list_item.attrib["name"]] = str([p.text for p in list_item.findall(fixtag(url, 'p'))])
            else:
                
                
#                if list_item.attrib["name"] not in result[cur_class].keys():
#                    result[cur_class][list_item.attrib["name"]] = []
#                    result[cur_class][list_item.attrib["name"]].extend(['blank']*result[cur_class]['row_count']) #extend(['blank' for i in range(result[cur_class]['row_count'])])
                
                result[cur_class][list_item.attrib["name"]] = result[cur_class].get(list_item.attrib["name"], ['blank']*result[cur_class]['row_count']) + ['List']
    
    
    
    
                for item in list_item.findall(fixtag(url, 'item')):
                    for p in item.findall(fixtag(url, 'p')):
                        param_name = 'Item-' + list_item.attrib["name"] + '-' + p.attrib["name"]
        
#                        if param_name not in temp.keys():
#                            temp[param_name] = []
                        
                        temp[param_name] = temp.get(param_name, []) + [p.text]

        for key in temp.keys():

#            if key not in result[cur_class].keys():
#                result[cur_class][key] = []
#                result[cur_class][key].extend(['blank']*result[cur_class]['row_count'])
                
            result[cur_class][key] = result[cur_class].get(key, ['blank']*result[cur_class]['row_count']) + [str(temp[key])]
        
        
        result[cur_class]['row_count'] += 1
           
        for check in result[cur_class].keys():
            if check != 'row_count':
                col_len = len(result[cur_class][check])
                if col_len < result[cur_class]['row_count']:
                    result[cur_class][check].append('blank')
                if (result[cur_class]['row_count'] - len(result[cur_class][check])) > 1: print('ERROR!!! Delta is', (result[cur_class]['row_count'] - len(result[cur_class][check])), 'Object param:', cur_class, check)
        
        elem.clear()
        for obj in result.keys():
            for col_name in result[obj].keys():
                if col_name != 'row_count':
                    if len(result[obj][col_name]) != len(result[obj]['dn']):
                        print('ERROR: columns len doesnt match:', obj, col_name)
                        
        t2 = time.time()
        timer.append(t2 - t1)
        c += 1
        
        
 
print("min: %.6f" % min(timer))
print("max: %.6f" % max(timer))
print("avg: %.6f" % (sum(timer)/len(timer)))
print("count: %.6f" % len(timer))
timer.sort(reverse = True)
print(timer[:30])
 


workbook = xlsxwriter.Workbook('XML Parser.xlsx', {'nan_inf_to_errors': True})
format = workbook.add_format()
format.set_rotation(90)
format.set_bold()
format.set_bg_color('#FFFF99')
format.set_align('center')

for obj in result.keys():
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
