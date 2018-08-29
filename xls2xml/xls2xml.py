#!/usr/bin/env python

# -*- coding:utf-8 -*-

import sys
import os
import xlrd
import types
from xml.dom import minidom

reload(sys)
sys.setdefaultencoding( "utf-8" )

def get_unit(type):
    if type == 'string':
        return 256
    elif type == 'uint64' or type == 'int64':
        return 8
    elif type == 'uint32' or type == 'int32' or type == 'float':
        return 4
    elif type == 'uin16' or type == 'int16':
        return 2
    elif type == 'uint8' or type == 'int8':
        return 1
    else:
        return 4


def check_float(numValue):
    strValue = str(numValue)
    strList = strValue.split('.')
    
    if len(strList) == 1:
        return numValue
    elif strList[1] == '0':
        return int(strList[0])
    else:
        return numValue


def process_cell(cell_value, export_type):
    if cell_value == '':
        return cell_value

    if export_type == '' or export_type != 'string':
        if type(cell_value) == types.UnicodeType or type(cell_value) == types.StringType:
            if cell_value.find('.') == -1:
                return str(string.atoi(cell_value))
            else:
                return str(string.atof(cell_value))
        elif type(cell_value) == types.FloatType:
            return str(check_float(cell_value))
        else:
            return str(cell_value)
    
    else:
        if type(cell_value) == types.UnicodeType or type(cell_value) == types.StringType:
            return str(cell_value)
        else:
            return str(check_float(cell_value))



def xls2xml(xls_filename, meta_filename, xml_filename, xml_release_filename, name):
    
    data = xlrd.open_workbook(xls_filename)
    try:
        type_table = data.sheet_by_name(u'type')
    except:
        print("error: no type sheet")
        return

    key_list = []
    alias_list = []
    type_list = []
    empty_data = True
    uint = 0

    for i in range(type_table.nrows):
        key_list.append(type_table.cell_value(i, 0))
        alias_list.append(type_table.cell_value(i, 1))
        if type_table.ncols >= 4:
            type_list.append(type_table.cell_value(i, 3).lower())
            if type_table.cell_value(i, 3).lower() != u'reject':
                empty_data = False
                uint += get_unit(type_table.cell_value(i, 3).lower())
        else:
            type_list.append(u'reject')


    if not empty_data:
        ### write meta struct###
        doc = minidom.Document()
        root = doc.createElement('metalib')
        root.setAttribute('name', name)
        root.setAttribute('version', '2')
        root.setAttribute('tagsetversios', '1')
        doc.appendChild(root)

        struct = doc.createElement('struct')
        struct.setAttribute('name', 'CFG'+name)
        sortkey = False
        struct.setAttribute('version', '1')
        root.appendChild(struct)

        for k in range(len(alias_list)):
            if type_list[k] != u'reject':
                entry = doc.createElement('entry')
                entry.setAttribute('name', alias_list[k])
                entry.setAttribute('type', type_list[k])
                entry.setAttribute('cname', key_list[k])
                entry.setAttribute('desc', key_list[k])
                if type_list[k] == 'string':
                    entry.setAttribute('size', str(get_unit(type_list[k])))
                struct.appendChild(entry)
   
        

        print(meta_filename)
        dst_file = file(meta_filename, 'wb')
        
        doc.writexml(dst_file, addindent = '  ', newl='\n', encoding = 'utf-8')
        dst_file.close()

        table = data.sheet_by_index(0)
        count = table.nrows - 1

        doc = minidom.Document()
        root = doc.createElement(name)
        doc.appendChild(root)

        # checking type is null or not
        
        index = []

        for i in range(len(key_list)):
            found = False
            for j in range(table.ncols):
                if key_list[i] == table.cell_value(0, j):
                    index.append(j)
                    found = True
                    break
            if not found:
                index.append(-1)


        for i in range(1, table.nrows):
            element = doc.createElement('Cfg' + name)
            for j in range(len(key_list)):
                if type_list[j] != u'reject':
                    #cfg = doc.createElement(alias_list[j])
                    if index[j] != -1:
                        element.setAttribute(alias_list[j], process_cell(table.cell_value(i, index[j]), type_list[j]))
                    else:
                        # type is null
                         element.setAttribute(alias_list[j], '')
                   # element.appendChild(cfg)
                    root.appendChild(element)

        
        dst_file = file(xml_filename, 'wb')
        doc.writexml(dst_file, addindent = '  ', newl='\n', encoding = 'utf-8')
        dst_file.close()

if __name__ == '__main__':
    pwd = os.path.split(os.path.abspath(sys.argv[0]))[0]
    print("the script is at " + pwd)

    xls_dir = os.path.join(os.path.join(pwd, 'xls'))

    xml_dir = os.path.join(os.path.join(pwd, 'xml'))
    print("the xls dir is   " + xls_dir)

    meta_dir = os.path.join(os.path.join(pwd, 'meta'))
    print("the meta dir is  "+  meta_dir)


    xls_ext = '.xls'
    meta_ext = 'Meta.xml'
    xml_ext = '.xml'


    files = []
    folders = [xls_dir]
    for folder in folders:
        folders += [os.path.join(folder, x) for x in os.listdir(folder) if os.path.isdir(os.path.join(folder, x))]
        files += [os.path.relpath(os.path.join(folder, x), start = xls_dir) for x in os.listdir(folder)\
                    if os.path.isfile(os.path.join(folder, x)) and os.path.splitext(x)[1] == xls_ext]
    print(folders)
    print(files)

    for filename in files:
        if os.path.splitext(filename)[1] == xls_ext:
            name = os.path.splitext(os.path.basename(filename))[0]
            if '+' in name:
                loc = name.find('+')
                name = name[0:loc]
        xls_filename = os.path.join(xls_dir, filename)
        meta_filename = os.path.join(meta_dir, name + meta_ext)
        xml_filename = os.path.join(xml_dir, name + xml_ext)
        print(name)
        xls2xml(xls_filename, meta_filename,xml_filename,"e", name)


