# -*- coding: utf-8 -*-
"""
Created on Wed Mar 20 10:24:47 2019

@author: chkiran
"""
import xml.dom.minidom as md
import pandas as pd
import numpy as np
import os


# import openpyxl as op

In_path = 'full.excelrp'
Out_path = 'Vod_poc.xlsx'


def CreateRow(xml_row_node):
    """Returns a single row(type :Pandas Series)"""
    List_of_datavalues = []
    xml_data_list = xml_row_node.getElementsByTagName('Data')
    for xml_data_node in xml_data_list:
            List_of_datavalues.append(xml_data_node.childNodes[0].nodeValue)
            pandas_row = pd.Series(np.asarray(List_of_datavalues))
    return pandas_row


def CreateTable(xml_Worksheet_node):
    """Returns a single table(type :Pandas DataFrame)"""
    pandas_row_list = []
    xml_row_node_list = xml_Worksheet_node.getElementsByTagName('Row')
    for xml_row_node in xml_row_node_list:
        pandas_row_list.append(CreateRow(xml_row_node))
    pandas_table = pd.DataFrame(pandas_row_list)
    return pandas_table


def ChangeNamesAndTypes(pandas_table):
    """Change column names to first row values and deletes first row
        Changes columns that contain all numeric values to numeric type """
    column_names = pandas_table.iloc[0]
    pandas_table = pandas_table.iloc[1:, ].rename(columns=column_names)
    pandas_table = pandas_table.apply(pd.to_numeric, errors='ignore')
    return pandas_table


def CreateWorkbook(xml_string):
    """Workbook contains worksheets
        Each Worksheet contains a Table and Table name
        Worksheet[0] has table and Worksheet[1] has Table name"""
    Workbook = []
    DOMTree = md.parseString(xml_string)
    xml_Worksheetnode_list = DOMTree.getElementsByTagName('Worksheet')
    for xml_Worksheet_node in xml_Worksheetnode_list:
        pandas_table = CreateTable(xml_Worksheet_node)
        pandas_table = ChangeNamesAndTypes(pandas_table)
        Name = xml_Worksheet_node.getAttribute('ss:Name')
        Name = ShortTableName(Name)
        Worksheet = [pandas_table, Name]
        Workbook.append(Worksheet)
    return Workbook


def AddIndex(Workbook):
    Table_names = []
    Hyperlinks_list = []
    i = 1
    for Worksheet in Workbook:
        name = Worksheet[1]
        Hyperlink = r'=HYPERLINK("#'+name+r'!B'+str(i)+r'",'+r'"link")'
        i = i+1
        Table_names.append(Worksheet[1])
        Hyperlinks_list.append(Hyperlink)
    Index_sheet = pd.DataFrame({'Index': Table_names, 'Hyperlinks': Hyperlinks_list})
    Workbook.insert(0, [Index_sheet, "Index"])
    return Workbook


def ShortTableName(name):
    """Reduses length of Table name to write to excel"""
    if r'(internel)' in name:
        Short_name = name.replace(r'(internal)', '')
    else:
        Short_name = name.replace(name.split('___')[0], '')[4:]
    return Short_name


def SaveToExcel(Workbook, Output_file):
    writer = pd.ExcelWriter(Output_file)
    for Worksheet in Workbook:
        Worksheet[0].to_excel(writer, sheet_name=Worksheet[1], index=False)
    writer.save()


def PreParseXml(xml_string):
    """ Remove tags that cause errors before parsing"""
    xml_string = xml_string.replace(r'<Category="rprof.profile.stat"/>', '')
    xml_string = xml_string.replace(r'<Category="rprof.profile.mem.context"/>', '')
    xml_string = xml_string.replace(r'<Category="name"/>', '')
    return xml_string


file = open(In_path, mode='r')
xml_string = file.read()
file.close()
print('working directory '+os.getcwd())
xml_string = PreParseXml(xml_string)
Workbook = CreateWorkbook(xml_string)
Workbook = AddIndex(Workbook)
SaveToExcel(Workbook, Out_path)
print('complete')



#for Worksheet in Worksheet_list:
#    
#     Name = Worksheet.getAttribute('ss:Name')
#     Row_list=Worksheet.getElementsByTagName('Row')
#     Table = pandas.DataFrame()
#     
#     for Row in Row_list:
#         Data_list = Row.getElementsByTagName('Data')
#         Value_list = []
#         for Data in Data_list:
#             Value_list.append(Data.childNodes[0].nodeValue)
#         row = pandas.Series(np.asarray(Value_list))
#         print(row)
#         
#     Table.append(row,ignore_index= True)
         