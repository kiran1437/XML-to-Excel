# -*- coding: utf-8 -*-
"""
Created on Wed Mar 20 10:24:47 2019

@author: chkiran
"""
import xml.dom.minidom as md
import pandas
import numpy as np
##import openpyxl as op

Input_file_path = 'Memory Summary_to_modify.xml'
Output_file_path = 'Vod_poc.xlsx'
DOMTree = md.parse(Input_file_path)
xml_Worksheetnode_list = DOMTree.getElementsByTagName('Worksheet')


def CreateRow(xml_row_node):
    List_of_datavalues = []
    xml_data_list = xml_row_node.getElementsByTagName('Data')
    for xml_data_node in xml_data_list:
            List_of_datavalues.append(xml_data_node.childNodes[0].nodeValue)
            pandas_row = pandas.Series(np.asarray(List_of_datavalues))
    return pandas_row


def CreateTable(xml_Worksheet_node):
    pandas_row_list = []
    xml_row_node_list = xml_Worksheet_node.getElementsByTagName('Row')
    for xml_row_node in xml_row_node_list:
        pandas_row_list.append(CreateRow(xml_row_node))
    pandas_table = pandas.DataFrame(pandas_row_list)
    return pandas_table


pandas_table_list = []
combo_list = []
for xml_Worksheet_node in xml_Worksheetnode_list:
    pandas_table = CreateTable(xml_Worksheet_node)
    pandas_table_list.append(pandas_table)
    Name = xml_Worksheet_node.getAttribute('ss:Name')
    table_and_name = [pandas_table, Name]
    combo_list.append(table_and_name)


writer = pandas.ExcelWriter(Output_file_path)

for itr in combo_list:
    print(itr[1])
    temp = itr[1].replace(itr[1].split('___')[0], '')[3:] + '_'
    print(temp)
    itr[0].to_excel(writer, sheet_name=temp, index=False)
writer.save()

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
         