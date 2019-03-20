# -*- coding: utf-8 -*-
"""
Created on Wed Mar 20 10:24:47 2019

@author: chkiran
"""
import xml.dom.minidom as md
import pandas 
import numpy as np
Dataframe_list = []
DOMTree = md.parse('Memory Summary_to_modify.xml')
xml_Worksheetnode_list = DOMTree.getElementsByTagName('Worksheet')




def CreateRow(xml_row_node):
    xml_data_list = xml_row_node.getElementsByTagName('Data')
    List_of_datavalues = []
    for xml_data_node in xml_data_list:
             List_of_datavalues.append(xml_data_node.childNodes[0].nodeValue)
             #print(xml_data_node.childNodes[0].nodeValue)
    pandas_row = pandas.Series(np.asarray(List_of_datavalues))
    return pandas_row

def CreateTable(xml_Worksheet_node):
    xml_row_node_list=xml_Worksheet_node.getElementsByTagName('Row')
    pandas_row_list = []
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
    table_and_name = [pandas_table,Name] 
    combo_list.append(table_and_name)

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
         