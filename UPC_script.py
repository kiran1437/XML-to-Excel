# -*- coding: utf-8 -*-
"""
Created on Wed Mar 20 10:24:47 2019

@author: chkiran
"""
import xml.dom.minidom as md
import pandas as pd
import numpy as np
##import openpyxl as op

In_path = 'Memory Summary_to_modify.xml'
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


def CreateWorkbook(Input_file):
    """Workbook contains worksheets
        Each Worksheet contains a Table and Table name
        Worksheet[0] has table and Worksheet[1] has Table name"""
    Workbook = []
    Thread_names = []
    DOMTree = md.parse(Input_file)
    xml_Worksheetnode_list = DOMTree.getElementsByTagName('Worksheet')
    for xml_Worksheet_node in xml_Worksheetnode_list:
        pandas_table = CreateTable(xml_Worksheet_node)
        pandas_table = ChangeNamesAndTypes(pandas_table)
        Name = xml_Worksheet_node.getAttribute('ss:Name')
        Thread_names.append(Name)
        Worksheet = [pandas_table, Name]
        Workbook.append(Worksheet)
    return Workbook


def ShortTableName(name):
    """Reduses length of Table name,Excel doesn't allow longer names """
    Short_name = name.replace(name.split('___')[0], '')[3:] + '_'
    return Short_name


def SaveToExcel(Workbook, Output_file):
    writer = pd.ExcelWriter(Output_file)
    for Worksheet in Workbook:
        Shortened_name = ShortTableName(Worksheet[1])
        Worksheet[0].to_excel(writer, sheet_name=Shortened_name, index=False)
    writer.save()


Workbook = CreateWorkbook(In_path)
SaveToExcel(Workbook, Out_path)




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
         