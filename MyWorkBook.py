# -*- coding: utf-8 -*-
"""
Created on Wed Mar 20 10:24:47 2019

@author: chkiran
"""
import xml.dom.minidom as md
import pandas as pd
import numpy as np
import os




class MyWorkBook:


#    def __init__(self,inputfile,outputfile):
#        self.inputfile = inputfile
#        self.outputfile = outputfile
#        file = open(inputfile, mode='r')
#        self.xml_string = file.read()
#        file.close()
#        self.Custom_Workbook = CreateWorkbook(self.xml_string)
#        self.Custom_Workbook = AddIndex(self.Custom_Workbook)
#        self.Writer = CreateExcelWriter(self.Custom_Workbook,self.outputfile)



    def CreateRow(xml_row_node):
        """Returns a single row(type :Pandas Series)"""
        List_of_datavalues = []
        xml_data_list = xml_row_node.getElementsByTagName('Data')
        for xml_data_node in xml_data_list:
                List_of_datavalues.append(xml_data_node.childNodes[0].nodeValue)
                pandas_row = pd.Series(np.asarray(List_of_datavalues))
        return pandas_row


    def CreateTable(xml_worksheet_node):
        """Returns a single table(type :Pandas DataFrame)"""
        pandas_row_list = []
        xml_row_node_list = xml_worksheet_node.getElementsByTagName('Row')
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
        """Custom_Workbook contains Custom_worksheets
            Each Custom_worksheet contains a Table and Table name
            Custom_worksheet[0] has table and Custom_worksheet[1] has Table name"""
        Custom_Workbook = []
        DOMTree = md.parseString(xml_string)
        xml_worksheetnode_list = DOMTree.getElementsByTagName('Custom_worksheet')
        for xml_worksheet_node in xml_worksheetnode_list:
            pandas_table = CreateTable(xml_worksheet_node)
            pandas_table = ChangeNamesAndTypes(pandas_table)
            Name = xml_worksheet_node.getAttribute('ss:Name')
            Name = ShortTableName(Name)
            Custom_worksheet = [pandas_table, Name]
            Custom_Workbook.append(Custom_worksheet)
        return Custom_Workbook


    def AddIndex(Custom_Workbook):
        Table_names = []
        Hyperlinks_list = []
        i = 1
        for Custom_worksheet in Custom_Workbook:
            name = Custom_worksheet[1]
            Hyperlink = r'=HYPERLINK("#'+name+r'!B'+str(i)+r'",'+r'"link")'
            i = i+1
            Table_names.append(Custom_worksheet[1])
            Hyperlinks_list.append(Hyperlink)
        Index_sheet = pd.DataFrame({'Index': Table_names, 'Hyperlinks': Hyperlinks_list})
        Custom_Workbook.insert(0, [Index_sheet, "Index"])
        return Custom_Workbook


    def ShortTableName(name):
        """Reduses length of Table name to write to excel"""
        if r'(internel)' in name:
            Short_name = name.replace(r'(internal)', '')
        else:
            Short_name = name.replace(name.split('___')[0], '')[4:]
        return Short_name


    def CreateExcelWriter(Custom_Workbook, Output_file):
        writer = pd.ExcelWriter(Output_file)
        for Custom_worksheet in Custom_Workbook:
            Custom_worksheet[0].to_excel(writer, sheet_name=Custom_worksheet[1], index=False)
        return writer


    def PreParseXml(xml_string):
        """ Remove tags that cause errors before parsing"""
        xml_string = xml_string.replace(r'<Category="rprof.profile.stat"/>', '')
        xml_string = xml_string.replace(r'<Category="rprof.profile.mem.context"/>', '')
        xml_string = xml_string.replace(r'<Category="name"/>', '')
        return xml_string


    def GetData(inputfile,outputfile):
        file = open(inputfile, mode='r')
        xml_string = file.read()
        file.close()
        Custom_Workbook = CreateWorkbook(self.xml_string)
        Custom_Workbook = AddIndex(self.Custom_Workbook)
        Writer = CreateExcelWriter(self.Custom_Workbook,self.outputfile)


#for Custom_worksheet in Custom_worksheet_list:
#
#     Name = Custom_worksheet.getAttribute('ss:Name')
#     Row_list=Custom_worksheet.getElementsByTagName('Row')
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
