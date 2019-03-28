# -*- coding: utf-8 -*-
"""
Created on Wed Mar 20 10:24:47 2019

@author: chkiran
"""
import xml as xm
import pandas as pd
import numpy as np
import xlsxwriter as xw



class MyWorkBook(object):


    def __init__(self,input_file,output_file):
        self.input_file = input_file
        self.output_file = output_file
        file = open(input_file, mode='r')
        self.xml_string = file.read()
        file.close()
        self.xml_string = self.pre_parse_xml(self.xml_string)
        self.my_workbook = self.create_workbook(self.xml_string)


    def create_row(self,xml_row_node):
        """Returns a single row(type :Pandas Series)"""
        list_of_datavalues = []
        xml_data_list = xml_row_node.getElementsByTagName('Data')
        for xml_data_node in xml_data_list:
            list_of_datavalues.append(xml_data_node.childNodes[0].nodeValue)
            pandas_row = pd.Series(np.asarray(list_of_datavalues))
        return pandas_row


    def create_table(self,xml_worksheet_node):
        """Returns a single table(type :Pandas DataFrame)"""
        pandas_row_list = []
        xml_row_node_list = xml_worksheet_node.getElementsByTagName('Row')
        for xml_row_node in xml_row_node_list:
            pandas_row_list.append(self.create_row(xml_row_node))
        pandas_table = pd.DataFrame(pandas_row_list)
        return pandas_table


    def change_names_types(self,pandas_table):
        """Change column names to first row values and deletes first row
            Changes columns that contain all numeric values to numeric type """
        column_names = pandas_table.iloc[0]
        pandas_table = pandas_table.iloc[1:, ].rename(columns=column_names)
        pandas_table = pandas_table.apply(pd.to_numeric, errors='ignore')
        return pandas_table


    def create_workbook(self,xml_string):
        """Workbook contains worksheets
            Each Worksheet contains a Table and Table name
            Worksheet[0] has table and Worksheet[1] has Thread name
            Worksheet[3] has process name"""
        workbook = []
        domtree = xm.dom.minidom.parseString(xml_string)
        xml_worksheetnode_list = domtree.getElementsByTagName('Worksheet')
        for xml_worksheet_node in xml_worksheetnode_list:
            pandas_table = self.create_table(xml_worksheet_node)
            pandas_table = self.change_names_types(pandas_table)
            name = xml_worksheet_node.getAttribute('ss:Name')
            if name.find(r'(internel)') == -1:
                name = self.short_table_name(name)
                worksheet = [pandas_table, name[0], name[1]]
                # Name[0] = thread name ,Name[1] process name
                workbook.append(worksheet)
        return workbook


    def add_index(self,workbook):
        table_names = []
        hyperlinks_list = []
        peak_values = []
        pool_sizes = []
        used_percent = []
        process_names = []
        i = 1
        for worksheet in workbook:
            name = worksheet[1]
            hyperlink = r'=HYPERLINK("#'+name+r'!B'+str(i)+r'",'+r'"link")'
            hyperlinks_list.append(hyperlink)
            table_names.append(worksheet[1])
            process_names.append(worksheet[2])
            peak_values.append(worksheet[0]['Peak'].max())
            pool_sizes.append(worksheet[0]['Pool Size'].max())
            percentage = worksheet[0]['Peak'].max()/worksheet[0]['Pool Size'].max()*100
            used_percent.append(round(percentage, 2))
            i = i+1
        index_sheet = pd.DataFrame({'Index': table_names, 'Process name': process_names,
                                    'Hyperlinks': hyperlinks_list,
                                    'Peaks': peak_values, 'Pool size': pool_sizes,
                                    'Used percent': used_percent})
        workbook.insert(0, [index_sheet, "Index",""])
        return workbook

    def short_table_name(self,name):
        """Reduses length of Table name to write to excel"""
        split_names = name.split('____')
        thread_name = split_names[1]
        process_name = split_names[0]
        return  thread_name, process_name


    def get_excel_writer(self,workbook):
        writer = pd.ExcelWriter(self.output_file,engine = 'xlsxwriter')
        for worksheet in workbook:
            worksheet[0].to_excel(writer, sheet_name=worksheet[1], index=False)
        return writer



    def pre_parse_xml(self,xml_string):
        """ Remove tags that cause errors before parsing"""
        xml_string = xml_string.replace(r'<Category="rprof.profile.stat"/>', '')
        xml_string = xml_string.replace(r'<Category="rprof.profile.mem.context"/>', '')
        xml_string = xml_string.replace(r'<Category="name"/>', '')
        return xml_string


    def add_details_sheet(self, workbook):
        details_table = pd.DataFrame(columns = ['Process name','Pool size', 'count'])
        details_table['count'] = workbook[0][0].groupby('Process name')['Process name'].count()
        details_table['Pool size'] = workbook[0][0].groupby('Process name')['Pool size'].sum()
        details_table['Process name'] = details_table.index
        workbook.insert(1, [details_table, "Details", ""])
        return workbook


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
