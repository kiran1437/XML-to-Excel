# -*- coding: utf-8 -*-
"""
Created on Fri Mar 22 13:33:54 2019

@author: chkiran
"""
import pandas as pd
import numpy as np
import xlsxwriter as xw

class graphs:
    def __init__(self):
        self = self


    def add_line_chart(self, writer, name, row_len):
        xlworkbook = writer.book
        xlworksheet = writer.sheets[name]
        row_len = str(row_len)
        chart = xlworkbook.add_chart({'type': 'line'})
        chart.add_series({
            'values': '='+name+'!$D$2:$D$'+row_len+'\'',
            'name': 'used'
            })
        chart.add_series({
            'values': '='+name+'!$E$2:$E$'+row_len+'\'',
            'name': 'pool size'
            })
        chart.set_x_axis({'name': 'Time'})
        chart.set_y_axis({'name': 'size'})
        chart.set_title({'name': name})
        xlworksheet.insert_chart('I2', chart)


    def add_line_charts(self,writer,my_workbook):
        for my_worksheet in my_workbook:
            name = my_worksheet[1]
            row_len = len(my_worksheet[0])+1
            self.add_line_chart(writer, name, row_len)





