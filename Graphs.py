# -*- coding: utf-8 -*-
"""
Created on Fri Mar 22 13:33:54 2019

@author: chkiran
"""


class graphs:
    def AddLineChart(writer, name, row_len):
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
        chart.set_x_axis({'name': 'Time', 'position_axis': 'on_tick'})
        chart.set_y_axis({'name': 'size'})
        chart.set_title({'name': name})
        xlworksheet.insert_chart('I2', chart)

