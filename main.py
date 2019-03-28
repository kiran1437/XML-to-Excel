import xml.dom.minidom as md
import pandas as pd
import numpy as np
import os
import Graphs as gp       # custom class
import MyWorkBook as mb   # custom class

# import openpyxl as op

INPATH = 'full.excelrp'
OUTPATH = 'Vod_poc.xlsx'

obj = mb.MyWorkBook(INPATH,OUTPATH)
my_workbook = obj.my_workbook
my_workbook = obj.add_index(my_workbook)
my_workbook = obj.add_details_sheet(my_workbook)
writer = obj.get_excel_writer(my_workbook)
my_graphs = gp.graphs()
my_graphs.add_line_charts(writer,my_workbook)
writer.save()
