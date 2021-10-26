# -*- coding: utf-8 -*-
"""
Created on Wed Oct 27 00:11:41 2021

@author: hasee
"""
import xlwings


app = xlwings.App(visible=True,add_book=False)
wb1 = app.books.open(r"D:\Script\GitHub\OA\test1.xlsx")
wb2 = app.books.open(r"D:\Script\GitHub\OA\test2.xlsx")

sht1 = wb1.sheets['Sheet1']
sht2 = wb2.sheets['Sheet1']
sht2.range('B1:B5').options(transpose=True).value = sht1.range('A1:A5').value

wb2.save()
wb2.close()
wb1.close()
app.quit()
