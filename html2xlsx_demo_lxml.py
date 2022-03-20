# -*- coding: utf-8 -*-
#%% [markdown]
# html2xlsx `lxml.etree.iterparse` and `xlsxwriter` Demo
#%%
from lxml import html, etree
import xlsxwriter

#%%
url = 'demo.html'
# content = html.fromstring(open(url, encoding='utf-8').read())

# %%
# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook('demo.xlsx')
worksheet = workbook.add_worksheet()

row = 0
for event, element in etree.iterparse(url, tag='TR', events=('end',)):
    print(row)
    for col, child in enumerate(element):
        # print(child.tag, child.text)
        worksheet.write(row, col, child.text)
    element.clear()
    row += 1
    if row > 1:
        break

workbook.close()