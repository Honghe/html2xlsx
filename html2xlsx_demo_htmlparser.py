# -*- coding: utf-8 -*-
#%% [markdown]
# html2xlsx `HTMLParser and` `xlsxwriter` Demo
#%%
import xlsxwriter

from html.parser import HTMLParser

# Ref: https://stackoverflow.com/questions/8477627/iteratively-parsing-html-with-lxml/8484265#8484265
class MyHTMLParser(HTMLParser):
    def __init__(self, callback):
        self.finished = False
        self.in_table = False
        self.in_row = False
        self.in_cell = False
        self.current_row = []
        self.current_cell = None
        self.row_idx = 0
        self.callback = callback
        HTMLParser.__init__(self)

    def handle_starttag(self, tag, attrs):
        attrs = dict(attrs)
        if not self.in_table:
            if tag == 'table':
                self.in_table = True
        else:
            if tag == 'tr':
                self.in_row = True
            elif tag == 'td':
                self.in_cell = True

    def handle_endtag(self, tag):
        if tag == 'tr':
            if self.in_table:
                if self.in_row:
                    self.in_row = False
                    self.callback(self.row_idx, self.current_row)
                    self.current_row = []
                    self.row_idx += 1
        elif tag == 'td':
            if self.in_table:
                if self.in_cell:
                    self.in_cell = False
                    self.current_row.append(self.current_cell)
                    self.current_cell = None

        elif (tag == 'table') and self.in_table:
            self.finished = True

    def handle_data(self, data):
        if self.in_cell:
            self.current_cell = data.strip() if data else data

#%%
url = 'demo.html'

# %%
# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook('demo.xlsx')
worksheet = workbook.add_worksheet()

def write_row(row, data):
    print(f'row {row}')
    worksheet.write_row(row, 0, data)

parser = MyHTMLParser(write_row)
parser.feed(open(url, encoding='utf-8').read())

workbook.close()
