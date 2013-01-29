# -*- coding: utf-8 -*- 
'''
Created on 13/11/2012

@author: scalvofe
'''
import os, sys
import win32com.client, re

class WordNotFoundException(Exception):
    def __init__(self, error):
        self.error = error
    def __str__(self):
        return repr(self.error)


class DocTextExtractor(object):
    def __init__(self, path, doc=None, visible=False):
        self.word = win32com.client.Dispatch("Word.Application")
        self.word.Visible = visible
        self.doc = self.readDoc(path)
        self.tables = []


    def readDoc(self,path):
        try:
            return self.word.Documents.Open(path)
        except Exception as ex:
            if re.search("could not be found", ex[2][2]):
                raise WordNotFoundException(ex[2][2])
            else:
                raise ex
    
    def readTables(self):
        for table in self.doc.Tables:
            self.tables.append(table)
        return self.tables

    def getTableCells(self,table):
        '''Por alguna extraña razón el COM lee las celdas comenzando por 1 en vez de por 0'''
        rows = len(table.Rows)
        cols = len(table.Columns)
        cells = [[table.Cell(r+1, c+1) for c in range(cols)] for r in range(rows)]
        return cells
                
            

if __name__ == '__main__':
    pass
            