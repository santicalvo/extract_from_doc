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
            print type(table)
            

if __name__ == '__main__':
    pass
            