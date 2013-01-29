# -*- coding: utf-8 -*- 
import sys, os
import office.apps.extractor.doc_extractor as doc_extractor

extractor = None

def parseTables():
    tables = extractor.readTables()
    for table in tables:
        cells = extractor.getTableCells(table)
        i = j = 0
        for rows in cells:
            for cell in rows:
                print i, j, cell.Range.Text[:-1]
                j+=1
            j=0
            i+=1
        print "---- fin table -----"
        #if table.Cell(1,0).Range.Text[:8] == "Pantalla":
#        if len(table.Rows) == 1 and len(table.Columns) == 2:
#            print "->",table.Cell(1, 1).Range.Text
#            print "-->",table.Cell(1, 2).Range.Text
#            print "-------------1 2----------------"
#        if len(table.Rows) == 2 and len(table.Columns) == 2:
#            print "--->",table.Cell(1, 1).Range.Text
#            print "---->",table.Cell(1, 2).Range.Text
#            print "----->",table.Cell(2, 1).Range.Text
#            print "------>",table.Cell(2, 2).Range.Text
#            print "--------------2 2---------------"

if __name__ == '__main__':
    path = os.path.join(os.getcwd(),"../files/guion_video_carroceria.docx")
    #path = os.path.join(os.getcwd(),"../files/tests1.docx")
    try:
        extractor = doc_extractor.DocTextExtractor(path)
        parseTables()
    except doc_extractor.WordNotFoundException as docNotFound:
        print "file not found: ", docNotFound
#    except Exception as ex:
#        raise ex
        



















    '''
    path = os.path.join(os.getcwd(),"../files/guion.docx")
    doc = doc_extractor.readDoc(path)
    for table in doc.Tables:
        #if table.Cell(1,0).Range.Text[:8] == "Pantalla":
        if len(table.Rows) == 1 and len(table.Columns) == 2:
            print len(table.Rows), len(table.Columns)
            print 0, 0, table.Cell(1, 1).Range.Text
            print 0, 1, table.Cell(1, 2).Range.Text
            print "-------------1 2----------------"
        if len(table.Rows) == 2 and len(table.Columns) == 2:
            print len(table.Rows), len(table.Columns)
            print 0, 0, table.Cell(1, 1).Range.Text
            print 0, 1, table.Cell(1, 2).Range.Text
            print 1, 0, table.Cell(2, 1).Range.Text
            print 1, 1, table.Cell(2, 2).Range.Text
            print "--------------2 2---------------"
    '''