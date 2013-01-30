# -*- coding: utf-8 -*- 
import sys, os
import xml.etree.ElementTree as xml
import office.apps.extractor.doc_extractor as doc_extractor
import guion_parser.guion_videos as parser

class CursoSeatXml(object):
    def __init__(self):
        self.root = xml.Element("paginas")
        
    def addPagina(self, num_pagina, texto):
        pagina = xml.Element("pagina")
        pagina.attrib["num"] = num_pagina
        nodo_texto = xml.SubElement(pagina, "texto")
        nodo_texto.text = texto
        self.root.append(pagina)
        
    def save(self,path):
        return False
        try:
            fil = open(path, "w")
            fil.write( '<?xml version="1.0"?>' )
            fil.write( xml.tostring(self.root, "utf-8") )
            fil.close()
            #print xml.tostring(self.root)
        except Exception as ex:
            print path, ex

def DocAxml(caps):
    for cap in caps:
        paginas = caps[cap]
        curso_xml = CursoSeatXml()
        for pagina in paginas:
            num_pagina = pagina["np"]
            texto = pagina["txt"]
            #print num_pagina, texto
            curso_xml.addPagina(num_pagina, texto)
        path = "../files/xml/"+cap+".xml"
        curso_xml.save(path)



if __name__ == '__main__':
    path = os.path.join(os.getcwd(),"../files/guion_video_carroceria.docx")
    #path = os.path.join(os.getcwd(),"../files/tests1.docx")
    try:
        lector =  parser.LectorGuionVideos(path)
        caps = lector.separaCapitulos()
        DocAxml( caps )
#        for key in caps:
#            print caps[key]
#            print "--- ini key %s ---" % key
#            for item in caps[key]:
#                print item
#            print "--- end key %s ---" % key
        #pantallas = lector.parseTables()
        #tablasDocAxml(pantallas)
    except doc_extractor.WordNotFoundException as docNotFound:
        print "file not found: ", docNotFound
#       except Exception as ex:
#       raise ex














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