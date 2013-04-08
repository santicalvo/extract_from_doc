# -*- coding: utf-8 -*- 
import sys, os
import string
import office.apps.extractor.doc_extractor as doc_extractor
#import guion_parser.guion_videos as parser
from curso_toxml.xmlcreator import *
#from guion_parser.guion_videos_chapisteria import *
from guion_parser import *


def DocAxml(caps):
    for cap in caps:
    #if cap == "curso1":
        paginas = caps[cap]
        #curso_xml = CursoSeatXml()
        curso_xml = CursoSeatXmlMinidom()
        for pagina in paginas:
            num_pagina = pagina["np"]
            texto = pagina["txt"]
            titul = pagina["titul"]
            #print num_pagina, texto
            curso_xml.addPagina(num_pagina, texto, titulo=titul)
        path = "../files/xml/"+cap+".xml"
        curso_xml.save(path)

#modificar con cada nuevo curso
def getParser(path):
    return guion_videos_chapisteria.LectorGuionChapisteria(path)



if __name__ == '__main__':
    path = os.path.join(os.getcwd(),"../files/guion_video_chapisteria.docx")
    #path = os.path.join(os.getcwd(),"../files/tests1.docx")
    try:
        parser = getParser(path)
        parser.doc_to_xmls( xml_path="../files/xml/" )
        #lector = parser( path )
        #caps = parser.separaCapitulos()
        #print caps
        #DocAxml( caps )
        #print caps
        #lector =  parser.LectorGuionVideos(path)
        #lector =  parser.LectorGuionChapisteria(path)
        #caps = lector.separaCapitulos()
        #print caps
        #DocAxml( caps )

    except doc_extractor.WordNotFoundException as docNotFound:
        print "file not found: ", docNotFound
#       except Exception as ex:
#       raise ex
#    except parser.ParserError as parseError:
#        print parseError














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