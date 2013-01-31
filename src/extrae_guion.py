# -*- coding: utf-8 -*- 
import sys, os
import string
import xml.dom.minidom as minidom
import xml.etree.ElementTree as xml
import office.apps.extractor.doc_extractor as doc_extractor
import guion_parser.guion_videos as parser

class CursoSeatXmlMinidom(object):
    def __init__(self):
        self.doc = minidom.Document()
        self.root = self.doc.createElement("paginas")
        self.doc.appendChild(self.root)
    
    def limpia_saltos(self,texto):
        if texto.find("\n") != -1 or texto.find("\r") != -1:
            texto = texto.replace("\n", "<br />")
            texto = texto.replace("\r", "<br />")
            txt = self.doc.createCDATASection(texto)
        else:
            txt = self.doc.createTextNode(texto) 
        return txt
    
    def addPagina(self, num_pagina, texto, titulo=""):
        pagina = self.doc.createElement("pagina")
        pagina.setAttribute("num", num_pagina)
        nodo_texto = self.doc.createElement("texto")
        ptext = self.limpia_saltos(texto)
        nodo_texto.appendChild(ptext)
        if titulo !="" and titulo != " ":
            tit = self.doc.createElement("titulo")
            ttext = self.doc.createTextNode(titulo)
            tit.appendChild(ttext)
            pagina.appendChild(tit)
        pagina.appendChild(nodo_texto)
        self.root.appendChild(pagina)
        
    def save(self,path, notyet=False):
        if notyet:
            return False
        try:
            xmlstr = self.doc.toxml("utf-8")
            f=open(path, "w")
            f.write(xmlstr)
            f.close()
        except Exception as ex:
            print path, ex

class CursoSeatXmlEtree(object):
    def __init__(self):
        self.root = xml.Element("paginas")
    
    def limpia_saltos(self,texto):
        #texto = "<br />".join(texto.split("\n"))
        #texto = "<br />".join(texto.split("\r"))
        #Limpiamos un extraño carácter que nos sale al adquirir el texto del word
        #texto = filter(lambda x: x in string.printable, texto)
        texto = texto.replace("\n", "<br />")
        texto = texto.replace("\r", "<br />")
        return texto.encode("utf-8")
    
    def addPagina(self, num_pagina, texto, titulo=""):
        pagina = xml.Element("pagina")
        pagina.attrib["num"] = num_pagina
        nodo_texto = xml.SubElement(pagina, "texto")
        nodo_texto.text = self.limpia_saltos(texto)
        if titulo !="" and titulo != " ":
            tit = xml.SubElement(pagina, "titulo")
            tit.text = titulo
        self.root.append(pagina)
        
    def save(self,path, notyet=False):
        if notyet:
            return False
        try:
            fil = open(path, "w")
            fil.write( '<?xml version="1.0"?>' )
            print xml.tostring(self.root)
            fil.write( xml.tostring(self.root, "utf-8") )
            fil.close()
            #print xml.tostring(self.root)
        except Exception as ex:
            print path, ex

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



if __name__ == '__main__':
    path = os.path.join(os.getcwd(),"../files/guion_video_carroceria.docx")
    #path = os.path.join(os.getcwd(),"../files/tests1.docx")
    try:
        lector =  parser.LectorGuionVideos(path)
        caps = lector.separaCapitulos()
        DocAxml( caps )

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