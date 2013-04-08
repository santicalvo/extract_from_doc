# -*- coding:utf-8 -*-
import xml.dom.minidom as minidom
import xml.etree.ElementTree as xml

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
            #print xml.tostring(self.root)
            fil.write( xml.tostring(self.root, "utf-8") )
            fil.close()
            #print xml.tostring(self.root)
        except Exception as ex:
            print path, ex
