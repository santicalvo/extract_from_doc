# -*- coding:utf-8 -*-
import re
import guion_videos
from curso_toxml.xmlcreator import CursoSeatXmlMinidom
#CursoSeatXmlMinidom = 

#TODO: ARREGLAR MÉTODOS!!!!!


ParserError = guion_videos.ParserError

class LectorGuionChapisteria(guion_videos.LectorGuionVideos):
    def __init__(self, path, nombre_xml="curso"):
        super(LectorGuionChapisteria, self).__init__(path, nombre_xml="curso")
        self.re_alt = re.compile(".*obser.*", re.I)
        self.alt_text = "Obser"
    
    def detect_pantalla(self, txt):
        pantalla = re.match(self.re_pantalla, txt)
        if pantalla:
            return pantalla.group(2)
        indice = re.match(self.re_indice, txt)
        if indice:
            return "indice"
        alt = re.match(self.re_alt, txt)
        if alt:
            return "alternative_regex"
        return False
    
    def parse_table_data(self,table):
        i = j = 0
        txt = ""
        if len(table) > 1:
            fila = 0
            obser = table[1][1].Range.Text[:-1]
            titulo = ""
        elif len(table) == 1:
            fila = 0
            titulo = ""
            obser = ""
        columna = 0
        match = self.detect_pantalla(table[fila][columna].Range.Text)
        #match = self.detect_pantalla(table[fila][columna].Range.Text[:-1])
        if match:
            txt = table[fila][columna+1].Range.Text[:-1]
       
        return {
              "num": match,
              "audio": txt,
              "Titulo1": titulo,
              "observacion": obser
            }
    def get_creador_xml(self):
        return CursoSeatXmlMinidomChapisteria()

        
#Escribimos xml
class CursoSeatXmlMinidomChapisteria(CursoSeatXmlMinidom):
    def addPantalla(self, tag="pagina", num_pagina=""):
        pagina = self.doc.createElement(tag)
        if num_pagina != "":
            pagina.setAttribute("num", num_pagina)
        return pagina
        
    def addNodo(self, nom, texto):
        if texto != "":
            nodo_texto = self.doc.createElement(nom)
            ptext = self.limpia_saltos(texto) 
            nodo_texto.appendChild(ptext)
            return nodo_texto
        return None

    
    