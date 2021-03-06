# -*- coding:utf-8 -*-
import re
import office.apps.extractor.doc_extractor as doc_extractor
from curso_toxml.xmlcreator import CursoSeatXmlMinidom
'''

Aqui guardaremos diferentes funciones que permitan parsear cada tipo de documento en concreto
Deberíamos conseguir que siempre se llamaran a estas funciones con listas de datos y que estas se
encarguen de reconocer los datos de dicha lista y organizarlos de la forma adecuada.

Cada guión es distinto. No consigo desgranar una forma que me permita parsear todas las estructuras 
de tablas que nos envíen. Esto no sirve para ser reusado, sólo como script rápido para generar los 
documentos
'''

#TODO: ARREGLAR MÉTODOS!!!!!

class ParserError(Exception):
    def __init__(self, error):
        self.error = error
    def __str__(self):
        return repr(self.error)
        
  
    
class LectorGuionVideos(object):
    def __init__(self, path, nombre_xml="curso"):
        self.extractor = doc_extractor.DocTextExtractor(path)
        self.re_pantalla = re.compile("(.*)([0-9]{1,2}\.[0-9]{1,2}\.[0-9]{1,2})(.*)")
        self.re_indice = re.compile(".*ndice.*")
        self.nombre_xml = nombre_xml
        self.capitulos = []
        
        
    def detect_pantalla(self, txt):
        pantalla = re.match(self.re_pantalla, txt)
        if pantalla:
            return pantalla.group(2)
        indice = re.match(self.re_indice, txt)
        if indice:
            return "indice"
        return False
    
    def num_indice(self, nums):
        return "indice"
    
    def num_alt(self, nums):
        return self.alt_text

    def numera_pantalla(self, nums):
        if nums[0] == "indice":
            return {
                    "curso": self.nombre_xml+"0", 
                    "num_pantalla": self.num_indice(nums)
                    }
        if nums[0] == "alternative_regex":
            return {
                    "curso": self.nombre_xml+"0", 
                    "num_pantalla": self.num_indice(nums)
                    }
        try:
            capitulo = "0" + nums[0] if int(nums[0]) < 10 else nums[0]
            tema = "0" + nums[1] if int(nums[1]) < 10 else nums[1]
            pantalla = "0" + nums[2] if int(nums[2]) < 10 else nums[2]
            return {
                    "curso": self.nombre_xml+nums[0],
                    "num_pantalla": capitulo+"_"+tema+"_"+pantalla
                    }
        except:
            raise ParserError( "Imposible numerar las pantallas. Error de entrada de datos .%s"% nums)

    def separaCapitulos(self):
        pantallas = self.parseTables()
        capitulos = {}
        for pantalla in pantallas:
            if pantalla["num"] is not False:
                numeracion_pantalla = self.numera_pantalla( pantalla["num"].split(".") )
                del pantalla["num"]
            try:
                nomfile = numeracion_pantalla["curso"]
                if nomfile not in capitulos.keys():
                    capitulos[nomfile] = []
                pantalla["np"] = numeracion_pantalla["num_pantalla"]
                capitulos[nomfile].append(pantalla)
            except Exception as ex:
                raise ParserError( "Imposible separar los capitulos .%s" % ex)
        return capitulos
        
    
    def parseTables(self):
        tables = self.extractor.readTables()
        pantallas = []
        for table in tables:
            cells = self.extractor.getTableCells(table)
            pantallas.append( self.parse_table_data(cells) )
        return pantallas
    
    def parse_table_data(self,table):
        i = j = 0
        txt = ""
        if len(table) > 1:
            fila = 1
            titulo = table[0][1].Range.Text[:-1]
        elif len(table) == 1:
            fila = 0
            titulo = ""
        columna = 0
        match = self.detect_pantalla(table[fila][columna].Range.Text)
        #match = self.detect_pantalla(table[fila][columna].Range.Text[:-1])
        if match:
            txt = table[fila][columna+1].Range.Text[:-1]
       
        return {
              "num": match,
              "txt": txt,
              "titul": titulo
            }
    
    def get_creador_xml(self):
        return CursoSeatXmlMinidom()
    
    def crea_xml(self, capitulo, pantallas):
        capitulo_xml = self.get_creador_xml()
        for pantalla in pantallas:
            num = pantalla["np"]
            del pantalla["np"]
            pant_nodo = capitulo_xml.addPantalla( num_pagina=num )
            for tag, tag_val in pantalla.items():
                nodo = capitulo_xml.addNodo(tag, tag_val)
                if nodo is not None:
                    pant_nodo.appendChild( nodo )
            capitulo_xml.root.appendChild( pant_nodo )
        return capitulo_xml
    
    def doc_to_xmls(self, xml_path="", capitulos=None):
        if capitulos is None:
            capitulos = self.separaCapitulos()
        for capitulo, pantallas in capitulos.items():
            xml = self.crea_xml( capitulo, pantallas )
            path = xml_path+capitulo+".xml"
            xml.save(path)
            
    def print_table(self,table):
        i = j = 0
        for rows in table:
            for cell in rows:
                print i, j, cell.Range.Text[:-1], self.detect_pantalla(cell.Range.Text[:-1])
                j+=1
            j=0
            i+=1
        print "---- fin table -----"
        
