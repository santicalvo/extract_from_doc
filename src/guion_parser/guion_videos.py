# -*- coding:utf-8 -*-
import re
import xml.etree.ElementTree as xml
'''

Aqui guardaremos diferentes funciones que permitan parsear cada tipo de documento en concreto
Deberíamos conseguir que siempre se llamaran a estas funciones con listas de datos y que estas se
encarguen de reconocer los datos de dicha lista y organizarlos de la forma adecuada.

Cada guión es distinto. No consigo desgranar una forma que me permita parsear todas las estructuras 
de tablas que nos envíen. Esto no sirve para ser reusado, sólo como script rápido para generar los 
documentos
'''
class DocToXml(object):
    def __init__(self):
        self.root = xml.Element("paginas")
        
    def addPagina(self, num_pagina):
        pagina = xml.Element("pagina")
        pagina.attrib["num"] = num_pagina
        self.root.append(pagina)
        
    def save(self,nom):
        nom = nom+".xml"
        file.open(nom, "w")
        xml.ElementTree.write(file)
        file.close()
        
    
    
class LectorGuionVideos(object):
    def __init__(self):
        self.re_pantalla = re.compile("(.*)([0-9]{1,2}\.[0-9]{1,2}\.[0-9]{1,2})(.*)")
        self.re_indice = re.compile(".*ndice.*")
        
    def __detect_pantalla(self, txt):
        pantalla = re.match(self.re_pantalla, txt)
        if pantalla:
            return pantalla.group(2)
        indice = re.match(self.re_indice, txt)
        if indice:
            return "indice"
        return False
    
    def parse_table_data(self,table):
        i = j = 0
        if len(table) > 1:
            fila = 1
        elif len(table) == 1:
            fila = 0
        columna = 0
        match = self.__detect_pantalla(table[fila][columna].Range.Text[:-1])
        if match:
            txt = table[fila][columna+1].Range.Text[:-1]
        return {
              "num": match,
              "txt":txt
            }
        
    def print_table(self,table):
        i = j = 0
        for rows in table:
            for cell in rows:
                print i, j, cell.Range.Text[:-1], self.detect_pantalla(cell.Range.Text[:-1])
                j+=1
            j=0
            i+=1
        print "---- fin table -----"
