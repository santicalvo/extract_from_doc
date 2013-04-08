# -*- coding:utf-8 -*-
import re
import guion_videos

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
        print table[fila][columna].Range.Text, obser
        match = self.detect_pantalla(table[fila][columna].Range.Text)
        #match = self.detect_pantalla(table[fila][columna].Range.Text[:-1])
        if match:
            txt = table[fila][columna+1].Range.Text[:-1]
       
        return {
              "num": match,
              "txt": txt,
              "titul": titulo,
              "obser": obser
            }
        
    def numera_pantalla(self, nums):
        if nums[0] == "indice":
            return {
                    "curso": self.nombre_xml+"0", 
                    "num_pantalla": self.__num_indice(nums)
                    }
        if nums[0] == "alternative_regex":
            return {
                    "curso": self.nombre_xml+"0", 
                    "num_pantalla": self.__num_indice(nums)
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
