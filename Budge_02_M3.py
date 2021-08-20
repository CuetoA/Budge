# -*- coding: utf-8 -*-
"""
Created on Tue Jan 12 09:12:32 2021

@author: Andrés Cueto Estrada
@Programa: Módulo 3, Lector de carpetas para ubicación de archivos
            Budge 0.1.1 M3

@Descripción: Este script ejecuta los módulos M1 y M2 para poder obtener sus datos y 
    trabajar con ellos con la finalidad de leer las carpetas cargadas en M1

"""
import os 



def lector_archivos_desde_carpetas (directorio, extension_permisible):
    lista_archivos = os.listdir(directorio)
    
    archivos_xlsx = []
    for elemento in lista_archivos:
        if elemento.find(extension_permisible):
            archivos_xlsx.append(elemento)
    lista_archivos = archivos_xlsx
    
    return lista_archivos
    







