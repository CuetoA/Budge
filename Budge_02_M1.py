# -*- coding: utf-8 -*-
"""
Created on Mon Jan 11 17:55:55 2021

@author: Andrés Cueto Estrada
@Programa: Módulo 1, Ingreso de datos para trabajar

@Descripción: En este script se pueden observar todos los datos que el usuario deberá 
    ingresar para que el programa pueda trabajar, tales como:
        - Ubicación de las carpetas:
            - Carpeta de Particiones
            - Carpeta del catálogo
            - Columnas de cada tipo de dato
            - Renglones de inicio
            - Obra en la que se trabaja
"""

#---------------------------   Datos ingresados para trabajar  ---------------------------
#---------------------------   M1.1 Carpetas de Origen   ---------------------------
directorio_catálogo = r'C:\Users\ADMIN\Desktop\Margaleff\1.0 Automatizaciones\1.02 Automatización búsqueda\Budge - Python\0 Test Documents\0.1 Test Catálogo\0.0 Catalogo Liverpool La Perla Claves Margaleff.xlsx'  
directorio_particiones = r'C:\Users\ADMIN\Desktop\Margaleff\1.0 Automatizaciones\1.02 Automatización búsqueda\Budge - Python\0 Test Documents\0.2 Test Particiones'
print(directorio_particiones)
directorio_destino = r'C:\Users\ADMIN\Desktop\Margaleff\1.0 Automatizaciones\1.02 Automatización búsqueda\Budge - Python\0 Test Documents\0.3 Test Reports'  # Destino del reporte
directorio_archivos_del_programa = r'C:\Users\ADMIN\Desktop\Margaleff\1.0 Automatizaciones\1.02 Automatización búsqueda\Budge - Python\Budge 0.1.1'





#---------------------------   M1.2.1 Ubicación de celdas de información CATÁLOGO   ---------------------------
# NOTA: Ingrese las columnas en MAYÚSCULAS
catalogo_columna_matricula = 'B'
catalogo_columna_nombre = 'C'
catalogo_columna_unidad = 'D'
catalogo_columna_cantidad = 'E'
catalogo_columna_pu = 'F'
catalogo_renglon_inicio = '6'  # Primer renglón en el que aparecen datos de forma contínua





#---------------------------   M1.2 Ubicación de celdas de información PARTICIONES   ---------------------------
particiones_columna_matricula = 'B'
particiones_columna_nombre = 'C'
particiones_columna_unidad = 'D'
particiones_columna_cantidad = 'E'
particiones_renglon_inicio = '17'  # Primer renglón en el que aparecen datos de forma contínua


class Documento:
    columna_matricula = ''
    columna_nombre = ''
    columna_unidad = ''
    columna_cantidad = ''
    columna_pu = ''
    renglon_inicio = ''
    
catalogo = Documento
catalogo.columna_matricula = catalogo_columna_matricula
catalogo.columna_nombre = catalogo_columna_nombre
catalogo.columna_unidad = catalogo_columna_unidad  
catalogo.columna_cantidad = catalogo_columna_cantidad
catalogo.columna_pu = catalogo_columna_pu 
catalogo.renglon_inicio = catalogo_renglon_inicio


particiones = Documento
particiones.columna_matricula = particiones_columna_matricula
particiones.columna_nombre = particiones_columna_nombre
particiones.columna_unidad = particiones_columna_unidad  
particiones.columna_cantidad = particiones_columna_cantidad
particiones.renglon_inicio = particiones_renglon_inicio
    



