# -*- coding: utf-8 -*-
"""
Created on Tue Jan 12 09:01:18 2021

@author: Andrés Cueto Estrada
@Programa: Módulo 2, Definición del objeto tipo Material
            Budge 0.1.1 M2

@Descripción: En este script se pueden observar la declaración del Objeto tipo Material 
    este tipo de objeto se definirá según su:
        - Material (nombre)
        - Matrícula
        - Precio Unitario
        - Cantidad presupuestada por el proveedor
        - Cantidad Contabilizada
    también conteiene algunos métodos que el mismo objeto requerirá, como la suma de los
    siguientes valores que se le ingresen.
"""

class Material:
    matricula = ''
    nombre = ''
    precio_unitario = ''
    cantidad_proveedor = ''
    cantidad_contabilizada = ''
