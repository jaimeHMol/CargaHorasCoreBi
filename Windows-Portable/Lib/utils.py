# -*- coding: utf-8 -*-
"""
    utils.py
    Created on Mon Apr 30 20:48:00 2018
    Agrupa todas las funcionalidad o utilidades generales que la aplicación
    CargaHorasCoreBi requiere y que pueden ser adaptadas para cualquier otra
    aplicación similar

    @author: @kingSelta
"""

# IMPORTS GLOBALES
# ----------------
import base64
import xml.etree.ElementTree as ET
import os
from datetime import datetime

# VARIABLES GLOBALES
# ----------------
vowels = list("aeiouy")
consonants = list("bcdfghjklmnñpeqrstvuwxz")
punctuation = list(".,@")

def cargarConfigCombos(resourcesAppDir, nameConfigFile):
    """ Abre y lee el archivo de configuración con los combos de la interfaz
        para cargar cada combo en su respectivo diccionario de tipo
        [codigo:nombre] el combo de los proyectos son diccionarios anidados
        dentro de otro diccionario pues los posibles valores dependen del valor
        de cliente (key) elegido.
        Carga también los demás valores de configuración de la aplicación
        como las urls utilizadas en el proceso de carga
    """
    with open(resourcesAppDir + nameConfigFile, 'r') as xml_file:
        xmlConfig = ET.parse(xml_file)
    config = xmlConfig.getroot()

    urlLogin = config.find("urlLogin").text
    urlImputacion = config.find("urlImputacion").text
    urlApiFeriados = config.find("urlApiFeriados").text

    comboCliente = {}
    comboProyecto = {}
    for cliente in config.iter("cliente"):
        listaProyectos = {}
        comboCliente[cliente.find("codigo").text] = cliente.find("nombre").text

        for proyecto in cliente.iter("proyecto"):
            listaProyectos[proyecto.find("codigo").text] = proyecto.find("nombre").text
        comboProyecto[cliente.find("codigo").text] = listaProyectos

    comboEtapa = {}
    for etapa in config.iter("etapa"):
        comboEtapa[etapa.find("codigo").text] = etapa.find("nombre").text

    return urlLogin, urlImputacion, urlApiFeriados, \
            comboCliente, comboProyecto, comboEtapa


def creaArchivoParamIni(resourcesAppDir, nameParamIniFile):
    """ Verifica si existe el archivo xml de parámetros inciales y si no
        existe lo crea de acuerdo a la estructura requerida
    """
    if not os.path.isfile(resourcesAppDir + nameParamIniFile):
        # Estructura base del xml de parametros iniciales
        dataXmlConfig = ET.Element("data")
        user = ET.SubElement(dataXmlConfig, "user")
        passw = ET.SubElement(dataXmlConfig, "pass")
        cantHoras = ET.SubElement(dataXmlConfig, "cantHoras")
        dateIni = ET.SubElement(dataXmlConfig, "dateIni")
        dateFin = ET.SubElement(dataXmlConfig, "dateFin")
        coment = ET.SubElement(dataXmlConfig, "coment")
        concat = ET.SubElement(dataXmlConfig, "concat")
        cliente = ET.SubElement(dataXmlConfig, "cliente")
        proy = ET.SubElement(dataXmlConfig, "proyecto")
        etapa = ET.SubElement(dataXmlConfig, "etapa")
        fact = ET.SubElement(dataXmlConfig, "facturable")

        # Valores por defecto para los parámetros iniciales
        user.text = ""
        passw.text = ""
        cantHoras.text = "08:00"
        dateIni.text = ""
        dateFin.text = ""
        coment.text = ""
        concat.text = ""
        cliente.text = "" # "68" (Link)
        proy.text = "" # "478" (Sigma)
        etapa.text = "" # "8" (Desarrollo)
        fact.text = "True"

        dataStrConfig = ET.tostring(dataXmlConfig, encoding="unicode",
                                    method="html")
        try:
            xmlParamIni = open(resourcesAppDir + nameParamIniFile, "w")
        except IOError:
            print("El archivo no pudo ser abierto")
        else:
            # El archivo generado queda en formato UTF8
            xmlParamIni.write(dataStrConfig)
            xmlParamIni.close()


def cargarParamIni(resourcesAppDir, nameParamIniFile, cantIter):
    """ Carga los valores de los parámetros iniciales almacenados en el
        respectivo archivo xml
    """

    # Llama la función que crea el archivo xml de parametros inciales
    # en caso que este no exista
    creaArchivoParamIni(resourcesAppDir, nameParamIniFile)


    # Lee los parametros iniciales guardados en el xml
    try:
        with open(resourcesAppDir + nameParamIniFile, 'r') as xml_file:
            xmlParamIni = ET.parse(xml_file)

        dataIni = xmlParamIni.getroot()
        userIni = dataIni.findtext("user")
        conConsonants = sum(userIni.count(c) for c in consonants)
        contVowels = sum(userIni.count(c) for c in vowels)
        contPunct = sum(userIni.count(c) for c in punctuation)
        busca = (7 + conConsonants + contVowels) * len(userIni) + contPunct
        passwIni = buscarCodif("selta" + str(busca),
                               str(dataIni.findtext("pass")))
        cantHorasIni = dataIni.findtext("cantHoras")
        dateIniIni = dataIni.findtext("dateIni")
        if dateIniIni == "":
            dateIniIni = datetime.now().strftime("%Y%m%d")
        dateFinIni = dataIni.findtext("dateFin")
        if dateFinIni == "":
            dateFinIni = datetime.now().strftime("%Y%m%d")
        comentIni = dataIni.findtext("coment")
        concatIni = dataIni.findtext("concat")
        clienteIni = dataIni.findtext("cliente")
        proyIni = dataIni.findtext("proyecto")
        etapaIni = dataIni.findtext("etapa")
        factIni = dataIni.findtext("facturable")

    except:
        # En caso de algún error con el archivo de paráremtros iniciales (que
        # se encuentre corrupto o algo por el estilo) lo borra y lo vuelve a
        # crear de cero
        cantIter = cantIter + 1
        xml_file.close()

        # Elimina el archivo de parámetros iniciales
        os.remove(resourcesAppDir + nameParamIniFile)

        if cantIter <= 1:
            # Vuelve a crear y cargar los parámetros iniciales
            return cargarParamIni(resourcesAppDir, nameParamIniFile, cantIter)
        else:
            raise

    xml_file.close()
    return dataIni, userIni, passwIni, cantHorasIni, dateIniIni, dateFinIni, \
            comentIni, concatIni, clienteIni, proyIni, etapaIni, factIni


def guardaParam(panel, dataIni, codCliente, codProy,  codEtapa,
                     resourceAppDir, nameParamIniFile, dateFormat):
    """ Guarda la información ingresada en la interfaz para que en la próxima
        ejecución esta quede con los últimos valores usados
    """

    userFin = dataIni.find("user")
    userFin.text = panel.userTxt.GetValue()
    passwFin = dataIni.find("pass")
    conConsonants = sum(userFin.text.count(c) for c in consonants)
    contVowels = sum(userFin.text.count(c) for c in vowels)
    contPunct = sum(userFin.text.count(c) for c in punctuation)
    busca = (7 + conConsonants + contVowels) * len(userFin.text) + contPunct
    passwFin.text = alterarCodif("selta" + str(busca), panel.passwTxt.GetValue())
    cantHorasFin = dataIni.find("cantHoras")
    cantHorasFin.text = panel.cantHorasTxt.GetValue()
    dateIniFin = dataIni.find("dateIni")
    dateIniFin.text = panel.date_ini.GetValue().Format(dateFormat)
    dateFinFin = dataIni.find("dateFin")
    dateFinFin.text = panel.date_fin.GetValue().Format(dateFormat)
    comentFin = dataIni.find("coment")
    comentFin.text = panel.comentTxt.GetValue()
    concatFin = dataIni.find("concat")
    concatFin.text = panel.concatTxt.GetValue()
    clienteFin = dataIni.find("cliente")
    clienteFin.text = codCliente
    proyFin = dataIni.find("proyecto")
    proyFin.text = codProy
    etapaFin = dataIni.find("etapa")
    etapaFin.text = codEtapa
    factFin = dataIni.find("facturable")
    factFin.text = str(panel.factTxt.GetValue())

    # Configura y escribe el xml con los nuevos valores ingresados
    tree = ET.ElementTree(dataIni)
    tree.write(resourceAppDir + nameParamIniFile, encoding="unicode",
               method="html")


def validarCamposPanel(panel):
    """ Valida que todos los campos obligatorios del panel hayan sido
        ingresados
    """
    error = False
    mensaje = ""

    # Valida que los campos obligatorios hayan sido ingresados
    if (panel.userTxt.GetValue() == ""
        or panel.passwTxt.GetValue() == ""
        or panel.cantHorasTxt.GetValue() == ""
        or panel.date_ini.GetValue() == ""
        or panel.date_fin.GetValue() == ""
        or panel.comentTxt.GetValue() == ""
        or panel.clienteTxt.GetSelection() == -1
        or panel.proyTxt.GetSelection() == -1
        or panel.etapaTxt.GetSelection() == -1):
        error = True
        mensaje = "Todos los campos obligatorios deben ser ingresados"
    else:
        # Valida valores consistentes de fechas
        if (panel.date_fin.GetValue().Format("%Y%m%d")
            < panel.date_ini.GetValue().Format("%Y%m%d")):
            error = True
            mensaje = "La fecha inicial no puede ser mayor a la fecha final"

    return error, mensaje


def alterarCodif(string, clear):
    """ Modifica la cadena de caracteres según el string ingresado
    """
    enc = []
    for i in range(len(clear)):
        string_c = string[i % len(string)]
        enc_c = chr((ord(clear[i]) + ord(string_c)) % 256)
        enc.append(enc_c)
    return base64.urlsafe_b64encode("".join(enc).encode()).decode()


def buscarCodif(string, enc):
    """ Busca la cadena de caracteres adecuada según el string ingresado
    """
    dec = []
    enc = base64.urlsafe_b64decode(enc).decode()
    for i in range(len(enc)):
        string_c = string[i % len(string)]
        dec_c = chr((256 + ord(enc[i]) - ord(string_c)) % 256)
        dec.append(dec_c)
    return "".join(dec)
