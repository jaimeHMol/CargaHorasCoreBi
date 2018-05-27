# -*- coding: utf-8 -*-
"""
    functionCargaHorasCorebi.py
    Created on Sat Nov 18 14:32:09 2017

    Contiene los procesos para realiza la carga de horas automatizada para
    Core Bi

    @author: @kingSelta
"""

# IMPORTS GLOBALES
# ----------------
import time
from datetime import datetime, timedelta
import json
import requests
# Importa la libreria selenium que debió ser instalada previamente
# usando pip en el comand prompt
from selenium import webdriver


# CONSTANTES GLOBALES
# -------------------
PRIMERANHOSISTEMA = 2017
FECHAPRIMERANHOSISTEMA = datetime.strptime('20170101','%Y%m%d')
PRIMERDIAANHO2017 = 6209
DRIVERNAME = "chromedriver.exe"

def cargaHoras(workDirectory,
               urlLogin, urlImputacion, urlApiFeriados,
               usuario,
               contrasenha,
               cantHoras,
               fechaIni,
               fechaFin,
               detalle,
               concatenador,
               cliente,
               proy,
               etapa,
               fact):
    """ Realiza la imputación de horas en el sistema de Core Bi
    """

    # Funciones para hacer más prolija la navegación controlada desde python
    def navigate(url):
        """ Acceder a la url ingresada
        """
        chm_driver.get(url)

    def scripting(script):
        """ Ejecuta el script ingresado
        """
        chm_driver.execute_script(script)

    def scripting_return(script):
        """ Ejecuta el script ingresado y devuelve el elemento resultado
        """
        element = chm_driver.execute_script('return ' + str(script))
        return element

    def selecting_by_name(selector):
        """ Selecciona un elemento buscandolo por name
        """
        element = chm_driver.find_elements_by_name(selector)
        return element

    def selecting_by_classname(selector):
        """ Selecciona un elemento buscandolo por classname
        """
        element = chm_driver.find_element_by_class_name(selector)
        return element

    def selecting_by_id(selector):
        """ Selecciona un elemento buscandolo por id
        """
        element = chm_driver.find_elements_by_id(selector)
        return element

    def selecting_by_tag(selector):
        """ Selecciona un elemento buscandolo por el tag ingresado
        """
        element = chm_driver.find_elements_by_tag_name(selector)
        return element

    def clicking_by_name(selector):
        """ Hace click en un elemento buscandolo por name
        """
        element = chm_driver.find_elements_by_name(selector)
        element[0].click()

    def clicking_by_id(selector):
        """ Hace click en un elemento buscandolo por id
        """
        element = chm_driver.find_elements_by_id(selector)
        element[0].click()

    def page_source():
        """ Recupera la fuente de página
        """
        elem = chm_driver.page_source
        return elem

    def clicking_by_classname(selector):
        """ Hace click en un elemento buscandolo por classname
        """
        element = chm_driver.find_elements_by_class_name(selector)
        element[0].click()

    def clicking_by_other_tag(tagName, selector):
        """ Hace click en un elemento buscandolo por el tag ingresado
        """
        element = chm_driver.find_elements_by_xpath('//*[@'+ tagName
                                                    + '="' + selector + '"]')
        element[0].click()

    def close():
        """ Cierra el navegador
        """
        chm_driver.close()
        chm_driver.quit()

    def wait(wtime):
        """ Pausa el navegador, wtime en segundos
        """
        chm_driver.implicitly_wait(wtime)

    def cantDiasAnho(year):
        """ Detecta si es año bisiesto y devuelve la cantidad de dìas del año
        """
        if year % 400 == 0:
            return 366
        if year % 100 == 0:
            return 365
        if year % 4 == 0:
            return 366
        else:
            return 365

    def fillForm(inCliente, inCantHoras, inDetalle, inFact, inEtapa, inProy):
        """ Llena el formulario principal de la imputación de horas y carga
            la información ingresada en el sistema
        """

        # Seleccionar cliente
        scripting("document.getElementsByName('ctl00$ContentPlaceHolder1$ddlCliente')[0].value = " + inCliente)
        scripting("__doPostBack('ctl00$ContentPlaceHolder1$ddlCliente')")

        scripting("document.getElementById('ContentPlaceHolder1_txtHoras').value = '" + inCantHoras + "'")
        scripting("__doPostBack('ctl00$ContentPlaceHolder1$txtHoras')")

        scripting("document.getElementById('ContentPlaceHolder1_txtDetalle').value = '" + inDetalle + "'")

        if inFact != "True":
            clicking_by_id('ContentPlaceHolder1_cbFacturable')

        # Seleccionar etapa
        scripting("document.getElementsByName('ctl00$ContentPlaceHolder1$ddlEtapa')[0].value = " + inEtapa)

        # Espera que termine de cargar la pagina luego del evento de refresh que tiene
        time.sleep(5) # Segundos
        # Seleccionar proyecto
        scripting("document.getElementsByName('ctl00$ContentPlaceHolder1$ddlProyecto')[0].value = " + inProy)

        # Imputar la carga del horario
        clicking_by_name('ctl00$ContentPlaceHolder1$btnCargar')


    # MAIN
    # ----
    try:
        # -- Variables
        # ------------------------
        estado = False
        warnings = ""
        info = ""
        # Convierte a tipo datetime las fechas que me llegan en el formato esperado (YYYYMMDD)
        fechaIniDate = datetime(int(fechaIni[0:4]), int(fechaIni[4:6]), int(fechaIni[6:8]))
         # Convierte a tipo datetime las fechas que me llegan en el formato esperado (YYYYMMDD)
        fechaFinDate = datetime(int(fechaFin[0:4]), int(fechaFin[4:6]), int(fechaFin[6:8]))
        listConcat = concatenador.split(";")
        tamListConcat = len(listConcat)
        indListConcat = 0

        # MAIN
        # ----------------------------
        # Ingresar al sistema de carga
        # Requiere tener el controlador de chrome en la misma carpeta donde
        # está este programa y su respectivo .bat
        chm_driver = webdriver.Chrome(executable_path=workDirectory
                                      + DRIVERNAME)

        # Ingresar a la pagina
        chm_driver.get(urlLogin)

        # Ingresar nombre de usario y contraseña
        scripting("document.getElementById('txtUsuario').value = '"
                  + usuario + "'")
        scripting("document.getElementById('txtPassword').value = '"
                  + contrasenha + "'")

        # Acceder
        clicking_by_name('btnLogin')

        # Valida si el usuario y contraseña son correctos
        loginError = selecting_by_id("lblError")
        if len(loginError) != 0:
            # Cerrar navegador
            close()

            # Salir
            estado = False
            error = "Usuario o contraseña errados. Favor verifique"
            return estado, error

        fechaLoop = fechaIniDate
        cargaPrevia = None
        contDias = 0
        contFest = 0
        while  fechaLoop <= fechaFinDate:

            # Valida si es día entre semana para hacer la carga
            if (int(fechaLoop.strftime("%w")) != 0
                and int(fechaLoop.strftime("%w")) != 6):
                # Ingresar a la imputación de horas
                chm_driver.get(urlImputacion)
                festivo = False
                # Recupera la fecha del sistema (que la web carga por defecto
                # en el calendario) para posteriormente llevarla al primer día de ese mes
                # pues en ese mes es que aparece el calendario de la página principal
                primerDiaFecAct = datetime.today()

                # Cálcula el id del día inicial a cargar
                idDiaAct = PRIMERDIAANHO2017 + (fechaLoop
                           - FECHAPRIMERANHOSISTEMA).days + 1

                while True:
                    # Lleva la fecha actual al primer día del mes y la almacena
                    primerDiaFecAct = primerDiaFecAct.replace(day=1)

                    # Lleva la fecha actual al ultimo día del mes y la almacena
                    mesSigFecAct = (primerDiaFecAct.replace(day=28)
                                    + timedelta(days=4))
                    ultimoDiaFecAct = (mesSigFecAct
                                       - timedelta(days=mesSigFecAct.day))

                    # Cálcula el id del primer día del mes de la fecha actual
                    idPrimerDiaFecAct = PRIMERDIAANHO2017 + (primerDiaFecAct
                                        - FECHAPRIMERANHOSISTEMA).days + 1

                    # Cálcula el id del último día del mes de la fecha actual
                    idUltimoDiaFecAct = PRIMERDIAANHO2017 + (ultimoDiaFecAct
                                         - FECHAPRIMERANHOSISTEMA).days + 1

                    if idDiaAct < idPrimerDiaFecAct:
                        # Resta un día al primer día del mes actual para
                        # devolverse al mes anterior
                        primerDiaFecAct = primerDiaFecAct - timedelta(days=1)
                        # Dado que es un link (href) para evitar hacer el
                        # doPostBack que a veces no funciona
                        clicking_by_other_tag("title", "Go to the previous month")

                    if idDiaAct > idUltimoDiaFecAct:
                        # Suma un día al último día del mes actual para
                        # adelantarse al siguiente mes
                        primerDiaFecAct = ultimoDiaFecAct + timedelta(days=1)
                        # Dado que es un link (href) para evitar hacer el
                        # doPostBack que a veces no funciona
                        clicking_by_other_tag("title", "Go to the next month")

                    if idPrimerDiaFecAct <= idDiaAct <= idUltimoDiaFecAct:
                        # Ingresa al día a cargar, que se encuentra en el mes
                        # que actualmente muestra el calendario
                        scripting("__doPostBack('ctl00$ContentPlaceHolder1$calendario','"
                                  + str(idDiaAct) + "')")
                        break


                # Valida que no se haya hecho una carga de horas anteriormente para ese día
                cargaPrevia = selecting_by_id("ContentPlaceHolder1_gvHorasCargadas")
                if len(cargaPrevia) == 0:

                    # Verifica si el día a cargar es feriado
                    # --------------------------------------
                    apiFeriados = urlApiFeriados
                    #incluyeFeriadosOp = "?incluir=opcional"
                    urlFullApiFeriados = apiFeriados + fechaLoop.strftime("%Y")
                    allFeriados = requests.get(urlFullApiFeriados)

                    # Convierte el json en un arreglo
                    allFeriadosString = json.dumps(allFeriados.json())
                    allFeriadosDecoded = json.loads(allFeriadosString)
                    # Busca el el día y mes en los nodos recuperados
                    for item in allFeriadosDecoded:
                        if item['dia'] == int(fechaLoop.strftime("%d")) \
                            and item['mes'] == int(fechaLoop.strftime("%m")):
                            festivo = True
                            contFest = contFest + 1

                            if item['tipo'] != "inamovible":
                                info = "Precaución: Se cargaron días "
                                "'no laborables' como feriados."

                    if festivo:
                        fillForm("1", "08:00", "Feriado", "False", "6", "105")
                        warnings = ("Advertencia: Se cargó "
                                        + str(contFest)
                                        + " día(s) feriado(s).")

                    else:
                        # Realiza el proceso de concatenazacion del campo
                        # detalle con el campo concat
                        detalleConcat = detalle + listConcat[indListConcat]
                        if indListConcat < tamListConcat - 1:
                            indListConcat = indListConcat + 1

                        fillForm(cliente, cantHoras, detalleConcat,
                                 fact, etapa, proy)
                        contDias = contDias + 1


            fechaLoop = fechaLoop + timedelta(days=1)

        # Cerrar sesión
        clicking_by_id('lbtnCerrarSesion')

        # Cerrar navegador
        close()

        estado = True
        info = "Info: Se cargó " + str(contDias) + " día(s) laborable(s) "
        "normal(es). \n" + info

        return estado, info + "\n" + warnings
    except BaseException as error:
        estado = False
        return estado, format(error)
