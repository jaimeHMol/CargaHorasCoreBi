# -*- coding: utf-8 -*-
"""
    GUICargaHorasCorebi.py
    Created on Fri Apr 20 15:09:30 2018
    Contiene la interfaz gráfica con su lógica y validaciones, para capturar
    los datos necesarios para realizar la carga de horas automatizada para
    Core Bi

    @author: @kingSelta
"""
#!/usr/bin/python
# -*- coding: iso-8859-1 -*-


# IMPORTS GLOBALES
# ----------------
import os
import sys
import distutils.util
from datetime import datetime
from wx.lib.wordwrap import wordwrap
import wx as wx
import wx.adv


# CONSTANTES GLOBALES
# -------------------
ROOTAPPDIR = os.getcwd()
RESOURCESAPPDIR = ROOTAPPDIR + "\\Resources\\"
LIBAPPDIR = ROOTAPPDIR + "\\Lib\\"
NAMEAPP = "Carga Horas Core Bi"
NAMEPARAMINIFILE = "paramIni.xml"
NAMECONFIGFILE = "config.xml"
NAMEICON = "cargaHorasCorebiAjust.ico"
FRAMETITLE = "Carga horas Core Bi"
DATEFORMAT = "%Y%m%d"
## HINT: solo serviría para windows
sys.path.insert(2, LIBAPPDIR)

# Imports propios
import functionCargaHorasCorebi
import utils


class MyPanel(wx.Panel):
    """ Definición del panel como clase
    """
    def __init__(self, parent):
        """ Contructor del panel
        """
        super(MyPanel, self).__init__(parent)

        sizer = wx.GridBagSizer()

        # Campos de usuario y contraseña (crea label, input y carga valor inicial)
        self.userLbl = wx.StaticText(self, label="Usuario *")
        self.userTxt = wx.TextCtrl(self, -1, size=(170, -1))
        # Carga valor inicial
        self.userTxt.SetValue(userIni)
        # Carga ayuda o descripción del campo
        self.userTxt.SetToolTip(wx.ToolTip("Dirección de correo electrónico "
                                           "de Core Bi"))

        sizer.Add(self.userLbl, (1, 0), (1, 1), wx.ALIGN_CENTER_VERTICAL
                  | wx.LEFT, border=15)
        sizer.Add(self.userTxt, (2, 0), (1, 1), wx.EXPAND | wx.LEFT, border=15)

        self.passwLbl = wx.StaticText(self, label="Contraseña *")
        self.passwTxt = wx.TextCtrl(self, -1, size=(170, -1),
                                    style=wx.TE_PASSWORD)
        # Carga valor inicial
        self.passwTxt.SetValue(passwIni)
        # Carga ayuda o descripción del campo
        self.passwTxt.SetToolTip(wx.ToolTip("Contraseña asignada para ingresar "
                                            "al sistema de carga de horas (no "
                                            "necesariamente la misma del correo "
                                            "electrónico)"))

        sizer.Add(self.passwLbl, (4, 0), (1, 1), wx.ALIGN_CENTER_VERTICAL
                  | wx.LEFT, border=15)
        sizer.Add(self.passwTxt, (5, 0), (1, 1), wx.EXPAND | wx.LEFT,
                  border=15)
        #----------------------------------------------------------------------

        # Botones de calendario (crea label, input y carga valor inicial)
        self.date_iniLbl = wx.StaticText(self, label="Fecha inicial *")
        self.date_ini = wx.adv.DatePickerCtrl(self, style=wx.adv.DP_DROPDOWN)
        # Carga valor inicial
        datetimeIniIni = datetime.strptime(dateIniIni, '%Y%m%d')
        self.date_ini.SetValue(datetimeIniIni)
        # Carga ayuda o descripción del campo
        self.date_ini.SetToolTip(wx.ToolTip("Fecha incial del periodo de "
                                            "tiempo a cargar"))

        sizer.Add(self.date_iniLbl, (7, 0), (1, 1), wx.ALIGN_CENTER_VERTICAL
                  | wx.LEFT, border=15)
        sizer.Add(self.date_ini, (8, 0), (1, 1), wx.EXPAND  | wx.LEFT,
                  border=15)


        self.date_finLbl = wx.StaticText(self, label="Fecha final *")
        self.date_fin = wx.adv.DatePickerCtrl(self, style=wx.adv.DP_DROPDOWN)
        # Carga valor inicial
        datetimeFinIni = datetime.strptime(dateFinIni, '%Y%m%d')
        self.date_fin.SetValue(datetimeFinIni)
        # Carga ayuda o descripción del campo
        self.date_fin.SetToolTip(wx.ToolTip("Fecha final del periodo de "
                                            " tiempo a cargar"))

        sizer.Add(self.date_finLbl, (10, 0), (1, 1), wx.ALIGN_CENTER_VERTICAL
                  | wx.LEFT, border=15)
        sizer.Add(self.date_fin, (11, 0), (1, 1), wx.EXPAND | wx.LEFT,
                  border=15)
        #----------------------------------------------------------------------

        # CAMPOS FIJOS: cliente, proyecto, etapa y facturable
        # (crea label, input y carga valor inicial)
        # mirar http://zetcode.com/wxpython/layout/
        self.camposFijos = wx.StaticBox(self)
        fixedFieldBoxsizer = wx.StaticBoxSizer(self.camposFijos, wx.VERTICAL)
        self.buttonEneable = wx.ToggleButton(self, label="Activar campos fijos?")
        self.Bind(wx.EVT_TOGGLEBUTTON, self.btnFixedFieldClk, self.buttonEneable)
        fixedFieldBoxsizer.Add(self.buttonEneable, flag=wx.CENTER)
        fixedFieldBoxsizer.AddSpacer(12)

        self.clienteLbl = wx.StaticText(self, label="Cliente *")
        self.clienteTxt = wx.Choice(self, choices=list(comboCliente.values()))
        self.Bind(wx.EVT_CHOICE, self.selecCliente, self.clienteTxt)
        # Carga valor inicial
        if clienteIni != "":
            self.clienteTxt.SetSelection(list(comboCliente.keys()).index(clienteIni))
            self.clienteTxt.Disable()
            # Carga los valores de proyectos según el cliente elegido
            comboProyectoCliente = comboProyecto[clienteIni]
        else:
            comboProyectoCliente = {}
        # Carga ayuda o descripción del campo
        self.clienteTxt.SetToolTip(wx.ToolTip("Cliente al que se le cargarán  "
                                              "las horas de trabajo"))

        fixedFieldBoxsizer.Add(self.clienteLbl)
        fixedFieldBoxsizer.Add(self.clienteTxt)
        fixedFieldBoxsizer.AddSpacer(10)

        self.proyLbl = wx.StaticText(self, label="Proyecto *")
        self.proyTxt = wx.Choice(self, choices=list(comboProyectoCliente.values()))
        # Carga valor inicial
        if proyIni != "":
            self.proyTxt.SetSelection(list(comboProyectoCliente.keys()).index(proyIni))
            self.proyTxt.Disable()
        # Carga ayuda o descripción del campo
        self.proyTxt.SetToolTip(wx.ToolTip("Proyecto en el que se realizaron "
                                           "las tareas a cargar"))

        fixedFieldBoxsizer.Add(self.proyLbl)
        fixedFieldBoxsizer.Add(self.proyTxt)
        fixedFieldBoxsizer.AddSpacer(10)

        self.etapaLbl = wx.StaticText(self, label="Etapa *")
        self.etapaTxt = wx.Choice(self, choices=list(comboEtapa.values()))
        # Carga valor inicial
        if etapaIni != "":
            self.etapaTxt.SetSelection(list(comboEtapa.keys()).index(etapaIni))
            self.etapaTxt.Disable()
        # Carga ayuda o descripción del campo
        self.etapaTxt.SetToolTip(wx.ToolTip("Etapa o fase del proyecto en que "
                                            "se realizan las tareas a cargar"))

        fixedFieldBoxsizer.Add(self.etapaLbl)
        fixedFieldBoxsizer.Add(self.etapaTxt)
        fixedFieldBoxsizer.AddSpacer(10)

        self.cantHorasLbl = wx.StaticText(self, label="Tiempo empleado *")
        self.cantHorasTxt = wx.TextCtrl(self)
         # Carga valor inicial
        self.cantHorasTxt.SetValue(cantHorasIni)
        self.cantHorasTxt.Disable()
        # Carga ayuda o descripción del campo
        self.cantHorasTxt.SetToolTip(wx.ToolTip("Cantidad de horas a cargar "
                                                "cada día. Debe tener formato "
                                                "HH:MM"))

        fixedFieldBoxsizer.Add(self.cantHorasLbl)
        fixedFieldBoxsizer.Add(self.cantHorasTxt)
        fixedFieldBoxsizer.AddSpacer(10)

        self.factTxt = wx.CheckBox(self, label="Facturable *")
        # Carga valor inicial
        self.factTxt.SetValue(bool(distutils.util.strtobool(factIni)))
        self.factTxt.Disable()
        # Carga ayuda o descripción del campo
        self.factTxt.SetToolTip(wx.ToolTip("Seleccionarlo en caso que las "
                                           "tareas a cargar sean facturables"))

        if (clienteIni == "" or proyIni == "" or etapaIni == "" or
            cantHorasIni == "" or factIni == ""):
            self.buttonEneable.SetValue(True)

            # Llama el evento de click sobre el botón como si se hubiera presionado
            evt = wx.PyCommandEvent(wx.EVT_TOGGLEBUTTON.typeId,
                                    self.buttonEneable.GetId())
            wx.PostEvent(self, evt)

        fixedFieldBoxsizer.Add(self.factTxt)

        sizer.Add(fixedFieldBoxsizer, (1, 2), (12, 1), flag=wx.EXPAND
                  | wx.RIGHT, border=15)
        #----------------------------------------------------------------------

        # Campo comentario (crea label, input y carga valor inicial)
        self.comentLbl = wx.StaticText(self, label="Detalle * (de las tareas "
                                       "a cargar)")
        self.comentTxt = wx.TextCtrl(self, style=wx.TE_MULTILINE)
        # Carga valor inicial
        self.comentTxt.SetValue(comentIni)
        # Carga ayuda o descripción del campo
        self.comentTxt.SetToolTip(wx.ToolTip("Explicación de las tareas a "
                                             "cargar"))

        sizer.Add(self.comentLbl, (13, 0), (1, 3), wx.LEFT | wx.RIGHT,
                  border=15)
        sizer.Add(self.comentTxt, (14, 0), (4, 3), wx.EXPAND | wx.LEFT
                  | wx.RIGHT, border=15)
        sizer.AddGrowableRow(14)

        # Campo concatenador (crea label, input y carga valor inicial)
        self.concatLbl = wx.StaticText(self, label="Concatenador (Opcional) "
                                       "Para generar un detalle diferente por "
                                       "cada concatenador,\n separado por ';'")
        self.concatTxt = wx.TextCtrl(self, style=wx.TE_MULTILINE)
        # Carga valor inicial
        self.concatTxt.SetValue(concatIni)
        # Carga ayuda o descripción del campo
        self.concatTxt.SetToolTip(wx.ToolTip("Para generar un detalle "
                                             "diferente por cada elemento que "
                                             "se ingrese en este campo, "
                                             "separandolos con el caracter "
                                             "';'"))

        sizer.Add(self.concatLbl, (19, 0), (2, 3), wx.EXPAND | wx.LEFT
                  | wx.RIGHT, border=15)
        sizer.Add(self.concatTxt, (21, 0), (3, 3), wx.EXPAND | wx.LEFT
                  | wx.RIGHT, border=15)
        #----------------------------------------------------------------------

        # Botones
        self.buttonExit = wx.Button(self, label="Salir")
        sizer.Add(self.buttonExit, (25, 0), (1, 1), wx.EXPAND | wx.LEFT
                  | wx.BOTTOM, border=15)
        self.buttonExit.Bind(wx.EVT_BUTTON, self.btnExitClk, self.buttonExit)

        self.buttonHelp = wx.Button(self, label="Acerca de")
        sizer.Add(self.buttonHelp, (25, 1), (1, 1), wx.EXPAND | wx.BOTTOM,
                  border=15)
        self.buttonHelp.Bind(wx.EVT_BUTTON, self.btnHelptClk, self.buttonHelp)

        self.buttonLoad = wx.Button(self, label="Cargar")
        sizer.Add(self.buttonLoad, (25, 2), (1, 1), wx.EXPAND | wx.RIGHT
                  | wx.BOTTOM, border=15)
        self.Bind(wx.EVT_BUTTON, self.btnLoadClk, self.buttonLoad)
        #----------------------------------------------------------------------

        # Configuracion final
        sizer.AddGrowableCol(1)

        # Asigna el sizer construido al panel que contiene todos los widgets
        self.SetSizerAndFit(sizer)
        #----------------------------------------------------------------------

    # Eventos
    def btnLoadClk(self, event):
        """ Realiza el llamado de la carga de horas de acuerdo a los valores
            seleccionados en la interfaz, luego de haber presionado el botón
            de carga
        """

        # Valida que todos los campos requeridos hayan sido ingresados
        error, mensaje = utils.validarCamposPanel(self)
        if error:
            wx.MessageBox(mensaje, NAMEAPP, wx.OK | wx.ICON_ERROR)
            return

        wait = wx.BusyInfo("Ejecutando el proceso, espere por favor, "
                           "puede tomar varios minutos")

        # Busca los codigos correspondientes según lo seleccionado en los
        # combos de cliente, proyecto y etapa
        codCliente = list(comboCliente.keys())[self.clienteTxt.GetSelection()]
        comboProyectoCliente = comboProyecto[codCliente]
        codProy = list(comboProyectoCliente.keys())[self.proyTxt.GetSelection()]
        codEtapa = list(comboEtapa.keys())[self.etapaTxt.GetSelection()]

        # Guarda (en archivo de parametros iniciales) los valores ingresados
        # en el formulario para poderlos cargar en posteriores ejecuciones
        utils.guardaParam(self, dataIni, codCliente, codProy, codEtapa,
                          RESOURCESAPPDIR, NAMEPARAMINIFILE, DATEFORMAT)

        # Llama al proceso principal de la carga de horas
        resultado, detalle = functionCargaHorasCorebi.cargaHoras(
            RESOURCESAPPDIR,
            urlLogin, urlImputacion, urlApiFeriados,
            self.userTxt.GetValue(),
            self.passwTxt.GetValue(),
            self.cantHorasTxt.GetValue(),
            self.date_ini.GetValue().Format(DATEFORMAT),
            self.date_fin.GetValue().Format(DATEFORMAT),
            self.comentTxt.GetValue(),
            self.concatTxt.GetValue(),
            codCliente,
            codProy,
            codEtapa,
            str(self.factTxt.GetValue()))

        del wait
        if resultado:
            wx.MessageBox("El proceso finalizó correctamente.\n\n" +
                          detalle, NAMEAPP)
        else:
            wx.MessageBox("El proceso encontró un error\n\n Traza del error: "
                          + detalle, NAMEAPP, wx.OK | wx.ICON_ERROR)

    def btnFixedFieldClk(self, event):
        """ Habilita o deshabilita los campos fijos, según la selección
            del usuario
        """
        if self.buttonEneable.GetValue():
            self.clienteTxt.Enable()
            self.proyTxt.Enable()
            self.etapaTxt.Enable()
            self.cantHorasTxt.Enable()
            self.factTxt.Enable()
        else:
            self.clienteTxt.Disable()
            self.proyTxt.Disable()
            self.etapaTxt.Disable()
            self.cantHorasTxt.Disable()
            self.factTxt.Disable()

    def selecCliente(self, event):
        """ Carga los posibles valores de proyectos según el cliente
            seleccionado
        """
        clienteSel = list(comboProyecto.keys())[self.clienteTxt.GetSelection()]
        comboProyectoCliente = comboProyecto[clienteSel]

        self.proyTxt.Clear()
        self.proyTxt.AppendItems(list(comboProyectoCliente.values()))

    def btnExitClk(self, event):
        """ Recupera y cierra el frame cuando se presiona el botón de salida
        """
        frame = self.GetParent()
        frame.Close()

    def btnHelptClk(self, event):
        """ Muestra la ventana "Acerca de" con la información del programa
        """
        # Llena la información del objeto de la venta de about
        info = wx.adv.AboutDialogInfo()
        info.Name = "Carga horas Core Bi"
        info.Version = "1.0.0"
        info.Copyright = "2018 Selta Soft"
        info.Description = wordwrap("Programa que automatiza la carga de "
                                    "horas trabajadas en el sistema de horas "
                                    "que la empresa tiene para tal fin, "
                                    "buscando hacer más amable y menos "
                                    "aburrida la vida de los empleados :)\n\n"
                                    "Características:\n"
                                    "* Automatiza todo el proceso de carga "
                                    "de horas en un solo formulario. Permitiendo "
                                    "cargar al mismo tiempo todos los días "
                                    "hábilies de un rango de tiempo.\n"
                                    "* Escrito en su totalidad en Python.\n"
                                    "* Utiliza Selenium para automatizar Chrome.\n"
                                    "* Almacena la última configuración para "
                                    "hacer aún más ágil la carga.\n"
                                    "* Valida y evita cargar días no hábiles "
                                    " (fines de semana).\n"
                                    "* Realiza la validación y carga de días "
                                    "feriados de Argentina, usando la API: "
                                    "https://pjnovas.gitbooks.io/no-laborables/\n\n",
                                    350, wx.ClientDC(self))
        info.WebSite = ("https://about.me/kingselta")
        info.Developers = ["Jaime A Herran M  @kingSelta"]
        info.License = wordwrap("Esta aplicación es de código abierto "
                                "hecha de manera libre, autodidacta, por "
                                "diversión y buscando evitar la fatiga XD \n"
                                "Sientete libre de usarla, reutilizarla o "
                                "mejorarla, en cuyo caso me gustaría "
                                "enterarme y ser mencionado, sin ser esto "
                                "imprescindible.\n\n Todos los nombres, logos, "
                                "y marcas pertenecen a sus respectivos dueños."
                                , 500, wx.ClientDC(self))

        # Llama al wx.AboutBox con la información previamente ingresada en el
        # correspondiente objeto
        wx.adv.AboutBox(info)
        #----------------------------------------------------------------------

class simpleapp_wx(wx.Frame):
    """ Definición de la aplicación como clase
    """
    def __init__(self, parent):
        """ Constructor del frame
        """
        super(simpleapp_wx, self).__init__(parent)
        #self.parent = parent
        self.initialize()

        # Define el icono para el frame
        icono = wx.Icon()
        icono.CopyFromBitmap(wx.Bitmap(RESOURCESAPPDIR + NAMEICON,
                                       wx.BITMAP_TYPE_ANY))
        self.SetIcon(icono)


    def initialize(self):
        """ Configuración inicial del frame y su panel
        """
        panel = MyPanel(self)

        # Crea y ajusta un sizer para el Frame principal
        sizer2 = wx.GridBagSizer()
        sizer2.AddGrowableCol(0)
        sizer2.AddGrowableRow(0)
        sizer2.Add(panel, (0, 0), (1, 1), wx.EXPAND)
        self.SetSizerAndFit(sizer2)

        # Muestra el Frame con todo configurado y listo
        self.SetTitle(FRAMETITLE)
        self.Centre()
        self.Show(True)


if __name__ == "__main__":

    # Verifica/crea y carga los parámetros inciales a partir del respectivo
    # archivo xml
    dataIni, userIni, passwIni, cantHorasIni, dateIniIni, dateFinIni, \
    comentIni, concatIni, clienteIni, proyIni, etapaIni, factIni \
                   = utils.cargarParamIni(RESOURCESAPPDIR, NAMEPARAMINIFILE, 0)


    # Carga los diccionarios con los que se llenan los combos de la GUI según el
    # archivo de configuración
    urlLogin, urlImputacion, urlApiFeriados, comboCliente, comboProyecto, \
         comboEtapa = utils.cargarConfigCombos(RESOURCESAPPDIR, NAMECONFIGFILE)

    # Crea y genera la interfaz gráfica
    app = wx.App()
    simpleapp_wx(None)
    app.MainLoop()
