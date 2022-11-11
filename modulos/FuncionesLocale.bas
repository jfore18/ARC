Attribute VB_Name = "FuncionesLocale"
Option Explicit
'Funciones del API de Windows usadas para tareas de LOCALE
Private Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Boolean
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long

'Constantes para manejo de LOCALE

'Separador de fecha o formato de fecha metodo antiguo
Private Const LOCALE_SDATE = &H1F
'Formato de fecha metodo nuevo
Private Const LOCALE_SSHORTDATE = &H1F

'Formato de hora
Private Const LOCALE_STIMEFORMAT = &H1003

'Separador decimal
Private Const LOCALE_SDECIMAL = &HE

'Numero de digitos decimales situados despues del separador decimal.
Private Const LOCALE_IDIGITS = &H11

'separador de Miles
Private Const LOCALE_STHOUSAND As Long = &HF&

'Simbolo de moneda
Private Const LOCALE_SCURRENCY = &H14

'Caracter usado para separar listas de items
Private Const LOCALE_SLIST As Long = &HC

'separador de miles para monedas
Private Const LOCALE_SMONTHOUSANDSEP As Long = &H17

'Separador de decimales para monedas
Private Const LOCALE_SMONDECIMALSEP As Long = &H16

'Forma como se agrupan los digitos para monedas
Private Const LOCALE_SMONGROUPING As Long = &H18

'forma como se agrupan los digitos
Private Const LOCALE_SGROUPING As Long = &H10

'signo negativo
Private Const LOCALE_SNEGATIVESIGN As Long = &H51

'sistema de medidas 0 metrico, 1 ingles EU
Private Const LOCALE_IMEASURE As Long = &HD

'simbolo para AM
Private Const LOCALE_S1159 As Long = &H28
'simbolo para PM
Private Const LOCALE_S2359 As Long = &H29


'Mensaje que indica un cambio de configuracion
Private Const WM_SETTINGCHANGE = &H1A

'Mensaje broadcast para windows
Private Const HWND_BROADCAST = &HFFFF&

Public formatoFecha As String

Public formatoHora As String

Public separadorDecimal As String

Public digitosDecimales As String

Public simboloMoneda As String

Public separadorListas As String

Public separadorMiles As String

Public agrupamientoNumeros As String

Public simboloNegativo As String

Public sistemaMedida As String

Public simboloAM As String

Public simboloPM As String

Public Sub cargarValoresPorDefecto()
    'Se cargan los valores por defecto en el caso en el que no se hayan establecido los valores por medio de las variables publicas.
    'Si se quieren cambiar los valores de los formatos, se debe hacer desde afuera del modulo de la siguiente manera: FuncionesLocale.formatoFecha="yyyy/mm/dd"
    If formatoFecha = "" Then
        formatoFecha = "dd/MM/yyyy"
    End If
    
    If formatoHora = "" Then
        formatoHora = "hh:mm:ss tt"
    End If
    
    If separadorDecimal = "" Then
        separadorDecimal = "."
    End If
    
    If digitosDecimales = "" Then
        digitosDecimales = "2"
    End If
    
    If simboloMoneda = "" Then
        simboloMoneda = "$"
    End If

    If separadorListas = "" Then
        separadorListas = ","
    End If

    If separadorMiles = "" Then
        separadorMiles = ","
    End If
    
    If agrupamientoNumeros = "" Then
        agrupamientoNumeros = "3;0"
    End If
    
    If simboloNegativo = "" Then
        simboloNegativo = "-"
    End If

    If sistemaMedida = "" Then
        sistemaMedida = "0"
    End If
    
    If simboloAM = "" Then
        simboloAM = "a.m."
    End If
    
    If simboloPM = "" Then
        simboloPM = "p.m."
    End If

End Sub

Public Sub ModificaInformacionLocale()


    Dim dwLCID As Long
    
    cargarValoresPorDefecto
    
    'Obtenemos el Identificador del Locale por defecto del sistema
    dwLCID = GetSystemDefaultLCID()

    'Modificamos la configuracion del Locale para cada uno de las caracteristicas del Locale que nos interesan.
    If SetLocaleInfo(dwLCID, LOCALE_SDATE, formatoFecha) = False Then
        MsgBox "Error: No se pudo cambiar la configuración de formato de fecha!", , "Configuracion del sistema"
        Exit Sub
    End If

    If SetLocaleInfo(dwLCID, LOCALE_STIMEFORMAT, formatoHora) = False Then
        MsgBox "Error: No se pudo cambiar la configuración de formato de hora!", , "Configuracion del sistema"
        Exit Sub
    End If

    If SetLocaleInfo(dwLCID, LOCALE_SDECIMAL, separadorDecimal) = False Then
        MsgBox "Error: No se pudo cambiar la configuración de separador de decimales!", , "Configuracion del sistema"
        Exit Sub
    End If

    If SetLocaleInfo(dwLCID, LOCALE_IDIGITS, digitosDecimales) = False Then
        MsgBox "Error: No se pudo cambiar la configuración de dígitos decimales!", , "Configuracion del sistema"
        Exit Sub
    End If

    If SetLocaleInfo(dwLCID, LOCALE_SCURRENCY, simboloMoneda) = False Then
        MsgBox "Error: No se pudo cambiar la configuración de simbolo de moneda!", , "Configuracion del sistema"
        Exit Sub
    End If

    If SetLocaleInfo(dwLCID, LOCALE_SLIST, separadorListas) = False Then
        MsgBox "Error: No se pudo cambiar la configuración de separador de listas!", , "Configuracion del sistema"
        Exit Sub
    End If

    If SetLocaleInfo(dwLCID, LOCALE_STHOUSAND, separadorMiles) = False Then
        MsgBox "Error: No se pudo cambiar la configuración de separador de miles!", , "Configuracion del sistema"
        Exit Sub
    End If

    If SetLocaleInfo(dwLCID, LOCALE_SMONTHOUSANDSEP, separadorMiles) = False Then
        MsgBox "Error: No se pudo cambiar la configuración de separador de miles en monedas!", , "Configuracion del sistema"
        Exit Sub
    End If

    If SetLocaleInfo(dwLCID, LOCALE_SMONDECIMALSEP, separadorDecimal) = False Then
        MsgBox "Error: No se pudo cambiar la configuración de separador de decimales en monedas!", , "Configuracion del sistema"
        Exit Sub
    End If

    If SetLocaleInfo(dwLCID, LOCALE_SMONGROUPING, agrupamientoNumeros) = False Then
        MsgBox "Error: No se pudo cambiar la configuración de formato de agrupamiento de numeros para monedas!", , "Configuracion del sistema"
        Exit Sub
    End If

    If SetLocaleInfo(dwLCID, LOCALE_SGROUPING, agrupamientoNumeros) = False Then
        MsgBox "Error: No se pudo cambiar la configuración de formato de agrupamiento de numeros!", , "Configuracion del sistema"
        Exit Sub
    End If

    If SetLocaleInfo(dwLCID, LOCALE_SNEGATIVESIGN, simboloNegativo) = False Then
        MsgBox "Error: No se pudo cambiar la configuración de signo negativo!", , "Configuracion del sistema"
        Exit Sub
    End If

    If SetLocaleInfo(dwLCID, LOCALE_IMEASURE, sistemaMedida) = False Then
        MsgBox "Error: No se pudo cambiar la configuración de sistema de medida!", , "Configuracion del sistema"
        Exit Sub
    End If
    
    If SetLocaleInfo(dwLCID, LOCALE_S1159, simboloAM) = False Then
        MsgBox "Error: No se pudo cambiar la configuración de simbolo AM!", , "Configuracion del sistema"
        Exit Sub
    End If
    
    If SetLocaleInfo(dwLCID, LOCALE_S2359, simboloPM) = False Then
        MsgBox "Error: No se pudo cambiar la configuración de simbolo PM!", , "Configuracion del sistema"
        Exit Sub
    End If

    'Enviamos un mensaje a todas las ventanas del sistema, el cual indica que se ha cambiado una configuracion del sistema.
    PostMessage HWND_BROADCAST, WM_SETTINGCHANGE, 0, 0

    

End Sub
