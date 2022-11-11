Attribute VB_Name = "funcionesARC"
Sub descargarFormularios()
    On Error Resume Next
    Unload frm_clave_usuarios
    Unload frm_Consulta_general_ARC
    Unload frmConsultaCarga
    Unload frmConsultaBitacora
    Unload frmConsultaResponsabilidades
End Sub

Function cerrarObjeto(ByVal objeto As Object)
    On Error Resume Next
    If Not objeto Is Nothing Then
        If objeto.State = adStateOpen Then objeto.Close
    End If
    Set objeto = Nothing
End Function

Public Sub bitacora(clase As Integer, Evento As Integer, accion As Integer, fechaContable As String, aplicacionFuente As String, transaccion As String, detalle As String, connectionTmp As ADODB.Connection)
    Dim sentenciaBitacora As String
    Dim fecha As String
    Dim fechaHora As String
    'Dim recordSetObjBitacora as
    'Se extrae la hora del sistema
    Set recordSetObjBitacora = connectionTmp.Execute("select to_char(SYSDATE,'yyyy/mm/dd')FECHA, to_char(SYSDATE,'yyyy/mm/dd HH:MI:SS AM') HORA from DUAL")
    fecha = "TO_DATE('" & recordSetObjBitacora("FECHA") & "', 'yyyy/mm/dd')"
    fechaHora = "TO_DATE('" & recordSetObjBitacora("HORA") & "', 'yyyy/mm/dd HH:MI:SS AM')"
    cerrarObjeto recordSetObjBitacora

    'Se establacen los valores

    If Trim(fechaContable) = "" Then
        fechaContable = "NULL"
    Else
        fechaContable = "TO_DATE('" & fechaContable & "', 'yyyy/mm/dd')"
    End If

    If Trim(detalle) = "" Then
        detalle = "NULL"
    Else
        detalle = "'" & Trim(Replace(detalle, "'", "")) & "'"
    End If

    If Trim(aplicacionFuente) = "" Then
        aplicacionFuente = "NULL"
    Else
        aplicacionFuente = "'" & Trim(Replace(aplicacionFuente, "'", "")) & "'"
    End If


    If Trim(transaccion) = "" Then
        transaccion = "NULL"
    Else
        transaccion = "'" & Trim(Replace(transaccion, "'", "")) & "'"
    End If
    If Stg_cod_Usuario = "" Then
        Stg_cod_Usuario = "22222221"
    End If

    'Se inserta el registro en la bitacora
    sentenciaBitacora = "insert into TBL_BITACORA (COD_ENTIDAD, CLASE, EVENTO, ACCION, FECHA_SISTEMA, HORA, FECHA_CONTABLE, APLICACION_FUENTE, TRANSACCION, USUARIO, DETALLE) " & _
                        "values (1," & clase & "," & Evento & "," & accion & "," & fecha & "," & fechaHora & "," & fechaContable & "," & aplicacionFuente & "," & transaccion & ",'" & Stg_cod_Usuario & "'," & detalle & ")"

    connectionTmp.Execute (sentenciaBitacora)
End Sub


'Carga un recorset desconectado
Function cargarRecordSet(sentencia As String) As ADODB.Recordset
    On Error GoTo error
    'Defines the recordSet to be returned
    Set rsGrilla = New ADODB.Recordset

    'Sets the Oracle connection
    PL_Conexion_Oracle

    'Indicates the location of the cursor service
    rsGrilla.CursorLocation = adUseClient

    'Opens the recordSet with a connection based sentence

    'MsgBox sentencia
    rsGrilla.Open sentencia, stringConexion, adOpenStatic, adLockReadOnly

    'Disconnects the recordset from the oracle Connection
    Set rsGrilla.ActiveConnection = Nothing

    'Close the active connection
    If Not cnObj1 Is Nothing Then
        If cnObj1.State = adStateOpen Then cnObj1.Close
    End If
    Set cnObj1 = Nothing

    'Returns the recordSet
    Set cargarRecordSet = rsGrilla



    Exit Function

error:

    MsgBox "No fue posible llenar el recordSet. " & Chr$(13) & Chr$(10) & Err.Number & ".  " & Err.Description

    'Closes the inner recordSet
    If Not rsGrilla Is Nothing Then
        If rsGrilla.State = adStateOpen Then rsGrilla.Close
    End If
    Set rsGrilla = Nothing

    'Closes the active connection
    If Not cnObj1 Is Nothing Then
        If cnObj1.State = adStateOpen Then cnObj1.Close
    End If
    Set cnObj1 = Nothing
End Function

Public Function rellenar(cadena As String, posiciones As Integer, caracter As String, lado As String) As String
    Dim I As Integer
    Dim COMPLEMENTO As String
    COMPLEMENTO = ""
    If Len(cadena) >= posiciones Then
        rellenar = cadena
    Else
        For I = Len(cadena) + 1 To posiciones
            COMPLEMENTO = COMPLEMENTO + caracter
        Next I
        If lado = "derecha" Then
            rellenar = cadena + COMPLEMENTO
        End If

        If lado = "izquierda" Then
            rellenar = COMPLEMENTO + cadena
        End If
    End If
End Function

Function verificaTipoMovimientoAsignado() As Boolean
    PL_Conexion_Oracle
    sentencia = "SELECT COUNT(*) CUENTA FROM TBL_USUARIO_MOVIMIENTO WHERE COD_USR = " & Stg_cod_Usuario
    Set rsobj = cnObj1.Execute(sentencia)
    If rsobj("CUENTA") > 0 Then
        verificaTipoMovimientoAsignado = True
    Else
        verificaTipoMovimientoAsignado = False
    End If
    rsobj.Close
    cnObj1.Close
    Set rsobj = Nothing
    Set cnObj1 = Nothing
    Exit Function
End Function
Function verificaTransacciones() As Boolean
    PL_Conexion_Oracle
    sentencia = "SELECT COUNT(*) CUENTA FROM TBL_USUARIO_TRANSACCION WHERE COD_USR = " & Stg_cod_Usuario
    Set rsobj = cnObj1.Execute(sentencia)
    If rsobj("CUENTA") > 0 Then
        verificaTransacciones = True
    Else
        verificaTransacciones = False
    End If
    rsobj.Close
    cnObj1.Close
    Set rsobj = Nothing
    Set cnObj1 = Nothing
    Exit Function
End Function
'Funcion que consulta la ultima fecha depurada en ARC para
'determinar si se consulta la base en linea o la historica
Function FG_Retorna_base_datos_consultar(fecha As Date) As Boolean
    PL_Conexion_Oracle
    sentencia = "SELECT TO_DATE(PARAMETRO_ALFANUMERICO,'DD/MM/YYYY') PARAMETRO_ALFANUMERICO FROM TBL_PARAMETRO WHERE COD_PARAMETRO = 11"
    Set rsobj = cnObj1.Execute(sentencia)

    If rsobj("PARAMETRO_ALFANUMERICO") < fecha Then
        'Fecha puede ser consultada en linea
        FG_Retorna_base_datos_consultar = True
    Else
        'Fecha puede ser consultada en tablas historicas
        FG_Retorna_base_datos_consultar = False
    End If
    rsobj.Close
    cnObj1.Close
    Set rsobj = Nothing
    Set cnObj1 = Nothing
    Exit Function
End Function


Public Sub fg_valida_tamaño_fecha()
'Esta función se realiza con el fin de validar que en el panel de control la fecha este bien configurada
'para este equipo

    Dim fecha As String

    fecha = Date
    Dtg_Fecha_Equipo = Format(Date, "DD/MM/YYYY")

    If Len(Dtg_Fecha_Equipo) <> Len("DD/MM/YYYY") Then
        MsgBox "Existen problemas con la configuración de la fecha en el panel de control. Por favor verifique que sea 'dd/mm/yyyy')"
        End

    End If

End Sub

Function FG_Digito_Chequeo_Auxiliar(Numero_Cuenta_Auxiliar As String) As String

'---------------------------------------------------------------------------
' Nombre:   FG_Digito_Chequeo_Auxiliar
'
' Propósito: Calcula el Digito de chequeo para las cuentas contables
'
' Descripción:  Multiplica cada uno de los digitos de la cuenta por su peso
'               equivalente al resultado de esta multiplicacion le optiene el
'               modulo 10, se lo suma a la divicion enterea por 10 del mismo
'               resultado y lo acumula en CALCULO por ultimo se le resta a
'               10 el modulo 10 de CALCULO y a este resultado se le calcula
'               modulo de 10.
'
' Parámetros: El parametro de entrada es Numero_cuenta de tipo String.
'
' Resultado:  La función retorna el digito de chequeo
'
' Autor: Alvaro Hurtado'
' Fecha: 95/11/24'
'---------------------------------------------------------------------------
    On Error Resume Next

    Ponderador = "1212121212"
    Calculo = 0

    If Len(Numero_Cuenta_Auxiliar) <> 10 Then
        FG_Digito_Chequeo_Auxiliar = "N"
        Exit Function
    End If

    For I = 1 To Len(Numero_Cuenta_Auxiliar)
        numero = (Val(Mid$(Numero_Cuenta_Auxiliar, I, 1)) * Val(Mid$(Ponderador, I, 1)))
        Calculo = Calculo + (numero Mod 10) + (numero \ 10)
    Next I

    Calculo = (10 - (Calculo Mod 10)) Mod 10

    FG_Digito_Chequeo_Auxiliar = Format$(Str(Calculo), "0")

End Function



Function fg_Encripta(cadena As String) As String

    Dim Longitud As Integer
    Dim contador As Integer
    Dim VAR_TEXTO As Integer
    Dim VALOR_NUMERICO As Integer
    Dim cadena_cifrada As String

    Longitud = Len(cadena)
    contador = 1
    fg_Encripta = ""
    While contador <= Longitud
        VALOR_NUMERICO = Asc(Mid(cadena, contador, 1))
        VAR_TEXTO = contador Mod 2
        If VAR_TEXTO = 0 Then
            VAR_TEXTO = 1
        Else
            VAR_TEXTO = 3

        End If
        VALOR_NUMERICO = VALOR_NUMERICO + VAR_TEXTO
        cadena_cifrada = cadena_cifrada & Chr(VALOR_NUMERICO)
        contador = contador + 1
    Wend
    fg_Encripta = cadena_cifrada

End Function
Function fg_Desencripta(cadena As String) As String

    Dim Longitud As Integer
    Dim contador As Integer
    Dim VAR_TEXTO As Integer
    Dim VALOR_NUMERICO As Integer
    Dim cadena_cifrada As String

    Longitud = Len(cadena)
    contador = 1
    fg_Desencripta = ""
    While contador <= Longitud
        VALOR_NUMERICO = Asc(Mid(cadena, contador, 1))
        VAR_TEXTO = contador Mod 2
        If VAR_TEXTO = 0 Then
            VAR_TEXTO = 1
        Else
            VAR_TEXTO = 3
        End If
        VALOR_NUMERICO = VALOR_NUMERICO - VAR_TEXTO
        cadena_cifrada = cadena_cifrada & Chr(VALOR_NUMERICO)
        contador = contador + 1
    Wend
    fg_Desencripta = cadena_cifrada

End Function

Function fg_Crea_archivo_Ini(servidor As String, Usuario As String, Clave As String) As String
    Dim FNAME As String
    Dim FNUM As Integer
    Dim Linea As String



    'Detecta el directorio donde esta corriendo la aplicacion
    STG_PATH_APL = STG_DIRECTORIO_SISTEMA & "\ARCN_NVO.INI"

    'Abre el archivo de configuracion ARCN_NVO.INI
    FNUM = FreeFile      ' Determine file number.
    FNAME = STG_PATH_APL
    On Error Resume Next
    Open FNAME For Output As FNUM

    If Err.Number = 0 Then
        On Error GoTo 0
        Print #FNUM, "[SERVIDOR_HOST  ]" & servidor
        Print #FNUM, "[DRIVER_HOST    ]OraOLEDB.Oracle"
        Print #FNUM, "[USUARIO_HOST   ]" & Usuario
        Print #FNUM, "[CLAVE_HOST     ]" & Clave
        Print #FNUM, "[ENTRADAS_LOCAL ]D:\APL\ARC\ENTRADAS\"
        Print #FNUM, "[SALIDAS_LOCAL  ]D:\APL\ARC\SALIDAS\"
        Print #FNUM, "[REPORTES_LOCAL ]D:\APL\ARC\REPORTES\"
        Print #FNUM, "[TIMEOUT        ]30"
        Close
    Else
        fg_Crea_archivo_Ini = Err.Number & " : " & Err.Description & Chr(13) & "Existen Problemas al escribir el archivo ARCN_NVO.INI"
        MsgBox STG_RESPUESTA
        Exit Function

    End If

    fg_Crea_archivo_Ini = "OK"

    Exit Function


End Function






Function FG_vlr_numerico(Digito_Entrada As String) As Integer

' PROPOSITO:    Valida que si un campo es de tipo
'               numérico reciba solo números, para letras
'               o símbolos, devuelve el código Ascii con
'               valor cero para que el campo no muestre
'               ningún caracter.

    On Error GoTo error

    If IsNumeric(Digito_Entrada) Or Asc(Digito_Entrada) = vbKeyBack Then
        FG_vlr_numerico = Asc(Digito_Entrada)
    Else
        FG_vlr_numerico = 0
    End If

    Exit Function
error:
    If Err.Number = 91 Then
        MsgBox "Existen Problemas de comunicación con la Base de datos , contacte al administrador de la red.  Ext. 3291", vbCritical
    Else
        MsgBox Err.Number & " " & Err.Description
    End If
End Function


Function FG_COMPARA_CENTRO() As Integer
    If Ing_Centro_Usuario <> Ing_Centro_Usuario_original Then
        FG_COMPARA_CENTRO = 1
    Else
        FG_COMPARA_CENTRO = 0
    End If
End Function



Function FG_Buscar_Mensajes(Codigo_Mensaje As Integer) As String

' PROPOSITO:  Valida si el número del mensaje existe en la tabla de
'             mensajes, en caso contrario devuelve Mensaje no Definido
    On Error GoTo error

    Set obj_acceso_tabla = CreateObject("clsacceso_tabla.cls_acceso_tabla")
    Set rsobj = obj_acceso_tabla.Fg_acceso_datos_recorset("TBL_MENSAJE", 1, "COD_MENSAJE;", Codigo_Mensaje & ";", "=;", "*", "COD_MENSAJE ASC")
    If Not rsobj.EOF Then
        FG_Buscar_Mensajes = rsobj("DES_MENSAJE")
    Else
        FG_Buscar_Mensajes = "Mensaje no definido"
    End If

    Exit Function
error:
    If Err.Number = 91 Then
        MsgBox "Existen Problemas de comunicación con la Base de datos , contacte al administrador de la red.  Ext. 3291", vbCritical
    Else
        MsgBox Err.Number & " " & Err.Description
    End If

End Function

Function FG_vlr_numerico_decimal(Digito_Entrada As String) As Integer

' PROPOSITO:    Valida que si un campo es de tipo
'               numérico reciba solo números y punto decimal, para letras
'               o símbolos, devuelve el código Ascii con
'               valor cero para que el campo no muestre
'               ningún caracter.

    On Error GoTo error

    If IsNumeric(Digito_Entrada) Or Digito_Entrada = "." Or Asc(Digito_Entrada) = vbKeyBack Then
        FG_vlr_numerico_decimal = Asc(Digito_Entrada)
    Else
        FG_vlr_numerico_decimal = 0
    End If

    Exit Function
error:
    If Err.Number = 91 Then
        MsgBox "Existen Problemas de comunicación con la Base de datos , contacte al administrador de la red.  Ext. 3291", vbCritical
    Else
        MsgBox Err.Number & " " & Err.Description
    End If

End Function

Function FG_vlr_numerico_fecha(Digito_Entrada As String) As Integer

' PROPOSITO:    Valida que si un campo es de tipo
'               numérico reciba solo números y punto decimal, para letras
'               o símbolos, devuelve el código Ascii con
'               valor cero para que el campo no muestre
'               ningún caracter.

    On Error GoTo error

    If IsNumeric(Digito_Entrada) Or Digito_Entrada = "/" Or Asc(Digito_Entrada) = vbKeyBack Then
        FG_vlr_numerico_fecha = Asc(Digito_Entrada)
    Else
        FG_vlr_numerico_fecha = 0
    End If

    Exit Function
error:
    If Err.Number = 91 Then
        MsgBox "Existen Problemas de comunicación con la Base de datos , contacte al administrador de la red.  Ext. 3291", vbCritical
    Else
        MsgBox Err.Number & " " & Err.Description
    End If

End Function
Function FG_Digito_chequeo(Numero_Cuenta As Variant) As String

'---------------------------------------------------------------------------
' Nombre:   FG_Digito_chequeo
'
' Propósito: Calcula el Digito de chequeo para las cuentas contables
'
' Descripción:  Multiplica cada uno de los digitos de la cuenta por su peso
'               equivalente al resultado de esta multiplicacion le optiene el
'               modulo 10, se lo suma a la divicion enterea por 10 del mismo
'               resultado y lo acumula en CALCULO por ultimo se le resta a
'               10 el modulo 10 de CALCULO y a este resultado se le calcula
'               modulo de 10.
'
' Parámetros: El parametro de entrada es Numero_cuenta de tipo String.
'
' Resultado:  La función retorna el digito de chequeo
'
' Autor: Banco de Bogotà
' Fecha: 97/06/04'
'---------------------------------------------------------------------------
    On Error GoTo error

    Ponderador = "1212121212"
    Calculo = 0

    If Len(Numero_Cuenta) <> 10 Then
        If Len(Numero_Cuenta) <> 8 Then
            FG_Digito_chequeo = "N"
            Exit Function
        End If
    End If

    For I = 1 To Len(Numero_Cuenta)
        numero = (Val(Mid$(Numero_Cuenta, I, 1)) * Val(Mid$(Ponderador, I, 1)))
        Calculo = Calculo + (numero Mod 10) + (numero \ 10)
    Next I

    Calculo = (10 - (Calculo Mod 10)) Mod 10

    FG_Digito_chequeo = Format$(Str(Calculo), "0")

    Exit Function
error:
    If Err.Number = 91 Then
        MsgBox "Existen Problemas de comunicación con la Base de datos , contacte al administrador de la red.  Ext. 3291", vbCritical
    Else
        MsgBox Err.Number & " " & Err.Description
    End If

End Function


Function FG_texto(entrada As Integer) As Integer

' PROPOSITO:    Valida que si un campo es de tipo
'               texto reciba solo letras, para números
'               o símbolos, devuelve el código Ascii con
'               valor cero para que el campo no muestre
'               ningún caracter.

    On Error GoTo error

    If (entrada = 8) Or (entrada = 32) Or (entrada >= 45) And (entrada <= 57) Or _
       ((entrada >= 65) And (entrada <= 90)) Or _
       ((entrada >= 97) And (entrada <= 122)) Or _
       (entrada = 209) Or (entrada = 241) Then
        FG_texto = entrada
    Else
        FG_texto = 0
    End If

    Exit Function
error:
    If Err.Number = 91 Then
        MsgBox "Existen Problemas de comunicación con la Base de datos , contacte al administrador de la red.  Ext. 3291", vbCritical
    Else
        MsgBox Err.Number & " " & Err.Description
    End If

End Function

Function FG_Digito_chequeoNit(Numero_Nit As Variant) As String
'---------------------------------------------------------------------------
'   NOMBRE: FG_Digito_chequeoNit
'
'   PROPOSITO: Calcula el Digito de chequeo para las cuentas contables
'
'   DESCRIPCION:
'   Multiplica cada uno de los digitos de la cuenta por su peso equivalente
'   al resultado de esta multiplicacion le obtiene el modulo 10, se lo suma
'   a la divicion enterea por 10 del mismo resultado y lo acumula en CALCULO
'   por ultimo se le resta a 10 el modulo 10 de CALCULO y a este resultado
'   se le calcula modulo de 10
'
'   PARAMETROS:
'   Numero_Nit = El parametro de entrada es el numero del nit
'
'   RETORNO:
'   La función retorna el digito de chequeo
'
'   AUTOR: Banco de Bogota
'   FECHA: Junio 11 de 1997
'---------------------------------------------------------------------------
    On Error GoTo error

    '   DEFINICION DE PONDERADORES
    Dim Ponderador() As Integer
    ReDim Ponderador(15) As Integer

    Ponderador(1) = 3
    Ponderador(2) = 7
    Ponderador(3) = 13
    Ponderador(4) = 17
    Ponderador(5) = 19
    Ponderador(6) = 23
    Ponderador(7) = 29
    Ponderador(8) = 37
    Ponderador(9) = 41
    Ponderador(10) = 43
    Ponderador(11) = 47
    Ponderador(12) = 53
    Ponderador(13) = 59
    Ponderador(14) = 67
    Ponderador(15) = 71


    '   Retorno invalido
    Numero_Nit = Trim(Numero_Nit)
    If (Len(Numero_Nit) = 0) Or (Not IsNumeric(Numero_Nit)) Then
        FG_Digito_chequeoNit = "N"
        Exit Function
    End If

    J = 1
    Calculo = 0
    For I = Len(Numero_Nit) To 1 Step -1
        Calculo = Calculo + (Val(Mid$(Numero_Nit, I, 1)) * Ponderador(J))
        J = J + 1
    Next I

    Digito = Calculo Mod 11
    If Digito > 1 Then
        Digito = 11 - (Calculo Mod 11)
    End If

    FG_Digito_chequeoNit = Format$(Str(Digito), "0")

    Exit Function
error:
    If Err.Number = 91 Then
        MsgBox "Existen Problemas de comunicación con la Base de datos , contacte al administrador de la red.  Ext. 3291", vbCritical
    Else
        MsgBox Err.Number & " " & Err.Description
    End If


End Function

Public Sub FG_Verifica_Formato_Numerico()

    On Error GoTo error
    Dim VALOR
    Dim CadErr As String

    VALOR = 1111.11
    If Mid(Format$(VALOR, "0,000.00"), 2, 1) <> "," Or Mid(Format$(VALOR, "0,000.00"), 6, 1) <> "." Then
        CadErr = "Formato Numérico Erroneo en el Panel de Control de Windows" & Chr(13) & Chr(13)
        CadErr = CadErr + "Correcto = 1,111.11  <----" & Chr(13)
        CadErr = CadErr + "Erroneo  = " & Format$(VALOR, "0,000.00") & Chr(13) & Chr(13)
        CadErr = CadErr + "El separador de Miles debe ser Coma (,)" & Chr(13)
        CadErr = CadErr + "El separador de Decimales debe ser Punto (.) " & Chr(13)
        CadErr = CadErr + "No. de digitos despues del decimal debe ser (2) "
        MsgBox CadErr
        End
    
    End If

    Exit Sub
error:
        MsgBox Err.Number & " " & Err.Description
    
End Sub



Function FG_Verifica_Formato_fecha() As Integer

    On Error GoTo error
    Dim fecha As Date
    Dim CadErr As String


    fecha = Format("16/02/2009", "General Date")
    sentencia = "SELECT TO_DATE( '" & fecha & "','DD/MM/YYYY') from DUAL"
    'MsgBox sentencia
    Set rsobj = cnObj1.Execute(sentencia)
    FG_Verifica_Formato_fecha = True

    Exit Function
error:
    If Err.Number = 91 Then
        MsgBox "Existen Problemas de comunicación con la Base de datos!", vbCritical
    Else
        MsgBox "Formato de fecha incorrecta.  Verifique que la configuración regional de su equipo se encuentre en Español- México"
        FG_Verifica_Formato_fecha = False
    End If


End Function

Public Sub actualizaFechaUltimaActividad()
    
    'Actualizamos el usuario con la fecha actual.
    sentencia = "UPDATE USRLARC.TBL_USUARIO SET FECHA_ULTIMA_ACTIVIDAD=TO_DATE('" & Format(Now, "yyyy/mm/dd HH:MM:SS") & "','YYYY/MM/DD HH24:MI:SS')" & _
    " WHERE COD_USR=" & Stg_cod_Usuario
    Set conexion = New ADODB.Connection
    conexion.ConnectionString = "provider = " & STG_PROVIDER_HOST & ";Data Source =" & STG_NOMBRE_BD_HOST & ";User ID=" & STG_USR_BASE_HOST & ";Password=" & STG_CLAVE_BASE_HOST & ";"
    'conexion.ConnectionString = "provider = MSDAORA ;Data Source =" & STG_NOMBRE_BD_HOST & ";User ID=" & STG_USR_BASE_HOST & ";Password=" & STG_CLAVE_BASE_HOST & ";"
    conexion.Open
    conexion.Execute (sentencia)
    
    'Cerramos objetos
    sentencia = ""
    conexion.Close
    Set conexion = Nothing
    
End Sub

Public Sub verificaConexiones()
    'Preguntamos por el numero de conexiones segun los parametros de la base de datos.
    sentencia = "SELECT  COUNT(*) NUMERO_CONEXIONES FROM V$SESSION  WHERE USERNAME NOT IN ('SYSTEM','DBSNMP') AND PROGRAM LIKE '%JDBC Thin Client%'"
    
    Dim numeroConexionesActuales As Integer
    
    Set recordSetObj = cargarRecordSet(sentencia)
    
    numeroConexionesActuales = recordSetObj("NUMERO_CONEXIONES")
    
    'Consultamos el parametro 41 para saber el numero de conexiones limite
    
    sentencia = "SELECT * FROM USRBNC.TBL_PARAMETRO WHERE COD_PARAMETRO=41"
    
    Dim numeroConexionesPermitidas As Integer
    
    Set recordSetObj = cargarRecordSet(sentencia)
    
    numeroConexionesPermitidas = recordSetObj("PARAMETRO_NUMERICO")
    
    'Si el numero de conexiones es mayor al limite, mostramos mensaje de parametro 42 y salimos de la aplicacion.
    
    If numeroConexionesActuales > numeroConexionesPermitidas Then
        sentencia = "SELECT * FROM USRBNC.TBL_PARAMETRO WHERE COD_PARAMETRO=42"
    
        Dim MENSAJE As String
        
        Set recordSetObj = cargarRecordSet(sentencia)
        
        MENSAJE = recordSetObj("PARAMETRO_ALFANUMERICO")
        
        MsgBox MENSAJE
        
        End
        
    End If
End Sub

Public Sub verificaReproceso()
    'Si el ambiente es de produccion, consulta el parametro de reproceso para mostrar el mensaje en el caso de que se haya habilitado
        Set rsobj = cnObj1.Execute("SELECT * from USRBNC.TBL_PARAMETRO WHERE COD_PARAMETRO=5")
        If Not (rsobj.EOF) Then
            If (rsobj("PARAMETRO_NUMERICO") = 1) Then
                lsMensaje = rsobj("PARAMETRO_ALFANUMERICO")
                lsMensaje = Replace(lsMensaje, "<br>", Strings.Chr(13))
                rsobj.Close
                cnObj1.Close
                MsgBox lsMensaje, vbOKOnly, "Información del sistema"
                
                End
            End If
        End If
    
End Sub

Public Function obtenerColumna(DATAG As DataGrid, nombreColumna As String) As Integer
    For indice = 0 To DATAG.Columns.Count - 1
        If DATAG.Columns(indice).Caption = nombreColumna Then
            obtenerColumna = indice
            Exit For
        End If
    Next
End Function

Public Sub cargaDatosDetalle(tipoSeleccionado As String, RSOBJ3 As ADODB.Recordset, conc As Boolean)

    Dim sqlTablas As String

    'Si se ha seleccionado tipo de movimiento Entrada o Mov Auxiliar, no se hace nada
    If tipoSeleccionado <> TR_ENTRADA And tipoSeleccionado <> MOV_LIBRO_AUXILIAR Then
        Set RSOBJ3 = Nothing

        'Si la consulta actual NO es sobre una conciliacion y el tipo es para codificar
        If Not conc And tipoSeleccionado = TR_CODIFICAR Then

            If ING_BD_CONSULTAR = True Then
                sqlTablas = " TBL_REGISTRO_TRADUCTOR A, TBL_TRANSACCION_TRADUCTOR B  "
            Else
                sqlTablas = " USRLARC.TBL_REGISTRO_TRADUCTOR@ARCHIST  A, TBL_TRANSACCION_TRADUCTOR B  "
            End If

            sentencia = "SELECT  A.COD_ENTIDAD,A.APLICACION_FUENTE,A.COD_TRANSACCION,A.FECHA_MOVIMIENTO,A.CENTRO_ORIGEN, " & _
                      " A.CENTRO_DESTINO,A.TIPO_CUENTA,A.COD_CUENTA,A.num_documento,A.FILLER," & _
                      " A.NUM_CAMPOS_MONETARIOS,A.TAM_CAMPOS_MONETARIOS,A.FECHA_SISTEMA, A.SEC_REGISTRO, B.DESC_TRANSACCION,A.COD_TIPO_REGISTRO  " & _
                      " FROM " & sqlTablas & _
                      " WHERE  A.COD_ENTIDAD = B.COD_ENTIDAD AND A.APLICACION_FUENTE = B.COD_APLICACION_FUENTE " & _
                      " AND A.COD_TRANSACCION = B.COD_TRANSACCION  " & _
                      " AND A.COD_ENTIDAD = 1 " & _
                      " AND A.APLICACION_FUENTE = '" & dbg_consulta.Columns(0) & "'" & _
                      " AND A.COD_TRANSACCION = '" & Format(dbg_consulta.Columns(1), "0###") & "'" & _
                      " AND FECHA_MOVIMIENTO = TO_DATE('" & Format(dbg_consulta.Columns(2), "DD/MM/YYYY") & "','DD/MM/YYYY') " & _
                      " AND FECHA_SISTEMA = TO_DATE('" & Format(dbg_consulta.Columns(3), "DD/MM/YYYY") & "','DD/MM/YYYY') " & _
                      " AND CENTRO_ORIGEN = " & Val(dbg_consulta.Columns(4)) & _
                      " AND COD_TIPO_REGISTRO = 1" & _
                      " AND CENTRO_DESTINO = " & Val(dbg_consulta.Columns(5))

            Set RSOBJ3 = cargarRecordSet(sentencia)
            Set rsdato = Nothing
            tipo_configuracion = 11
        Else
            If tipoSeleccionado = 9 Then
                If ING_BD_CONSULTAR = True Then
                    sqlTablas = " TBL_REGISTRO_TRADUCTOR A, TBL_TRANSACCION_TRADUCTOR B  "
                Else
                    sqlTablas = " USRLARC.TBL_REGISTRO_TRADUCTOR@ARCHIST  A, TBL_TRANSACCION_TRADUCTOR B  "
                End If

                sentencia = "SELECT  A.COD_ENTIDAD,A.APLICACION_FUENTE,A.COD_TRANSACCION,A.FECHA_MOVIMIENTO,A.CENTRO_ORIGEN," & _
                          " A.CENTRO_DESTINO,A.TIPO_CUENTA,A.COD_CUENTA,A.num_documento,A.FILLER," & _
                          " A.NUM_CAMPOS_MONETARIOS,A.TAM_CAMPOS_MONETARIOS, A.FECHA_SISTEMA, A.SEC_REGISTRO, B.DESC_TRANSACCION " & _
                          " FROM  " & sqlTablas & _
                          " WHERE  A.COD_ENTIDAD = B.COD_ENTIDAD AND A.APLICACION_FUENTE = B.COD_APLICACION_FUENTE " & _
                          " AND A.COD_TRANSACCION = B.COD_TRANSACCION  " & _
                          " AND A.COD_ENTIDAD = 1 " & _
                          " AND A.APLICACION_FUENTE = '" & dbg_consulta.Columns(2) & "'" & _
                          " AND A.COD_TRANSACCION = '" & Format(dbg_consulta.Columns(3), "0###") & "'" & _
                          " AND FECHA_SISTEMA = TO_DATE('" & dbg_consulta.Columns(5) & "','DD/MM/YYYY') " & _
                          " AND  FECHA_MOVIMIENTO = TO_DATE('" & Format(dbg_consulta.Columns(4), "DD/MM/YYYY") & "','DD/MM/YYYY') " & _
                          " AND CENTRO_ORIGEN = " & Val(dbg_consulta.Columns(6)) & _
                          " AND COD_TIPO_REGISTRO = 1" & _
                          " AND CENTRO_DESTINO = " & Val(dbg_consulta.Columns(7))

                Set RSOBJ3 = cargarRecordSet(sentencia)
                Set rsdato = Nothing
                tipo_configuracion = 11

            End If

            If tipoSeleccionado = TR_INCONSISTENTES Then

                If ING_BD_CONSULTAR = True Then
                    sqlTablas = " TBL_REGISTRO_TRADUCTOR A, TBL_ERROR B , TBL_TRANSACCION_TRADUCTOR C   "
                Else
                    sqlTablas = " USRLARC.TBL_REGISTRO_TRADUCTOR@ARCHIST  A, TBL_ERROR B , TBL_TRANSACCION_TRADUCTOR C  "
                End If

                sentencia = "SELECT  A.COD_ENTIDAD,A.APLICACION_FUENTE,A .COD_TRANSACCION,A.FECHA_MOVIMIENTO,A.CENTRO_ORIGEN," & _
                          " A.CENTRO_DESTINO,A.TIPO_CUENTA,A.COD_CUENTA,A.num_documento,A.FILLER," & _
                          " A.NUM_CAMPOS_MONETARIOS,A.TAM_CAMPOS_MONETARIOS, A.FECHA_SISTEMA,A.SEC_REGISTRO,A.COD_TIPO_REGISTRO,C.desc_transaccion" & _
                          " FROM " & sqlTablas & _
                          " WHERE  A.APLICACION_FUENTE = C.COD_APLICACION_FUENTE (+) AND A.COD_ENTIDAD = C.COD_ENTIDAD (+) AND A.COD_TRANSACCION = C.COD_TRANSACCION (+) AND nvl(FLAG_FILTRO,0) = 0 " & _
                          " AND A.COD_ENTIDAD = 1 " & _
                          " AND FECHA_SISTEMA = TO_DATE('" & Format(dbg_consulta.Columns(3), "DD/MM/YYYY") & "','DD/MM/YYYY') " & _
                          " AND A.APLICACION_FUENTE = '" & dbg_consulta.Columns(0) & "'" & _
                          " AND A.COD_TRANSACCION = '" & Format(dbg_consulta.Columns(1), "0###") & "'" & _
                          " AND  FECHA_MOVIMIENTO = TO_DATE('" & Format(dbg_consulta.Columns(2), "DD/MM/YYYY") & "','DD/MM/YYYY') and  nvl(FLAG_FILTRO,0) = 0 " & _
                          " AND CENTRO_ORIGEN = " & Val(dbg_consulta.Columns(4)) & _
                          " AND COD_TIPO_REGISTRO = 1" & _
                          " AND CENTRO_DESTINO = " & Val(dbg_consulta.Columns(5))

                Set RSOBJ3 = cargarRecordSet(sentencia)
                Set rsdato = Nothing
                tipo_configuracion = 5

            End If

            If tipoSeleccionado = TR_CODIFICAR Or tipoSeleccionado = TR_NORMALIZAR Then

                If ING_BD_CONSULTAR = True Then
                    sqlTablas = " TBL_CONCILIACION A, TBL_TRANSACCION_TRADUCTOR B "
                Else
                    sqlTablas = " USRLARC.TBL_CONCILIACION@ARCHIST  A, TBL_TRANSACCION_TRADUCTOR B  "
                End If

                sentencia = "SELECT  A.COD_ENTIDAD,A.APLICACION_FUENTE,A.COD_TRANSACCION,A.FECHA_MOVIMIENTO,A.CENTRO_ORIGEN," & _
                          " A.CENTRO_DESTINO,A.TIPO_CUENTA,A.COD_CUENTA,A.num_documento,A.FILLER," & _
                          " A.NUM_CAMPOS_MONETARIOS,A.TAM_CAMPOS_MONETARIOS, A.FECHA_SISTEMA, A.SEC_REGISTRO , B.DESC_TRANSACCION " & _
                          " FROM  " & sqlTablas & _
                          " WHERE  A.COD_ENTIDAD = B.COD_ENTIDAD AND A.APLICACION_FUENTE = B.COD_APLICACION_FUENTE  " & _
                          " AND A.COD_TRANSACCION = B.COD_TRANSACCION  " & _
                          " AND  COD_ENTIDAD_CONCIL = 1 " & _
                          " AND APLICACION_FUENTE_CONCIL = '" & dbg_consulta.Columns(0) & "'" & _
                          " AND COD_TRANSACCION_CONCIL = '" & Format(dbg_consulta.Columns(1), "0###") & "'" & _
                          " AND FECHA_SISTEMA = TO_DATE('" & Format(dbg_consulta.Columns(3), "DD/MM/YYYY") & "','DD/MM/YYYY') " & _
                          " AND  FECHA_MOVIMIENTO = TO_DATE('" & Format(dbg_consulta.Columns(2), "DD/MM/YYYY") & "','DD/MM/YYYY') " & _
                          " AND CENTRO_ORIGEN = " & Val(dbg_consulta.Columns(4)) & _
                          " AND CENTRO_DESTINO = " & Val(dbg_consulta.Columns(5))

                Set RSOBJ3 = cargarRecordSet(sentencia)


                Set rsdato = Nothing

                tipo_configuracion = 11

            End If
            If tipoSeleccionado = TR_CONCILIAR Then

                If ING_BD_CONSULTAR = True Then
                    sqlTablas = " TBL_CONCILIACION A, TBL_TRANSACCION_TRADUCTOR B  "
                Else
                    sqlTablas = " USRLARC.TBL_CONCILIACION@ARCHIST  A, TBL_TRANSACCION_TRADUCTOR B  "
                End If

                sentencia = "SELECT  A.COD_ENTIDAD,A.APLICACION_FUENTE,A.COD_TRANSACCION,A.FECHA_MOVIMIENTO,A.CENTRO_ORIGEN," & _
                          " A.CENTRO_DESTINO,A.TIPO_CUENTA,A.COD_CUENTA,A.num_documento,A.FILLER," & _
                          " A.NUM_CAMPOS_MONETARIOS,A.TAM_CAMPOS_MONETARIOS, A.FECHA_SISTEMA, A.SEC_REGISTRO , B.DESC_TRANSACCION " & _
                          " FROM " & sqlTablas & _
                          " WHERE  A.COD_ENTIDAD = B.COD_ENTIDAD AND A.APLICACION_FUENTE = B.COD_APLICACION_FUENTE  " & _
                          " AND A.COD_TRANSACCION = B.COD_TRANSACCION  " & _
                          " AND  COD_ENTIDAD_CONCIL = 1 " & _
                          " AND APLICACION_FUENTE_CONCIL = '" & dbg_consulta.Columns(0) & "'" & _
                          " AND COD_TRANSACCION_CONCIL = '" & Format(dbg_consulta.Columns(1), "0###") & "'" & _
                          " AND FECHA_SISTEMA = TO_DATE('" & Format(dbg_consulta.Columns(5), "DD/MM/YYYY") & "','DD/MM/YYYY') " & _
                          " AND  FECHA_MOVIMIENTO = TO_DATE('" & Format(dbg_consulta.Columns(4), "DD/MM/YYYY") & "','DD/MM/YYYY') " & _
                          " AND CENTRO_ORIGEN = " & Val(dbg_consulta.Columns(6)) & _
                          " AND CENTRO_DESTINO = " & Val(dbg_consulta.Columns(7))

                Set RSOBJ3 = cargarRecordSet(sentencia)

                Set rsdato = Nothing

                tipo_configuracion = 11

            End If
        End If

    End If

End Sub
