Attribute VB_Name = "variablesARC"
' Constantes para el Estado_comunicación (Modo Captura)
Public Const ING_FECHA_INVALIDA = 11005
Public Const Ing_Linea = 0
Public Const Ing_Fuera_Linea = 1
Public Const Ing_Linea_Central = 5

Global STG_RESPUESTA As String
Global ING_RESPUESTA As Integer
Global ING_BD_CONSULTAR As Boolean     ' 0 - Normal 1 Historica
Public STG_PATH_APL As String
Public STG_DIRECTORIO_SISTEMA As String
Public STG_NOMBRE_BD_HOST As String
Public STG_PROVIDER_HOST As String
Public STG_USR_BASE_HOST As String
Public STG_CLAVE_BASE_HOST As String
Public STG_PATH_SALIDAS As String
Public STG_PATH_ENTRADAS As String
Public STG_PATH_ENVIOS As String
Public STG_PATH_RECEPCION As String
Public STG_PATH_TRASMISION As String
Public STG_PATH_GENERACION As String
Public STG_PATH_BATCH As String
Public STG_PATH_REPORTES As String
Public STG_CONEXION_REPORTE As String
Public ING_TIMEOUT As Integer
Public ING_MODO_CAPTURA As Integer  '0 EN LINEA 1 FUERA LINEA 3 TMP
Public ing_tipo_validacion As Integer    '0 Central  1 Local
Public Stg_porcentaje_iva As String * 5
Public ING_VALIDA_LOCAL As Integer
Public rsobj As ADODB.Recordset
Public rsGrilla As ADODB.Recordset
Public rsGrillaCarga As ADODB.Recordset
Public cnObj1 As ADODB.Connection
Public stringConexion As String
Public STG_FILTRO As String
Public STG_FILTRO_BANA As String
Public SENTENCIA_FILTRO As String
Public ING_Estado As Integer
Public indicadorCambioClave As Integer
Public sentencia As String
Public sNombre_Computador As String
Public sUsuario_Windows As String
Public Ing_cod_entidad As Integer
Public Dtg_Fecha_movimiento As Date
Public Dtg_Fecha_movimiento1 As Date
Public Dtg_Fecha_Equipo As String
Public Ing_año As Integer
Public Stg_fecha1 As String
Public Dtg_fecha_sistema As Date
Global ambiente As String
Global Stg_cod_Usuario As String
Global Stg_nombre_Usuario As String
Global Stg_Clave_Usuario As String
Global Stg_Perfil_Usuario_Acceso As String         'Perfil del Usuario que Ingreso
Global Stg_Fecha_Vencimiento As String        'Fecha Vencimiento Usuario que Ingreso
Global Stg_Estado_Usuario As String           'Estado del Usuario que Ingreso

Public Ing_Nivel As Integer
Public Stg_Tabla As String
Public Stg_Where As String
Public CAMPOS_CONSULTA As String
Public FROM_CONSULTA As String
Public ORDER_CONSULTA As String

'Eventos para la bitacora
Public Const eventoPorDefecto = 0
Public Const eventoCreacionDeAplicacionFuente = 1
Public Const eventoModificacionDeAplicacionFuente = 2
Public Const eventoEliminacionDeAplicacionFuente = 3
Public Const eventoCreacionDeNegocio = 4
Public Const eventoModificacionDeNegocio = 5
Public Const eventoEliminacionDeNegocio = 6
Public Const eventoCreacionDeRequerimiento = 7
Public Const eventoModificacionDeRequerimiento = 8
Public Const eventoEliminacionDeRequerimiento = 9
Public Const eventoCreacionDeTransaccion = 10
Public Const eventoModificacionDeTransaccion = 11
Public Const eventoEliminacionDeTransaccion = 12
Public Const eventoEntregaDeTransaccion = 13
Public Const eventoDevolucionDeTransaccion = 14
Public Const eventoCreacionDeConciliacion = 15
Public Const eventoModificacionDeConciliacion = 16
Public Const eventoEliminacionDeConciliacion = 17
Public Const eventoCreacionDeTraduccion = 18
Public Const eventoModificacionDeTraduccion = 19
Public Const eventoEliminacionDeTraduccion = 20
Public Const eventoCertificacionDeTraduccion = 21
Public Const eventoRetiroCertificacionDeTraducción = 22
Public Const eventoEnvioAPruebas = 23
Public Const eventoCreacionUsuarioTransaccion = 24
Public Const eventoEliminacionUsuarioTransaccion = 25
Public Const eventoEnvioAProduccion = 26
Public Const eventoIngresoSistema = 27

'Constantes para la bitacora
Public Const claseAdministrativa = 1
Public Const claseTransaccional = 2
Public Const claseProcesos = 3
Public Const claseSistema = 4


'Constantes para las acciones
Public Const accionPorDefecto = 0
Public Const accionTransmisionExitosa = 1
Public Const accionTransmisionFallida = 2
Public Const accionProcesoFallido = 3
Public Const accionProcesoExitoso = 4
Public Const accionInicioProceso = 5
Public Const accionFinProceso = 6
Public Const accionTransaccionProcesada = 7
Public Const accionAccesoDenegado = 8
Public Const accionContrasenaErrada = 9
Public Const accionUsuarioNoExiste = 10
Public Const accionCreacionDeUsuario = 11
Public Const accionModificacionDeUsuario = 12
Public Const accionIngresoExitosoDeUsuarioARC = 13
Public Const accionIngresoExitosoDeUsuarioSALOC = 14

'Constantes para el tipo de movimiento de ARC
Public Const TR_ENTRADA = 1
Public Const TR_CONCILIAR = 2
Public Const TR_CODIFICAR = 3
Public Const TR_INCONSISTENTES = 4
Public Const TR_NORMALIZAR = 5
Public Const MOV_LIBRO_AUXILIAR = 6




Public Sub PL_Conexion_Oracle()
    On Error GoTo error
    stringConexion = "Provider=" & STG_PROVIDER_HOST & ";Data Source=" & STG_NOMBRE_BD_HOST & ";User ID=" & STG_USR_BASE_HOST & ";Password=" & STG_CLAVE_BASE_HOST
    
    Set cnObj1 = New ADODB.Connection
    cnObj1.Open stringConexion
    Exit Sub
error:
    Dim strErr As String
    strErr = "No se puede conectar a la base de datos." & Chr$(13) & Chr$(10) & Err.Number & ".  " & Err.Description
    Err.Raise 10014, "clsDbAccess.dbConnect", strErr
    Set cnObj1 = Nothing
End Sub
