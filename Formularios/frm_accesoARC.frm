VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frm_AccesoARC 
   BackColor       =   &H00800000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8580
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   11835
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawStyle       =   2  'Dot
   Icon            =   "frm_accesoARC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   MouseIcon       =   "frm_accesoARC.frx":0442
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   11835
   Begin VB.Frame frmIngresaAD 
      Height          =   3495
      Left            =   2160
      TabIndex        =   13
      Top             =   8400
      Visible         =   0   'False
      Width           =   7815
      Begin VB.TextBox txt_clave_arc 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   3720
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   15
         Top             =   1440
         Width           =   1935
      End
      Begin VB.CommandButton boton_autenticar_ad 
         Caption         =   "Autenticar"
         Height          =   495
         Left            =   1800
         TabIndex        =   16
         Top             =   2520
         Width           =   1335
      End
      Begin VB.CommandButton boton_cancelar_ad 
         Caption         =   "Cancelar"
         Height          =   495
         Left            =   4440
         TabIndex        =   17
         Top             =   2520
         Width           =   1335
      End
      Begin MSMask.MaskEdBox txt_usuario_arc 
         Height          =   375
         Left            =   3720
         TabIndex        =   14
         Top             =   960
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##########"
         PromptChar      =   " "
      End
      Begin VB.Label Label11 
         Caption         =   "Actualizacion usuario ARC"
         Height          =   255
         Left            =   3240
         TabIndex        =   20
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label14 
         Caption         =   "Usuario ARC:"
         Height          =   255
         Left            =   1680
         TabIndex        =   19
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label15 
         Caption         =   "Contraseña ARC:"
         Height          =   255
         Left            =   1680
         TabIndex        =   18
         Top             =   1560
         Width           =   1695
      End
   End
   Begin VB.TextBox txt_clave 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   7845
      MaxLength       =   50
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3622
      Width           =   1935
   End
   Begin VB.TextBox txt_nombre 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3540
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   3
      Top             =   5205
      Width           =   6225
   End
   Begin VB.CommandButton Cmd_Terminar 
      Caption         =   "&TERMINAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8025
      TabIndex        =   6
      Top             =   6330
      Width           =   1695
   End
   Begin VB.CommandButton Cmd_Cancelar 
      Caption         =   "&CANCELAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5145
      TabIndex        =   5
      Top             =   6330
      Width           =   1695
   End
   Begin VB.CommandButton Cmd_aceptar 
      Caption         =   "&INGRESAR"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2265
      TabIndex        =   4
      Top             =   6330
      Width           =   1695
   End
   Begin MSMask.MaskEdBox txt_usuario 
      Height          =   420
      Left            =   3585
      TabIndex        =   0
      Top             =   3600
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   741
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   " "
   End
   Begin VB.ComboBox comboDominios 
      Height          =   315
      Left            =   3600
      TabIndex        =   2
      Top             =   4440
      Width           =   6135
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "DOMINIO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   2160
      TabIndex        =   12
      Top             =   4395
      Width           =   1170
   End
   Begin VB.Label lblAmbiente 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00800000&
      Caption         =   "Ambiente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   255
      Left            =   6720
      TabIndex        =   11
      Top             =   2160
      Width           =   2655
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00800000&
      Caption         =   "Versión: Enero de 2018"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   255
      Left            =   5880
      TabIndex        =   10
      Top             =   1920
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   1380
      Left            =   1800
      Picture         =   "frm_accesoARC.frx":0884
      Top             =   600
      Width           =   7845
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "USUARIO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Index           =   0
      Left            =   2100
      TabIndex        =   9
      Top             =   3712
      Width           =   1275
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "CONTRASEÑA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   5865
      TabIndex        =   8
      Top             =   3682
      Width           =   1815
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "NOMBRE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   2115
      TabIndex        =   7
      Top             =   5280
      Width           =   1155
   End
End
Attribute VB_Name = "frm_AccesoARC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim numeroIntentos As Integer

Sub PL_CargaParametrosEstacion()

    Dim FNAME As String
    Dim FNUM As Integer
    Dim Linea As String

    Dim dirWinSystem As String
    Dim buf As String
    Dim ret As Long


    ' Obtener el directorio de System
    buf = String$(260, Chr$(0))
    ret = GetSystemDirectory(buf, Len(buf))
    dirWinSystem = Left$(buf, ret)
    STG_DIRECTORIO_SISTEMA = dirWinSystem
    STG_PATH_APL = STG_DIRECTORIO_SISTEMA & "\ARCN.INI"

    'Abre el archivo de configuracion ARCN.INI
    FNUM = FreeFile      ' Determine file number.
    FNAME = STG_PATH_APL
    On Error Resume Next
    Open FNAME For Input As FNUM

    If Err.Number = 0 Then
        On Error GoTo 0
        Line Input #FNUM, Linea
        STG_NOMBRE_BD_HOST = Mid$(Linea, 18, Len(Linea) - 11)
        Line Input #FNUM, Linea
        STG_PROVIDER_HOST = Mid$(Linea, 18, Len(Linea) - 11)
        Line Input #FNUM, Linea
        STG_USR_BASE_HOST = Mid$(Linea, 18, Len(Linea) - 11)
        Line Input #FNUM, Linea
        STG_CLAVE_BASE_HOST = Mid$(Linea, 18, Len(Linea) - 11)
        STG_CLAVE_BASE_HOST = fg_Desencripta(STG_CLAVE_BASE_HOST)
        Line Input #FNUM, Linea
        STG_PATH_ENTRADAS = Mid$(Linea, 18, Len(Linea) - 11)
        Line Input #FNUM, Linea
        STG_PATH_SALIDAS = Mid$(Linea, 18, Len(Linea) - 11)
        Line Input #FNUM, Linea
        STG_PATH_REPORTES = Mid$(Linea, 18, Len(Linea) - 11)
        Line Input #FNUM, Linea
        ING_TIMEOUT = Mid$(Linea, 18, Len(Linea) - 11)
        Close
    End If




    STG_CONEXION_REPORTE = "DSN = " & STG_NOMBRE_BD_HOST & ";UID = " & STG_USR_BASE_HOST & ";PWD = " & STG_CLAVE_BASE_HOST


    Exit Sub
    On Error GoTo 0

End Sub
Sub pl_cargar_variables_globales()

    On Error GoTo error
    'dsantan Preguntamos por el codigo de usuario dado el usuario_nt

    sentencia = "SELECT COD_USR FROM USRLARC.TBL_USUARIO WHERE USUARIO_NT='" & Trim(txt_usuario.Text) & "'"
    Dim recordSetUsuario As Recordset
    Set recordSetUsuario = cnObj1.Execute(sentencia)

    If (Not recordSetUsuario.EOF) Then

        Stg_cod_Usuario = recordSetUsuario("COD_USR")

    End If

    'Stg_cod_Usuario = Trim(txt_usuario)
    Stg_nombre_Usuario = rsobj("NOMBRE")
    txt_nombre = Stg_nombre_Usuario
    Stg_Clave_Usuario = rsobj("CLAVE")
    Stg_Fecha_Vencimiento = rsobj("FECHA_VENCE")
    Stg_Estado_Usuario = rsobj("ESTADO")
    Stg_Perfil_Usuario_Acceso = rsobj("COD_PERFIL_USUARIO")

    txt_nombre.Locked = True

    sentencia = "SELECT SYSDATE FROM DUAL"
    Dim rsobj2 As ADODB.Recordset
    Set rsobj2 = cnObj1.Execute(sentencia)
    Dtg_fecha_sistema = Format(rsobj2("SYSDATE"), "yyyy/mm/dd HH:MM:SS")
    Dtg_Fecha_movimiento = Format(rsobj2("SYSDATE"), "yyyy/mm/dd HH:MM:SS")
    rsobj2.Close
    Set rsobj2 = Nothing


    sentencia = "SELECT TO_CHAR(MIN(FECHA_CONTABLE),'dd/mm/yyyy') FROM USRBNC.TBL_FECHA WHERE GENERO_MINUTA = 0"
    Set rsobj2 = cnObj1.Execute(sentencia)
    Dtg_Fecha_movimiento = rsobj2(0)
    rsobj2.Close
    Set rsobj2 = Nothing
    Exit Sub


error:

    MsgBox Err.Number & " " & Err.Description
    txt_clave.Enabled = True
    txt_clave.SetFocus

End Sub

Private Sub boton_autenticar_ad_Click()
    autenticacionDoble
End Sub

Private Sub boton_cancelar_ad_Click()
    frmIngresaAD.Visible = False
End Sub



Private Sub cmd_aceptar_Click()
'MDI_Administrador_contable.Toolbar1.Enabled = True
 
    frm_AccesoARC.Visible = False
    
    If Not (MsgBox("Esta ingresando con el usuario " & Stg_cod_Usuario & " - " & Stg_nombre_Usuario & " ¿Desea continuar? ", vbQuestion + vbYesNo, " Confirmación de Salida ") = vbYes) Then
        End
    End If
    
    'CVAPD00007856 Se realiza llamado de procedimiento que actualiza el campo FECHA_ULTIMA_ACTIVIDAD de la tabla TBL_USUARIO
    actualizaFechaUltimaActividad
    
    PL_Conexion_Oracle
    
    'Se inserta registro en la bitacora para indicar que el usuario ingresó a la aplicacion.
    bitacora claseSistema, eventoIngresoSistema, accionIngresoExitosoDeUsuarioARC, "", "", "", "El Usuario " & Trim(Stg_cod_Usuario) & ", ingresó exitosamente a la aplicacion ARC. ", cnObj1
    
    cerrarObjeto cnObj1
    
    pl_habilita_menu
    
End Sub

Private Sub cmd_cancelar_clave_Click()
    End
End Sub

Private Sub Cmd_Cancelar_Click()
    txt_usuario.Text = "          "
    txt_clave.Text = ""
    txt_nombre.Text = ""
    txt_usuario.BackColor = &HFFFFFF
    txt_clave.BackColor = &HFFFFFF
    txt_clave.Locked = False
    Cmd_aceptar.Enabled = False
    txt_usuario.SetFocus

End Sub

Private Sub cmd_terminar_Click()
    Unload Me
    End
End Sub



Private Sub comboDominios_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        comboDominios.Text = ""
    End If
    If KeyAscii = vbKeyReturn Then
        txt_nombre.Enabled = True
        txt_nombre.SetFocus
    End If
End Sub





Private Sub txt_clave_arc_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        boton_autenticar_ad_Click
        Exit Sub
    End If
End Sub
Private Sub Form_Load()
    On Error GoTo error
    numeroIntentos = 0
    indicadorCambioClave = 0

    'dsantan: Obtenemos la lista de dominios
    Call SSPLogon.getDomainListInComboBox(comboDominios, "BANCODEBOGOTA")
    'Extrae la informacion del equipo y el usuario de red actuales
    Pg_Informacion_sistema
    'Modificamos la informacion de LOCALE para asegurar que los formatos de fechas y numeros sean los correctos
    Call FuncionesLocale.ModificaInformacionLocale
    
    'Verifica que el formato numerico sea correcto
    FG_Verifica_Formato_Numerico

    'Valida que el formato de fecha sea correcto para ARC
    fg_valida_tamaño_fecha
    'Detecta el directorio window\system32 para cargar la configuracion del archivo ARC.ini en donde se almacena la configuracion de
    'la base de datos a usar
    PL_CargaParametrosEstacion
    'Actualiza el la etiqueta con el ambiente de base de datos dependiendo de la configuracion cargada desde el ARC.ini
    If STG_NOMBRE_BD_HOST = "DSNMC" And STG_USR_BASE_HOST = "RESTORE" Then
        lblAmbiente = "HISTORICO"
    ElseIf STG_NOMBRE_BD_HOST = "DSNMC" Then
        lblAmbiente = "AMBIENTE DE DESARROLLO"
    ElseIf STG_NOMBRE_BD_HOST = "PRNMC" Then
        lblAmbiente = "AMBIENTE DE PRUEBAS"
    ElseIf STG_NOMBRE_BD_HOST = "PDMC" Then
        lblAmbiente = "AMBIENTE DE PRODUCCION"
    Else
        lblAmbiente = "Desconocido"
        MsgBox "ERROR: ¡Ambiente desconocido!", vbOKOnly, "Información del sistema"
        End
    End If

    'Se establece la conexion a la base de datos
    PL_Conexion_Oracle
   'Se verifica que el formato de fecha en la base de datos sea consistente con la configuracion regional establecida
    If Not FG_Verifica_Formato_fecha Then
        End
    End If
    
    'Si el mensaje por reproceso esta establecido, se muestra y sale de la aplicacion
    verificaReproceso
    
    'Se verifica el numero de conexiones a la base de datos y si supera el limite establecido por el parametro, se muestra
    'mensaje y se sale de la aplicacion
    verificaConexiones
     Exit Sub

error:

    If Err.Number = 10014 Then
        'fme_cambiar_clave.Visible = True
        MsgBox "Error en la configuración de la base de datos"
        End
        
    Else
        MsgBox Err.Description
    End If

End Sub



Private Sub txt_clave_GotFocus()
    txt_clave.SelStart = 0
    txt_clave.SelLength = Len(txt_usuario)

    If Len(Trim(txt_usuario)) = 0 Then
        MsgBox "Por favor Ingrese el Codigo de Usuario", vbInformation
        txt_usuario.Text = "          "
        txt_usuario.SetFocus
        Exit Sub
    End If

End Sub

Private Sub txt_clave_KeyPress(KeyAscii As Integer)
'KeyAscii = Asc(UCase(Chr(KeyAscii))) ' Mayuscula
    If KeyAscii = vbKeyEscape Then
        txt_clave.Text = ""
    End If
    If KeyAscii = vbKeyReturn Then
        comboDominios.SetFocus
    End If
End Sub


Private Sub txt_nombre_GotFocus()
    Dim clave_usr As String
    Dim sentencia As String
    Dim diasParaExpirar As String


    On Error GoTo error

    MousePointer = 13

    ' verifica que se haya ingresado un usuario
    If Trim(txt_usuario) = "" Then
        MsgBox "Por favor digite el usuario.", vbInformation
        txt_usuario.SetFocus
        Exit Sub
    End If

    'verifica que se haya ingresado la Clave
    If Trim(txt_clave) = "" Then
        MsgBox "Por favor digite la contraseña.", vbInformation
        txt_clave.SetFocus
        Exit Sub
    End If

    clave_usr = fg_Encripta(Trim(Me.txt_clave.Text))
    '------------------------------------------------------------------
    'Conexion a la base de datos
    PL_Conexion_Oracle

    'sentencia = "SELECT * from TBL_USUARIO WHERE COD_USR='" & Trim(txt_usuario.Text) & "'"
    'DSANTAN: REQ 11813: Consulto el usuario en la base ARC
    sentencia = "SELECT * from TBL_USUARIO WHERE USUARIO_NT='" & Trim(txt_usuario.Text) & "'"
    Set rsobj = cnObj1.Execute(sentencia)


    'Si Usuario existe
    If (Not rsobj.EOF) Then

        'Autentico AD y Traigo cedula y  continuo el programa.
        If rsobj("ESTADO") = "I" Then
            MsgBox "Acceso denegado, su usuario esta inactivo. Por favor contactar al Administrador de Usuarios.", vbCritical, "SEGURIDAD"
            End

        Else
            'If rsobj("CLAVE") <> clave_usr Then
            '    txt_clave.Enabled = True
            '    txt_clave.Text = ""
            '    txt_clave.SetFocus
            '    numeroIntentos = numeroIntentos + 1
            '    MsgBox "Contraseña errada, por favor digitela nuevamente.", vbCritical, "SEGURIDAD"
            '    txt_clave.Text = ""
            '    txt_clave.SetFocus
            '
            '    If numeroIntentos = 3 Then
            '        sentencia = "UPDATE TBL_USUARIO SET ESTADO='I' WHERE COD_USR='" & Trim(txt_usuario.Text) & "'"
            '        cnObj1.Execute (sentencia)
            '        MsgBox "Acceso denegado, si ha olvidado su contraseña, por favor hable con el Administrador de Usuarios.", vbCritical, "SEGURIDAD"
            '        End
            '    End If

            Dim flagAutenticadoAD

            flagAutenticadoAD = False

            If SSPLogon.authenticateUser(txt_usuario.Text, comboDominios.Text, txt_clave.Text) Then
                'MsgBox ("Autenticado!")
                flagAutenticadoAD = True
            Else
                'MsgBox ("No Autenticado!")
                flagAutenticadoAD = False
            End If

            If (Not flagAutenticadoAD) Then
                MsgBox "El usuario o la clave son incorrectos. Por favor intente de nuevo o contacte al administrador del Directorio Activo."
                txt_clave.Enabled = True
                txt_clave.Text = ""
                txt_clave.SetFocus
                txt_usuario = ""
            Else
                pl_cargar_variables_globales
                
                'Traemos los roles del usuario actual

                sentencia = "SELECT * FROM TBL_USUARIO_ROL " & _
                            "WHERE COD_USR='" & Stg_cod_Usuario & "'"

                Set rsobj = cnObj1.Execute(sentencia)

                Dim tipoUsuario As String

                'Para diferenciar los perfiles de usuario (roles) de saloc y arc se agrega el perfil 8:Consulta ARC Parcial (Antiguo Consulta en SALOC) y 9:Consulta ARC Total (Antiguo Analista de Contabilidad en SALOC)
                'recorremos los roles y verificamos si tiene rol de consulta parcial(8) o de consulta total(9)
                While Not rsobj.EOF
                    tipoUsuario = rsobj("COD_PERFIL_USUARIO")

                    'si el tipo de usuario iterado es 8 o 9
                    If (tipoUsuario = 8) Or (tipoUsuario = 9) Then
                        Stg_Perfil_Usuario_Acceso = tipoUsuario
                    End If
                    rsobj.MoveNext
                Wend
                'Aun si no encuentra el rol consulta parcial o consulta total en la tabla TBL_USUARIO_ROL,
                'ya el metodo pl_cargar_variables_globales ha cargado el que viene en la tabla TBL_USUARIO.
                'De esta forma se busca en las dos tablas


                'si no tiene ninguno de los roles, se sale de la aplicacion y muestra un mensaje que indica que el
                'usuario no tiene privilegios suficientes para entrar a ARC
                If Not ((Stg_Perfil_Usuario_Acceso = 8) Or (Stg_Perfil_Usuario_Acceso = 9)) Then
                    MsgBox "El usuario no tiene suficientes privilegios para ingresar a ARC."
                    txt_clave.Enabled = True
                    txt_clave.Text = ""
                    txt_usuario.SetFocus
                    txt_usuario = ""
                    MousePointer = 0
                    Exit Sub
                End If


                txt_clave.Locked = True
                txt_usuario.BackColor = &H80000004
                txt_clave.BackColor = &H80000004
                Cmd_aceptar.Enabled = True
                MDI_Administrador_contable.mnu_ingresar.Enabled = True
                Cmd_aceptar.SetFocus

               'Muestra un resúmen de la carga de archivos del día
                sentencia = "(SELECT COD_TIPO_REGISTRO  FROM TBL_CONTROL_ARCHIVO " & _
                            "WHERE FECHA_SISTEMA=TO_DATE('" & Format(Dtg_Fecha_movimiento, "DD/MM/YYYY") & "','DD/MM/YYYY') AND ESTADO_PROCESO='I' " & _
                            "UNION SELECT 7 FROM DUAL) " & _
                            "MINUS " & _
                            "SELECT COD_TIPO_REGISTRO  FROM TBL_CONTROL_ARCHIVO " & _
                            "WHERE FECHA_SISTEMA=TO_DATE('" & Format(Dtg_Fecha_movimiento, "DD/MM/YYYY") & "','DD/MM/YYYY') AND ESTADO_PROCESO='T' "


                Set rsobj = cnObj1.Execute(sentencia)

                Dim tipoRegistro As String
                Dim cadenaAlerta As String
                If Not rsobj.EOF Then
                    cadenaAlerta = "La información para la fecha contable '" & Format(Dtg_Fecha_movimiento, "DD/MM/YYYY") & "' no ha sido cargada completamente." & Chr(13) & "Los siguientes procesos no se han ejecutado correctamente: " & Chr(13)
                    While Not rsobj.EOF
                        tipoRegistro = rsobj("COD_TIPO_REGISTRO")
                        Select Case tipoRegistro
                        Case "1": cadenaAlerta = cadenaAlerta & Chr(13) & " ARCHIVO ENTRADAS "

                        Case "2": cadenaAlerta = cadenaAlerta & Chr(13) & " ARCHIVO CONCILIACION "

                        Case "3": cadenaAlerta = cadenaAlerta & Chr(13) & " ARCHIVO TRADUCTOR "

                        Case "4": cadenaAlerta = cadenaAlerta & Chr(13) & " ARCHIVO RECHAZOS "

                        Case "5": cadenaAlerta = cadenaAlerta & Chr(13) & " ARCHIVO LIBRO AUXILIAR COLGAAP "

                        Case "6": cadenaAlerta = cadenaAlerta & Chr(13) & " ARCHIVO LIBRO AUXILIAR COLGAAP BACK DATE "

                        Case "7": cadenaAlerta = cadenaAlerta & Chr(13) & " DETALLE DE PARTIDAS SIN CRUCE "
                        
                        Case "10": cadenaAlerta = cadenaAlerta & Chr(13) & " ARCHIVO LIBRO AUXILIAR IFRS "
                        
                        Case "11": cadenaAlerta = cadenaAlerta & Chr(13) & " ARCHIVO LIBRO AUXILIAR IFRS BACK DATE "
                        End Select
                        rsobj.MoveNext
                    Wend
                       MsgBox cadenaAlerta
                End If
            End If
            'End If
        End If

    Else
        'Deshabilitamos la autenticacion doble, ya que solamente tenia sentido para la migracion de usuarios
            'si no existe
            'Autenticacion doble (arc+ AD)
            'frmIngresaAD.Visible = True
            'txt_usuario_arc.SetFocus
        
        'Codigo anterior
        MsgBox "El usuario " & txt_usuario.Text & " no está registrado en ARC, por favor digitelo nuevamente o contacte al Administrador de la aplicación", vbInformation, "Información de Acceso"
        txt_usuario = "          "
        txt_clave.Text = ""
        txt_usuario.SetFocus
        MousePointer = 0
        Exit Sub

    End If
    MousePointer = 0

    '------------------------------------------------------------------
    'Se cierra el recordSet y la conexion
    rsobj.Close
    cnObj1.Close
    Set rsobj = Nothing
    Set cnObj1 = Nothing

    Exit Sub

error:
    MsgBox Err & " " & Err.Description
    End


End Sub

Sub autenticacionDoble()
'Autenticamos Active Directory
'DSANTAN:Autenticamos con Active Directory
    Dim flagAutenticadoAD

    flagAutenticadoAD = False

    txt_usuario.Text = Trim(txt_usuario.Text)

    If SSPLogon.authenticateUser(txt_usuario.Text, comboDominios.Text, txt_clave.Text) Then
        'MsgBox ("Autenticado!")
        flagAutenticadoAD = True
    Else
        'MsgBox ("No Autenticado!")
        flagAutenticadoAD = False
    End If


    'FIN DSANTAN:Autenticamos con Active Directory
    'Autenticamos en ARC

    Dim flagAutenticadoARC

    flagAutenticadoARC = False

    sentencia = "select COD_USR, COD_ENTIDAD, U.COD_PERFIL_USUARIO,DESC_PERFIL_USUARIO, NOMBRE, CLAVE, ESTADO, FECHA_VENCE, ASIGNACION_TOTAL " & _
                "from TBL_USUARIO U, TBL_PERFIL_USUARIO P " & _
                "where U.COD_PERFIL_USUARIO=P.COD_PERFIL_USUARIO and COD_ENTIDAD=1 and U.COD_USR='" & Trim(txt_usuario_arc.Text) & "'"
    Set rsobj = cnObj1.Execute(sentencia)

    'txt_usuario_arc.Text = Trim(txt_usuario_arc.Text)

    Dim claveEncriptada

    claveEncriptada = fg_Encripta(Trim(txt_clave_arc.Text))

    If Not rsobj.EOF Then
        If UCase(rsobj("CLAVE")) = UCase(claveEncriptada) Then

            flagAutenticadoARC = True
        End If
    End If
    'si ambas autenticaciones son correctas
    If flagAutenticadoAD And flagAutenticadoARC Then
        'Relacionamos usuario NT con la cedula introducida
        sentencia = " UPDATE TBL_USUARIO " & _
                  " SET USUARIO_NT = '" & txt_usuario.Text & "'" & _
                  " WHERE COD_USR = '" & txt_usuario_arc.Text & "' " & _
                  " AND COD_ENTIDAD=1"
        Set rsobj = cnObj1.Execute(sentencia)

        bitacora claseAdministrativa, eventoIngresoSistema, accionModificacionDeUsuario, "", "", "", "El Usuario ARC :" & Trim(txt_usuario_arc.Text) & ", fue actualizado con el UsuarioNT: " & txt_usuario.Text, cnObj1

        MsgBox "Usuario Actualizado. Por favor ingrese de nuevo con su usuario de Directorio Activo."

        frmIngresaAD.Visible = False


        'sino
    Else
        'mostramos error
        If Not flagAutenticadoAD Then
            MsgBox "Usuario o clave de directorio activo incorrectos. Por favor verifique e intente de nuevo.", vbCritical, "SEGURIDAD"
        End If

        If Not flagAutenticadoARC Then
            MsgBox "Usuario o clave de ARC incorrectos. Por favor verifique e intente de nuevo.", vbCritical, "SEGURIDAD"
        End If

        'Limpiamos las claves para obligar al usuario a verificar su clave.
        'txt_clave_nt.Text = ""
        txt_clave_arc.Text = ""

    End If


End Sub


Private Sub txt_usuario_arc_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txt_clave_arc.SetFocus
        Exit Sub
    End If
End Sub

Private Sub txt_usuario_GotFocus()
    txt_usuario.SelStart = 0
    txt_usuario.SelLength = Len(txt_usuario)
End Sub
Private Sub txt_usuario_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = vbKeyEscape Then
        txt_usuario.Text = "          "
    End If
    If KeyAscii = vbKeyReturn Then
        txt_clave.SetFocus
        Exit Sub
    End If
End Sub

