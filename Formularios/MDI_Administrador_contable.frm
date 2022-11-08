VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.MDIForm MDI_Administrador_contable 
   BackColor       =   &H8000000C&
   Caption         =   "Administrador de Relaciones Contables"
   ClientHeight    =   8445
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11070
   Icon            =   "MDI_Administrador_contable.frx":0000
   LinkTopic       =   "MDI_Libros_Auxiliares"
   MouseIcon       =   "MDI_Administrador_contable.frx":08CA
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   3960
      Top             =   4920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Menu mnu_archivo 
      Caption         =   "Archivo"
      Begin VB.Menu mnu_ingresar 
         Caption         =   "Ingresar"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnu_salir 
         Caption         =   "Terminar"
      End
   End
   Begin VB.Menu Mnu_consulta 
      Caption         =   "Movimiento Traductor"
      Begin VB.Menu mnu_consulta_general 
         Caption         =   "Consulta de Movimiento"
      End
   End
   Begin VB.Menu mnuResponsabilidades 
      Caption         =   "Responsabilidades"
      Begin VB.Menu menuConsultaResponsabilidades 
         Caption         =   "Consulta de responsabilidades"
      End
   End
   Begin VB.Menu mnuProcesos 
      Caption         =   "Monitoreo de Procesos"
      Begin VB.Menu mnuProcesosCarga 
         Caption         =   "Consulta Procesos de Carga"
      End
      Begin VB.Menu mnuBitacora 
         Caption         =   "Consulta Bitacora"
      End
   End
End
Attribute VB_Name = "MDI_Administrador_contable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Ing_Nivel As Integer

Private Sub cmb_calculadora_DropDown()
    frm_calculadora.Show
    frm_calculadora.Move 8000, 50
End Sub

Private Sub MDIForm_Load()
    MousePointer = 13
    pl_deshabilita_menu
    frm_AccesoARC.Show
    MousePointer = 0
End Sub


Sub Pl_desconecta_usuario()
    End

End Sub
Private Sub MDIForm_Unload(Cancel As Integer)
    End
End Sub

Private Sub menuConsultaResponsabilidades_Click()
'Despliega la forma de consulta de responsabilidades
'descargarFormularios
    MousePointer = 13
    frmConsultaResponsabilidades.Show
    MousePointer = 0
End Sub


Private Sub mnu_Captura_fias_Click()
'descargarFormularios
    frm_Captura_fias.Show
End Sub

Private Sub Mnu_consulta_fias_Click()
    descargarFormularios
    frm_consulta_fias_ARC.Show
End Sub

Private Sub mnu_consulta_general_Click()
'si el usuario tiene rol de consulta parcial, validamos que tenga asignaciones en TBL_USUARIO_MOVIMIENTO y TBL_USUARIO_TRANSACCION
    If Stg_Perfil_Usuario_Acceso = 8 Then
        'Valida que el usuario tenga algún tipo de movimiento asignado en la tabla TBL_USUARIO_MOVIMIENTO
        If Not (verificaTipoMovimientoAsignado) Then
            MsgBox "El usuario " & Stg_cod_Usuario & " no tiene ningún tipo de movimiento asignado. por favor contacte al administrador de usuarios ARC."
            Exit Sub
        End If

        'Valida que el usuario tenga alguna transaccion asociada en la tabla TBL_USUARIO_TRANSACCION
        If Not (verificaTransacciones) Then
            MsgBox "El usuario " & Stg_cod_Usuario & " no tiene ninguna transaccion asociada . por favor contacte al administrador de usuarios ARC."
            Exit Sub
        End If
    End If
    'Despliega la forma de consulta de movimiento
    'descargarFormularios
    frm_Consulta_general_ARC.Show
End Sub

Private Sub mnu_filtro_Click()
    On Error Resume Next
    Unload MDI_Administrador_contable.ActiveForm
    frm_filtro_transaccion.Show
    'frm_Recepcion_archivos_ARC.Move 1800, 700
    MousePointer = 0
End Sub

Private Sub mnu_ingresar_Click()
    On Error Resume Next
    If Not (MDI_Administrador_contable.ActiveForm Is Nothing) Then
        Unload MDI_Administrador_contable.ActiveForm
    End If

    If ING_MODO_CAPTURA = Ing_Linea Or ING_MODO_CAPTURA = Ing_Linea_Central Then
        If Not (frm_Acceso_ARC = Empty) Then
            frm_Acceso_ARC.Show
        End If
    Else
        MsgBox "Operacion no se puede realizar Fuera de Linea", vbInformation
    End If
End Sub

Private Sub mnu_recepcion_archivos_Click()
    On Error Resume Next
    Unload MDI_Administrador_contable.ActiveForm
    frm_Recepcion_archivos_ARC.Show
    'frm_Recepcion_archivos_ARC.Move 1800, 700
    MousePointer = 0
End Sub

Private Sub mnu_salir_Click()
    mb_valor = vbQuestion + vbYesNo
    mB_mensaje = MsgBox("Está seguro que desea salir de la aplicación?", mb_valor, "Administrador de relaciones contables")
    If mB_mensaje = vbYes Then
        Pl_desconecta_usuario
        End
    End If
End Sub

Private Sub mnuProcesoCarga_Click()
'Despliega la forma de consulta del proceso de carga
    descargarFormularios
    frmConsultaCarga.Show
End Sub

Private Sub mnuBitacora_Click()
descargarFormularios
frmConsultaBitacora.Show
End Sub

Private Sub mnuProcesosCarga_Click()
'Despliega la forma de consulta del proceso de carga
    descargarFormularios
    frmConsultaCarga.Show
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "Salir"
        mb_valor = vbQuestion + vbYesNo
        mB_mensaje = MsgBox("Está seguro que desea salir de la aplicación?", mb_valor, "Administrador de relaciones contables")
        If mB_mensaje = vbYes Then
            Pl_desconecta_usuario
            End
        End If
    End Select
End Sub

Private Sub txt_clave_GotFocus()
    txt_clave.SelStart = 0
    txt_clave.SelLength = Len(txt_clave)
End Sub

Private Sub txt_clave_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))    ' Mayuscula
    If KeyAscii = vbKeyEscape Then
        txt_clave = "    "
        txt_clave.SetFocus
    Else
        If KeyAscii = vbKeyReturn Then
            If txt_usuario = "" Then
                MsgBox "Digite el Código Usuario"
                txt_usuario.SetFocus
            Else
                Pl_valida_usuario
            End If
        End If
    End If
End Sub

Sub Pl_valida_usuario()
' verifica que se haya ingresado un usuario

    Dim stl_set As String
    Dim Stg_Where As String

    If Trim(txt_usuario) = "" Then
        MsgBox " Por favor digite un usuario", vbInformation
        txt_usuario.SetFocus
        Exit Sub
    End If

    'verifica que se haya ingresado la Clave
    If Trim(txt_clave) = "" Then
        MsgBox " Por favor digite una Clave", vbInformation
        txt_clave.SetFocus
        Exit Sub
    End If

    ' Verifica que el usuario y clave existan en la tabla de usuarios

    sentencia = "SELECT * FROM  TBL_USUARIO WHERE COD_USR='" & Trim(txt_usuario.Text) & "'"
    Set rsobj = cnObj1.Execute(sentencia)

    If rsobj.EOF Then
        MsgBox " El Usuario no Existe, Por favor digite Nuevamente su Usuario", vbInformation
        Set OBJDATO = Nothing
        txt_usuario.SetFocus
    Else

        sentencia = "SELECT * FROM  TBL_USUARIO WHERE COD_USR='" & Trim(txt_usuario.Text) & "' AND CLAVE_USR = '" & txt_clave.Text & "'"
        Set rsobj = cnObj1.Execute(sentencia)

        ''ing_Valor_existeUsuario = OBJDATO.fg_Existe_2("TBL_USUARIO", "COD_USR", txt_usuario.Text, "CLAVE_USR", txt_clave.Text, ing_tipo_validacion)
        ''Set OBJDATO = Nothing
        If rsobj.EOF Then
            MsgBox " Clave Errada, Por favor Digitela Nuevamente", vbInformation
            txt_clave.Enabled = True
            txt_clave.Text = ""
            txt_clave.SetFocus
            num_intentos = num_intentos + 1
        End If
    End If


    txt_usuario = ""
    txt_clave = ""
    fme_usuario.Visible = False


    MousePointer = 0
    Exit Sub

error:
    MsgBox Err & " " & Err.Description

End Sub

Private Sub txt_usuario_GotFocus()
    txt_usuario.SelStart = 0
    txt_usuario.SelLength = Len(txt_usuario)
End Sub

Private Sub txt_usuario_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))    ' Mayuscula
    If KeyAscii = vbKeyEscape Then
        txt_usuario = "    "
        txt_usuario.SetFocus
    Else
        If KeyAscii = vbKeyReturn Then
            txt_clave.SetFocus
        End If
    End If
End Sub

