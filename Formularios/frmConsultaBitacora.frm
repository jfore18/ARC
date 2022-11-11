VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmConsultaBitacora 
   Caption         =   "Consulta Bitácora"
   ClientHeight    =   8670
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15420
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   15420
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Caption         =   "Resultado de la búsqueda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   6375
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   14775
      Begin MSDataGridLib.DataGrid grll_bitacora 
         Height          =   5535
         Left            =   240
         TabIndex        =   12
         Top             =   480
         Width           =   14175
         _ExtentX        =   25003
         _ExtentY        =   9763
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1335
      Left            =   6000
      TabIndex        =   1
      Top             =   120
      Width           =   4095
      Begin VB.CommandButton btn_salir 
         Caption         =   "Salir"
         Height          =   615
         Left            =   2760
         TabIndex        =   11
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton btn_limpiar 
         Caption         =   "Limpiar"
         Height          =   615
         Left            =   1320
         TabIndex        =   10
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton btn_buscar 
         Caption         =   "Buscar"
         Height          =   615
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame frm_Busqueda 
      Caption         =   "Búsqueda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1815
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      Begin MSMask.MaskEdBox mskFecha 
         Height          =   315
         Left            =   1200
         TabIndex        =   8
         Top             =   1320
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cmb_accion 
         Height          =   315
         Left            =   1200
         TabIndex        =   7
         Text            =   "Seleccione"
         Top             =   840
         Width           =   4095
      End
      Begin VB.ComboBox cmb_evento 
         Height          =   315
         Left            =   1200
         TabIndex        =   6
         Text            =   "Seleccione"
         ToolTipText     =   "Eventos"
         Top             =   360
         Width           =   4095
      End
      Begin VB.Label lbl_fecha 
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lbl_Accion 
         Caption         =   "Acción"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lbl_evento 
         Caption         =   "Evento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmConsultaBitacora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_buscar_Click()
'Se validan los datos de búsqueda
    If validarCampos Then
        'Se buscan los registros de bitácora que cumplan los criterios especificados
         cargarGrilla
    End If
End Sub

'Subrutina para limpiar los datos del formulario
Private Sub btn_limpiar_Click()
 Me.cmb_evento.ListIndex = -1
 Me.cmb_accion.ListIndex = -1
 Me.mskFecha = "__/__/____"
End Sub

' Opcion para salir de la consulta de bitacora
Private Sub btn_salir_Click()
Unload Me
End Sub

Private Sub DataGrid1_Click()

End Sub
' REQUERIMIENTO: REQCVAPD00337472
' Modificado por: Ana María Bocanegra
' Diciembre de 2017
' Se construye nueva funcionalidad en el aplicativo para monitoreo de los procesos de carga
Private Sub Form_Load()
'Asignamos la fecha del día
Me.mskFecha.Text = Format(Date, "dd/mm/yyyy")
'Asingacion de valor a la clase de bitacora: 3-Procesos
'Cargamos la lista de eventos
 cargueEventos
'Cargamos la lista de acciones
 cargueAcciones
End Sub

'Subrutina para cargar los eventos correspondientes al cargue de archivos de ARC
 Private Sub cargueEventos()
 'Conexión con la base de datos
  PL_Conexion_Oracle
 'Declaración de variables para las consultas
  Dim sentencia As String
 
 sentencia = "SELECT CODIGO,DESCRIPCION FROM USRLARC.TBL_EVENTO WHERE CODIGO BETWEEN 36 AND 52 ORDER BY 1"
 Set rsobj = cnObj1.Execute(sentencia)
 Me.cmb_evento.Clear
 Do While Not rsobj.EOF
        Me.cmb_evento.AddItem rsobj("CODIGO") & " - " & rsobj("DESCRIPCION")
        rsobj.MoveNext
 Loop

 Me.cmb_evento.ListIndex = -1
    rsobj.Close
    cnObj1.Close
    Set rsobj = Nothing
    Set cnObj1 = Nothing
End Sub
'Subrutina para cargar las acciones correspondiente al cargue de archivos de ARC
Private Sub cargueAcciones()
'Conexión con la base de datos
  PL_Conexion_Oracle
 'Declaración de variables para las consultas
  Dim sentencia As String
 
 sentencia = "SELECT CODIGO,DESCRIPCION FROM USRLARC.TBL_ACCION WHERE CODIGO IN(5,6,16,17)ORDER BY CODIGO"
 Set rsobj = cnObj1.Execute(sentencia)
 Me.cmb_accion.Clear
 Do While Not rsobj.EOF
        Me.cmb_accion.AddItem rsobj("CODIGO") & " - " & rsobj("DESCRIPCION")
        rsobj.MoveNext
 Loop

 Me.cmb_accion.ListIndex = -1
    rsobj.Close
    cnObj1.Close
    Set rsobj = Nothing
    Set cnObj1 = Nothing
End Sub
' Subrutina para validar que se haya seleccionado una feccha para la consulta de los procesos de ARC
Private Function validarCampos() As Boolean
    On Error GoTo error
    If Me.mskFecha = "__/__/____" Then
      MsgBox "Debe seleccionar una fecha de consulta"
      Me.mskFecha.SetFocus
      validarCampos = False
      Exit Function
    End If
    validarCampos = True
    Exit Function
error:
    validarCampos = False
    MsgBox vbCrLf & Err.Number & " . " & Err.Description
End Function

' Subrutina para configurar la grilla de resultados

Sub configurarGrilla()
    Dim I As Integer
    Me.grll_bitacora.Columns(0).Caption = "Fecha Sistema"
    Me.grll_bitacora.Columns(1).Caption = "Hora Sistema"
    Me.grll_bitacora.Columns(2).Caption = "Evento"
    Me.grll_bitacora.Columns(3).Caption = "Accion"
    Me.grll_bitacora.Columns(4).Caption = "Detalle"

    Me.grll_bitacora.Columns(0).Width = 1400
    Me.grll_bitacora.Columns(1).Width = 1400
    Me.grll_bitacora.Columns(2).Width = 4000
    Me.grll_bitacora.Columns(3).Width = 2000
    Me.grll_bitacora.Columns(4).Width = 5000
    'Configuracion formatos
    Me.grll_bitacora.Columns(0).NumberFormat = ("DD/MM/YYYY")
    'Me.grll_bitacora.Columns(1).NumberFormat = ("HH:MM:SS")
   
    'Bloquea y vuelve visible todas las columnas
    For I = 0 To 4
        grll_bitacora.Columns(I).Visible = True
        grll_bitacora.Columns(I).Locked = True
    Next I
    'Alineación de las columnas
    grll_bitacora.Columns(0).Alignment = dbgLeft
    grll_bitacora.Columns(1).Alignment = dbgLeft
    grll_bitacora.Columns(2).Alignment = dbgLeft
    grll_bitacora.Columns(3).Alignment = dbgLeft
    grll_bitacora.Columns(4).Alignment = dbgLeft
End Sub
'Subrutina para cargar la grilla con los resultados de la consulta
Private Sub cargarGrilla()
'Definicion de variables
 Dim sentencia As String
 Dim criterioBusqueda As Boolean

 criterioBusqueda = False
 Screen.MousePointer = vbHourglass

 'Se construye la sentencia dependiendo de los critérios de búsqueda
  sentencia = "SELECT FECHA_SISTEMA,HORA,EVENTO,ACCION,DETALLE FROM USRLARC.VTA_BITACORA WHERE CODIGO_CLASE='3' AND FECHA_SISTEMA=to_date('" & Me.mskFecha.Text & "','DD/MM/YYYY')"
    
  'Si no selecciona evento, se asignan los eventos correspondientes al cargue de ARC: del 36 al 52
   If Me.cmb_evento.ListIndex = -1 Then
    sentencia = sentencia & "and CODIGO_EVENTO BETWEEN 36 AND 52"
   End If
  
  'Filtro por evento
   If Me.cmb_evento.ListIndex <> -1 And Trim(Me.cmb_evento.Text) <> "" Then
    sentencia = sentencia & "and CODIGO_EVENTO='" & Left(Me.cmb_evento.Text, 2) & "'"
   End If
  
  'Filtro por acción
   If Me.cmb_accion.ListIndex <> -1 And Trim(Me.cmb_accion.Text) <> "" Then
    sentencia = sentencia & "and CODIGO_ACCION ='" & Left(cmb_accion.Text, 2) & "'"
   End If
   sentencia = sentencia & "order by hora"
  Set rsGrilla = cargarRecordSet(sentencia)
  If rsGrilla.EOF Then
        grll_bitacora.ClearFields
        grll_bitacora.Refresh
        grll_bitacora.Visible = False
        MsgBox "No existen registros en la bitácora que cumplan el críterio de búsqueda.", vbOKOnly, "Información de búsqueda"
    Else
        Set grll_bitacora.DataSource = rsGrilla
        grll_bitacora.Visible = True
        grll_bitacora.Refresh
        configurarGrilla
    End If
    Screen.MousePointer = vbDefault
End Sub

