VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmConsultaCarga 
   Caption         =   "Consulta de usuarios"
   ClientHeight    =   7635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Caption         =   "Procesos de carga fallidos"
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
      Height          =   3075
      Left            =   120
      TabIndex        =   7
      Top             =   4200
      Width           =   10455
      Begin MSDataGridLib.DataGrid grillaResultadosFallidos 
         Height          =   2535
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   4471
         _Version        =   393216
         ForeColor       =   8388608
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
   Begin VB.Frame Frame2 
      Caption         =   "Parámetros de consulta"
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
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   10455
      Begin VB.CommandButton Command2 
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5520
         TabIndex        =   5
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Consultar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
      Begin MSMask.MaskEdBox msk_Fecha_Proceso 
         Height          =   315
         Left            =   1320
         TabIndex        =   0
         Top             =   360
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha proceso:"
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
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Procesos de carga exitosos"
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
      Height          =   2835
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   10455
      Begin MSDataGridLib.DataGrid grillaResultadosExitosos 
         Height          =   2295
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   4048
         _Version        =   393216
         AllowUpdate     =   -1  'True
         ForeColor       =   8388608
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
End
Attribute VB_Name = "frmConsultaCarga"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub configurarGrilla()
    Dim I As Integer
    Me.grillaResultadosExitosos.Columns(0).Caption = "Proceso"
    Me.grillaResultadosExitosos.Columns(1).Caption = "Fecha"
    Me.grillaResultadosExitosos.Columns(2).Caption = "Hora"
    Me.grillaResultadosExitosos.Columns(3).Caption = "Fecha contable"
    Me.grillaResultadosExitosos.Columns(4).Caption = "Cantidad de registros"

    Me.grillaResultadosExitosos.Columns(0).Width = 3000
    Me.grillaResultadosExitosos.Columns(1).Width = 1500
    Me.grillaResultadosExitosos.Columns(2).Width = 1300
    Me.grillaResultadosExitosos.Columns(3).Width = 1500
    Me.grillaResultadosExitosos.Columns(4).Width = 2200
    'Configuracion formatos
    Me.grillaResultadosExitosos.Columns(1).NumberFormat = ("DD/MM/YYYY")
    Me.grillaResultadosExitosos.Columns(2).NumberFormat = ("HH:MM:SS")
    Me.grillaResultadosExitosos.Columns(3).NumberFormat = ("DD/MM/YYYY")
    Me.grillaResultadosExitosos.Columns(4).NumberFormat = ("###,###,###,###")


    'Bloquea y vuelve visible todas las columnas
    For I = 0 To 4
        grillaResultadosExitosos.Columns(I).Visible = True
        grillaResultadosExitosos.Columns(I).Locked = True
    Next I
    'Alineación de las columnas
    grillaResultadosExitosos.Columns(0).Alignment = dbgLeft
    grillaResultadosExitosos.Columns(1).Alignment = dbgLeft
    grillaResultadosExitosos.Columns(2).Alignment = dbgLeft
    grillaResultadosExitosos.Columns(3).Alignment = dbgLeft
    grillaResultadosExitosos.Columns(4).Alignment = dbgRight
End Sub


Sub configurarGrilla2()
    Dim I As Integer
    Me.grillaResultadosFallidos.Columns(0).Caption = "Proceso"
    Me.grillaResultadosFallidos.Columns(1).Caption = "Fecha"
    Me.grillaResultadosFallidos.Columns(2).Caption = "Hora"
    Me.grillaResultadosFallidos.Columns(3).Caption = "Fecha contable"

    Me.grillaResultadosFallidos.Columns(0).Width = 3000
    Me.grillaResultadosFallidos.Columns(1).Width = 1500
    Me.grillaResultadosFallidos.Columns(2).Width = 1300
    Me.grillaResultadosFallidos.Columns(3).Width = 1500


    'Configuracion formatos
    Me.grillaResultadosFallidos.Columns(1).NumberFormat = ("DD/MM/YYYY")
    Me.grillaResultadosFallidos.Columns(3).NumberFormat = ("DD/MM/YYYY")

    'Bloquea y vuelve visible todas las columnas
    For I = 0 To 3
        grillaResultadosFallidos.Columns(I).Visible = True
        grillaResultadosFallidos.Columns(I).Locked = True
    Next I
    'Alineación de las columnas
    grillaResultadosFallidos.Columns(0).Alignment = dbgLeft
    grillaResultadosFallidos.Columns(1).Alignment = dbgCenter
    grillaResultadosFallidos.Columns(2).Alignment = dbgCenter
    grillaResultadosFallidos.Columns(3).Alignment = dbgCenter
End Sub

Private Sub consultarCargue()
    Me.MousePointer = 11
    Dim sentencia As String

    If (msk_Fecha_Proceso.Text = "  /  /    ") Then
        MsgBox "Debe diligenciar la fecha de proceso."
        Exit Sub
    End If

    sentencia = "SELECT  DECODE(COD_TIPO_REGISTRO," & _
                "'1',' Entradas'," & _
                "'2',' Conciliación'," & _
                "'3',' Traductor'," & _
                "'4',' Rechazos'," & _
                "'5',' Libro Auxiliar'," & _
                "'6',' Libro Auxiliar Back Date'," & _
                "'7',' Generación de partidas sin cruce'),FECHA_SISTEMA,HORA_PROCESO," & _
                "FECHA_MOVIMIENTO,CANT_REGISTROS FROM TBL_CONTROL_ARCHIVO WHERE " & _
                "COD_TIPO_REGISTRO NOT IN('8')AND ESTADO_PROCESO='T' " & _
                "AND FECHA_SISTEMA = to_date ('" & msk_Fecha_Proceso & "','DD/MM/YYYY')" & _
                "ORDER BY COD_TIPO_REGISTRO,FECHA_SISTEMA,HORA_PROCESO,FECHA_MOVIMIENTO"

    Set rsGrilla = cargarRecordSet(sentencia)

    If rsGrilla.EOF Then
        grillaResultadosExitosos.ClearFields
        grillaResultadosExitosos.Refresh
        grillaResultadosExitosos.Visible = False
    Else
        Set grillaResultadosExitosos.DataSource = rsGrilla
        grillaResultadosExitosos.Visible = True
        grillaResultadosExitosos.Refresh

        configurarGrilla
    End If


    Dim I As Integer
    Dim encontroRegistro7 As Boolean

    'Registros fallidos

    sentencia = "SELECT  DECODE(COD_TIPO_REGISTRO," & _
                " '1',' Entradas'," & _
                " '2',' Conciliación'," & _
                " '3',' Traductor'," & _
                " '4',' Rechazos'," & _
                " '5',' Libro Auxiliar'," & _
                " '6',' Libro Auxiliar Back Date'," & _
                " '7',' Generación de partidas sin cruce'),FECHA_SISTEMA,HORA_PROCESO,FECHA_MOVIMIENTO,CANT_REGISTROS " & _
                " FROM TBL_CONTROL_ARCHIVO WHERE COD_TIPO_REGISTRO NOT IN('8') AND ESTADO_PROCESO='I' " & _
                " AND FECHA_SISTEMA = to_date ('" & msk_Fecha_Proceso & "','DD/MM/YYYY')" & _
                " AND COD_TIPO_REGISTRO NOT IN (SELECT COD_TIPO_REGISTRO " & _
                " FROM TBL_CONTROL_ARCHIVO WHERE COD_TIPO_REGISTRO NOT IN('8') AND ESTADO_PROCESO='T' " & _
                " AND FECHA_SISTEMA = to_date ('" & msk_Fecha_Proceso & "','DD/MM/YYYY'))" & _
                " ORDER BY COD_TIPO_REGISTRO,FECHA_SISTEMA,HORA_PROCESO,FECHA_MOVIMIENTO"

    Set rsGrilla = cargarRecordSet(sentencia)
    If rsGrilla.EOF Then
        grillaResultadosFallidos.ClearFields
        grillaResultadosFallidos.Refresh
        grillaResultadosFallidos.Visible = False
    Else
        Set grillaResultadosFallidos.DataSource = rsGrilla
        grillaResultadosFallidos.Visible = True
        grillaResultadosFallidos.Refresh

        configurarGrilla2
    End If
    Me.MousePointer = 0
End Sub
Private Sub Command1_Click()
    consultarCargue
End Sub


Private Sub Command2_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub Form_Load()
    Me.msk_Fecha_Proceso = Format(Dtg_Fecha_movimiento, "DD/MM/YYYY")
    consultarCargue
End Sub


Private Sub msk_Fecha_Proceso_GotFocus()
    msk_Fecha_Proceso.SelStart = 0
    msk_Fecha_Proceso.SelLength = Len(msk_Fecha_Proceso)
End Sub

Private Sub msk_Fecha_Proceso_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        msk_Fecha_Proceso = ""
    End If
    If KeyAscii = vbKeyReturn Then
        consultarCargue
    End If
End Sub
