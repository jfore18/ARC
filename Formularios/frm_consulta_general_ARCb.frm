VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frm_Consulta_general_ARC 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Consulta Movimiento Traductor"
   ClientHeight    =   9435
   ClientLeft      =   -3525
   ClientTop       =   450
   ClientWidth     =   13995
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   9435
   ScaleWidth      =   13995
   WindowState     =   2  'Maximized
   Begin VB.ComboBox Cbo_filtro 
      Height          =   315
      ItemData        =   "frm_consulta_general_ARCb.frx":0000
      Left            =   8760
      List            =   "frm_consulta_general_ARCb.frx":000D
      TabIndex        =   68
      Text            =   "Cbo_filtro"
      Top             =   240
      Width           =   3015
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Height          =   735
      Left            =   11280
      TabIndex        =   60
      Top             =   7320
      Width           =   1215
      Begin VB.CommandButton cmd_Cierre 
         Height          =   495
         Left            =   360
         Picture         =   "frm_consulta_general_ARCb.frx":007D
         Style           =   1  'Graphical
         TabIndex        =   61
         ToolTipText     =   "Cerrar Detalles"
         Top             =   170
         Width           =   495
      End
   End
   Begin VB.Frame frm_desc 
      BackColor       =   &H00C0C0C0&
      Height          =   735
      Left            =   120
      TabIndex        =   51
      Top             =   6480
      Visible         =   0   'False
      Width           =   11655
      Begin VB.TextBox txt_desc 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Lucida Fax"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   300
         Left            =   1320
         ScrollBars      =   1  'Horizontal
         TabIndex        =   52
         Top             =   285
         Width           =   10095
      End
      Begin VB.Label Label16 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción  :"
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
         Height          =   240
         Left            =   120
         TabIndex        =   53
         Top             =   300
         Width           =   1035
      End
   End
   Begin VB.Frame frameParametros 
      Height          =   6255
      Left            =   4440
      TabIndex        =   41
      Top             =   480
      Visible         =   0   'False
      Width           =   4335
      Begin VB.ComboBox cmb_libro_auxiliar 
         Height          =   315
         Left            =   1680
         TabIndex        =   70
         Text            =   "Combo1"
         Top             =   5280
         Width           =   2295
      End
      Begin VB.ComboBox cmb_tipo_asiento 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frm_consulta_general_ARCb.frx":0947
         Left            =   1600
         List            =   "frm_consulta_general_ARCb.frx":0951
         TabIndex        =   14
         Top             =   4560
         Width           =   2535
      End
      Begin VB.CheckBox chkFlagAjuste 
         Height          =   315
         Left            =   1600
         TabIndex        =   13
         Top             =   4200
         Width           =   255
      End
      Begin VB.ComboBox cmb_aplicacion 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1600
         TabIndex        =   2
         Top             =   240
         Width           =   2535
      End
      Begin VB.ComboBox cmb_tipo_cuenta 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1600
         TabIndex        =   8
         Top             =   2400
         Width           =   2535
      End
      Begin MSMask.MaskEdBox msk_transaccion 
         Height          =   300
         Left            =   1605
         TabIndex        =   3
         Top             =   600
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   19
         Mask            =   "&&&&,&&&&,&&&&,&&&&"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox msk_destino 
         Height          =   315
         Left            =   1605
         TabIndex        =   5
         Top             =   1320
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   19
         Mask            =   "####,####,####,####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox msk_Fecha_ini 
         Height          =   315
         Left            =   1605
         TabIndex        =   6
         Top             =   1680
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox msk_cuenta 
         Height          =   315
         Left            =   1600
         TabIndex        =   9
         Top             =   2760
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   16
         Format          =   "0###############"
         Mask            =   "################"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox msk_origen 
         Height          =   315
         Left            =   1605
         TabIndex        =   4
         Top             =   960
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   19
         Mask            =   "####,####,####,####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox msk_valor_min 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   1600
         TabIndex        =   11
         Top             =   3480
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   22
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox msk_valor_max 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   1600
         TabIndex        =   12
         Top             =   3840
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   22
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox msk_secuencia 
         Height          =   315
         Left            =   1600
         TabIndex        =   10
         Top             =   3120
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         AutoTab         =   -1  'True
         HideSelection   =   0   'False
         MaxLength       =   2
         Format          =   "0#"
         Mask            =   "##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox msk_Fecha_Proceso 
         Height          =   315
         Left            =   1600
         TabIndex        =   7
         Top             =   2040
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox msk_cta_aux 
         Height          =   315
         Left            =   1605
         TabIndex        =   15
         Top             =   4920
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   11
         Format          =   "###########"
         Mask            =   "###########"
         PromptChar      =   " "
      End
      Begin VB.Label Label11 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Libro Auxiliar"
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
         Height          =   315
         Left            =   120
         TabIndex        =   69
         Top             =   5280
         Width           =   1425
      End
      Begin VB.Label Label18 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Valor máximo"
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
         Height          =   315
         Left            =   120
         TabIndex        =   59
         Top             =   3840
         Width           =   1425
      End
      Begin VB.Label Label17 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Valor minimo"
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
         Height          =   315
         Left            =   120
         TabIndex        =   58
         Top             =   3480
         Width           =   1425
      End
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Campo de valor"
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
         Height          =   315
         Left            =   120
         TabIndex        =   57
         Top             =   3120
         Width           =   1395
      End
      Begin VB.Label labelCuenta 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cuenta"
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
         Height          =   315
         Left            =   120
         TabIndex        =   56
         Top             =   2760
         Width           =   1425
      End
      Begin VB.Label LabelTipoAsiento 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo Asiento"
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
         Height          =   315
         Left            =   120
         TabIndex        =   55
         Top             =   4560
         Width           =   1425
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cuenta auxiliar"
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
         Height          =   315
         Left            =   120
         TabIndex        =   54
         Top             =   4920
         Width           =   1425
      End
      Begin VB.Label Label15 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Proceso"
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
         Height          =   315
         Left            =   120
         TabIndex        =   50
         Top             =   2040
         Width           =   1425
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Aplicación "
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
         Height          =   315
         Left            =   120
         TabIndex        =   48
         Top             =   240
         Width           =   1425
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Unidad Destino"
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
         Height          =   315
         Left            =   120
         TabIndex        =   47
         Top             =   1320
         Width           =   1425
      End
      Begin VB.Label Label9 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo Cuenta"
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
         Height          =   315
         Left            =   120
         TabIndex        =   46
         Top             =   2400
         Width           =   1425
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Unidad Origen"
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
         Height          =   315
         Left            =   120
         TabIndex        =   45
         Top             =   960
         Width           =   1425
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Transacción"
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
         Height          =   315
         Left            =   120
         TabIndex        =   44
         Top             =   600
         Width           =   1425
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Contable"
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
         Height          =   315
         Left            =   120
         TabIndex        =   43
         Top             =   1680
         Width           =   1425
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Reg. no cruzan"
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
         Height          =   315
         Left            =   120
         TabIndex        =   42
         Top             =   4200
         Width           =   1425
      End
   End
   Begin TabDlg.SSTab tb_detalle 
      Height          =   2895
      Left            =   0
      TabIndex        =   29
      Top             =   3480
      Visible         =   0   'False
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   5106
      _Version        =   393216
      Style           =   1
      Tabs            =   7
      Tab             =   3
      TabsPerRow      =   7
      TabHeight       =   617
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Transacciones Consolidadas"
      TabPicture(0)   =   "frm_consulta_general_ARCb.frx":096E
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "DBG_DETALLE"
      Tab(0).Control(1)=   "dtg_valorConc"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Transacciones Detalladas"
      TabPicture(1)   =   "frm_consulta_general_ARCb.frx":098A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "dtg_valor_det"
      Tab(1).Control(1)=   "dtg_detalle2"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Asientos COLGAAP"
      TabPicture(2)   =   "frm_consulta_general_ARCb.frx":09A6
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "dbg_conta"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Asientos IFRS"
      TabPicture(3)   =   "frm_consulta_general_ARCb.frx":09C2
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "dbg_ifrs"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Detalle Horizontal"
      TabPicture(4)   =   "frm_consulta_general_ARCb.frx":09DE
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "dtg_detalle3"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Detalle Diferencias"
      TabPicture(5)   =   "frm_consulta_general_ARCb.frx":09FA
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "dtg_detalle4"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "Configuración Conciliación"
      TabPicture(6)   =   "frm_consulta_general_ARCb.frx":0A16
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "dtg_detalle5"
      Tab(6).ControlCount=   1
      Begin MSDataGridLib.DataGrid dtg_detalle5 
         Height          =   2175
         Left            =   -74880
         TabIndex        =   71
         Top             =   480
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   3836
         _Version        =   393216
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
      Begin MSDataGridLib.DataGrid dbg_ifrs 
         Height          =   2280
         Left            =   120
         Negotiate       =   -1  'True
         TabIndex        =   39
         Top             =   480
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   4022
         _Version        =   393216
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
               LCID            =   3082
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
               LCID            =   3082
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
      Begin MSDataGridLib.DataGrid DBG_DETALLE 
         Height          =   2265
         Left            =   -74880
         Negotiate       =   -1  'True
         TabIndex        =   30
         Top             =   480
         Width           =   8715
         _ExtentX        =   15372
         _ExtentY        =   3995
         _Version        =   393216
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
               LCID            =   3082
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
               LCID            =   3082
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
      Begin MSDataGridLib.DataGrid dbg_conta 
         Height          =   2265
         Left            =   -74880
         Negotiate       =   -1  'True
         TabIndex        =   31
         Top             =   480
         Width           =   12210
         _ExtentX        =   21537
         _ExtentY        =   3995
         _Version        =   393216
         BackColor       =   16777215
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
               LCID            =   3082
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
               LCID            =   3082
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
      Begin MSDataGridLib.DataGrid dtg_valorConc 
         Height          =   2265
         Left            =   -66000
         TabIndex        =   36
         Top             =   480
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   3995
         _Version        =   393216
         ForeColor       =   0
         HeadLines       =   1
         RowHeight       =   14
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
            Name            =   "Verdana"
            Size            =   6.75
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
            MarqueeStyle    =   2
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid dtg_valor_det 
         Height          =   2265
         Left            =   -66000
         TabIndex        =   37
         Top             =   480
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   3995
         _Version        =   393216
         ForeColor       =   0
         HeadLines       =   1
         RowHeight       =   14
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
            Name            =   "Verdana"
            Size            =   6.75
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
            MarqueeStyle    =   2
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid dtg_detalle3 
         Height          =   2280
         Left            =   -74880
         Negotiate       =   -1  'True
         TabIndex        =   40
         Top             =   480
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   4022
         _Version        =   393216
         AllowUpdate     =   0   'False
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
               LCID            =   3082
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
               LCID            =   3082
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
      Begin MSDataGridLib.DataGrid dtg_detalle4 
         Height          =   2265
         Left            =   -74880
         Negotiate       =   -1  'True
         TabIndex        =   49
         Top             =   480
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   3995
         _Version        =   393216
         AllowUpdate     =   0   'False
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
               LCID            =   3082
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
               LCID            =   3082
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
      Begin MSDataGridLib.DataGrid dtg_detalle2 
         Height          =   2280
         Left            =   -74880
         Negotiate       =   -1  'True
         TabIndex        =   62
         Top             =   480
         Width           =   8715
         _ExtentX        =   15372
         _ExtentY        =   4022
         _Version        =   393216
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
               LCID            =   3082
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
               LCID            =   3082
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
   Begin VB.Frame frm_procesos 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   3773
      TabIndex        =   24
      Top             =   3165
      Visible         =   0   'False
      Width           =   4335
      Begin VB.Image Image6 
         Height          =   480
         Left            =   240
         Picture         =   "frm_consulta_general_ARCb.frx":0A32
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label25 
         BackColor       =   &H00800000&
         Caption         =   "Por favor espere, se está procesando su solicitud  ....."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   615
         Left            =   960
         TabIndex        =   25
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Height          =   750
      Left            =   4080
      TabIndex        =   20
      Top             =   7320
      Width           =   6975
      Begin VB.CommandButton Cmd_Cancelar 
         BackColor       =   &H00C0C0C0&
         Caption         =   "C&ancelar"
         Height          =   375
         Left            =   3720
         TabIndex        =   18
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmd_Imprimir 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Imprimir"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2040
         TabIndex        =   17
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton Cmd_Salir 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Salir"
         Height          =   375
         Left            =   5400
         TabIndex        =   19
         ToolTipText     =   "Salir de la Consulta"
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton Cmd_Consultar 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Consultar"
         Height          =   375
         Left            =   360
         TabIndex        =   16
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   735
      Left            =   120
      TabIndex        =   26
      Top             =   7320
      Width           =   3855
      Begin MSMask.MaskEdBox msk_cantidad 
         Height          =   255
         Left            =   2400
         TabIndex        =   28
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   -2147483637
         ForeColor       =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0"
         PromptChar      =   " "
      End
      Begin VB.Label Label26 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cantidad de Registros :"
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
         TabIndex        =   27
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame frm_imprimir 
      Caption         =   "Impresión"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   3293
      TabIndex        =   32
      Top             =   2385
      Visible         =   0   'False
      Width           =   5295
      Begin VB.CommandButton cmd_salir_imp 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Salir"
         Height          =   495
         Left            =   2880
         TabIndex        =   35
         Top             =   1920
         Width           =   1575
      End
      Begin VB.CommandButton cmd_imp 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Imprimir"
         Height          =   495
         Left            =   840
         TabIndex        =   34
         Top             =   1920
         Width           =   1575
      End
      Begin VB.ListBox lst_imprimir 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00C00000&
         Height          =   1410
         ItemData        =   "frm_consulta_general_ARCb.frx":0D3C
         Left            =   120
         List            =   "frm_consulta_general_ARCb.frx":0D43
         Style           =   1  'Checkbox
         TabIndex        =   33
         Top             =   360
         Width           =   5055
      End
   End
   Begin VB.Frame frameSuperior 
      Height          =   600
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   11820
      Begin VB.CommandButton BotonSQL 
         Caption         =   "Command1"
         Height          =   255
         Left            =   6720
         TabIndex        =   64
         Top             =   240
         Width           =   75
      End
      Begin VB.TextBox sqlText 
         Height          =   285
         Left            =   6960
         TabIndex        =   63
         Top             =   240
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton cmdParametrosConsulta 
         Caption         =   "&Parámetros de consulta"
         Height          =   315
         Left            =   4440
         TabIndex        =   1
         Top             =   200
         Width           =   2175
      End
      Begin VB.ComboBox cmb_tipo 
         Height          =   315
         ItemData        =   "frm_consulta_general_ARCb.frx":0D55
         Left            =   1800
         List            =   "frm_consulta_general_ARCb.frx":0D57
         TabIndex        =   0
         Top             =   200
         Width           =   2610
      End
      Begin VB.Label lbl_archivo 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de movimiento:"
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
         Height          =   195
         Left            =   0
         TabIndex        =   22
         Top             =   240
         Width           =   1725
      End
   End
   Begin TabDlg.SSTab sst_vlr 
      Height          =   4995
      Left            =   6480
      TabIndex        =   38
      Top             =   600
      Visible         =   0   'False
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   8811
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      ForeColor       =   16744448
      TabCaption(0)   =   "Valores"
      TabPicture(0)   =   "frm_consulta_general_ARCb.frx":0D59
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "dbg_valor"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Totales"
      TabPicture(1)   =   "frm_consulta_general_ARCb.frx":0D75
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "dtg_suma"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Errores"
      TabPicture(2)   =   "frm_consulta_general_ARCb.frx":0D91
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "dtg_error"
      Tab(2).ControlCount=   1
      Begin MSDataGridLib.DataGrid dtg_error 
         Height          =   2535
         Left            =   -74880
         TabIndex        =   67
         Top             =   480
         Visible         =   0   'False
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   4471
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
      Begin MSDataGridLib.DataGrid dbg_valor 
         Height          =   2175
         Left            =   -74880
         TabIndex        =   65
         Top             =   480
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   3836
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   14
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
            Name            =   "Verdana"
            Size            =   6.75
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
      Begin MSDataGridLib.DataGrid dtg_suma 
         Height          =   2175
         Left            =   120
         TabIndex        =   66
         Top             =   480
         Visible         =   0   'False
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   3836
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   14
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
            Name            =   "Verdana"
            Size            =   6.75
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
   Begin MSDataGridLib.DataGrid dbg_consulta 
      Height          =   5000
      Left            =   0
      TabIndex        =   23
      Top             =   600
      Visible         =   0   'False
      Width           =   6465
      _ExtentX        =   11404
      _ExtentY        =   8811
      _Version        =   393216
      AllowUpdate     =   0   'False
      ForeColor       =   0
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
         MarqueeStyle    =   2
         ScrollBars      =   3
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm_Consulta_general_ARC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim WHERE_CONSULTA As String
Dim WHERE_CONSULTA1 As String
Dim CAMPOS_CONSULTA As String
Dim FROM_CONSULTA As String
Dim sqlCamposGrilla As String
Dim ORDER_CONSULTA As String
Dim TIPO_CONSULTA As Integer
Dim cargoErrores As Boolean
Dim cargoTotales As Boolean
Dim RSOBJ3 As ADODB.Recordset
Dim tipo_configuracion As Integer
Dim data As DataGrid
Dim rsobj As ADODB.Recordset
Dim posx As Integer
Dim posy As Integer
Dim carga_cont_1 As Integer
Dim carga_cont As Integer
Dim carga_conc As Integer
Dim carga_det As Integer
Dim carga_diferencias As Integer
Dim codEntidad As String
Dim aplicacionFuente As String
Dim fechaMov As String
Dim centroOrigen As String
Dim centroDestino As String
Dim codTransaccion As String
Dim Flag_nuevo As Integer
Dim secuencia_Reg As String
Dim INL_BD_CONSULTAR As Integer
Sub configurarPaneles()

'Panel superior
    frameSuperior.Width = Screen.Width

    'Panel de parametros
    frameParametros.Top = 600

    'Panel de valores
    sst_vlr.Left = Screen.Width / 2
    sst_vlr.Width = Screen.Width / 2 - 50
    dbg_valor.Left = 100
    dbg_valor.Width = Screen.Width / 2 - 300
    dtg_error.Left = 100
    dtg_error.Width = Screen.Width / 2 - 300
    dtg_suma.Left = 100
    dtg_suma.Width = Screen.Width / 2 - 300
End Sub

'Valida que los parametros de la consulta sean correctos
Function parametrosValidos(tipoMovimiento As Integer)
    parametrosValidos = True
    Dim numeroParametros As Integer
    numeroParametros = 0


    If (Me.msk_Fecha_ini = "  /  /    ") And (Me.msk_Fecha_Proceso = "  /  /    ") Then
        MsgBox "Se requiere la Fecha Contable ó Fecha de proceso"
        parametrosValidos = False
        Exit Function
    End If


    If (Me.cmb_aplicacion.ListIndex <> -1) Then
        numeroParametros = numeroParametros + 1
    End If

    If (Me.msk_transaccion <> "    ,    ,    ,    ") Then
        numeroParametros = numeroParametros + 1
    End If

    If (Me.msk_origen <> "    ,    ,    ,    ") Then
        numeroParametros = numeroParametros + 1
    End If

    If (Me.msk_destino <> "    ,    ,    ,    ") Then
        numeroParametros = numeroParametros + 1
    End If

    If (Me.cmb_tipo_asiento.ListIndex <> -1) Then
        numeroParametros = numeroParametros + 1
    End If

    If (Me.msk_cta_aux <> "           ") Then
        numeroParametros = numeroParametros + 1
    End If

    If (Val(Trim(Me.msk_valor_max)) <> 0) Then
        numeroParametros = numeroParametros + 1
    End If

    If (Val(Trim(Me.msk_valor_min)) <> 0) Then
        numeroParametros = numeroParametros + 1
    End If

    Select Case tipoMovimiento
    Case 1:
        If Me.chkFlagAjuste.Value = 0 Then
            'If (Me.cmb_aplicacion.ListIndex = -1) Then RCC - 31012007
            If (Me.cmb_aplicacion.Text) = "" Then
                frameParametros.Visible = True
                MsgBox "Se requiere la Aplicación Fuente"
                parametrosValidos = False
                Exit Function
            End If
            If STG_NOMBRE_BD_HOST <> "PRNMC" And STG_NOMBRE_BD_HOST <> "DSNMC" Then
                If (msk_transaccion = "    ,    ,    ,    ") Then
                    frameParametros.Visible = True
                    MsgBox "Se requiere  la Transacción."
                    parametrosValidos = False
                    Exit Function
                End If
            End If
        End If
    Case 2, 3:
        If STG_NOMBRE_BD_HOST <> "PRNMC" And STG_NOMBRE_BD_HOST <> "DSNMC" Then

            If (Me.cmb_aplicacion.Text) = "" Then
                frameParametros.Visible = True
                MsgBox "Se requiere la Aplicación Fuente"
                parametrosValidos = False
                Exit Function
            End If

            If (msk_transaccion = "    ,    ,    ,    ") Then
                frameParametros.Visible = True
                frameParametros.Visible = True
                MsgBox "Se requiere la Transacción."
                parametrosValidos = False
                Exit Function
            End If
        End If

    Case 5:
        If (Me.cmb_aplicacion.ListIndex <> -1) Then
            If (Mid(Me.cmb_aplicacion.Text, 1, 4) <> "CONC") Then
                MsgBox "Solo se acepta aplicación CONC"
                parametrosValidos = False
                Exit Function
            End If

        End If

    Case 6:
        If (numeroParametros < 1) Then
            MsgBox "Se requieren al menos dos parámetros."
            parametrosValidos = False
            Exit Function
        End If
    End Select
    Exit Function
End Function
Sub PL_CERRAR_CONEXIONES()
    On Error Resume Next
    Set rsobj = Nothing
    Set rsobj2 = Nothing
    Set rsdato = Nothing
    Set rsobjp1 = Nothing
    Set rsdato2 = Nothing
    Set RSOBJ3 = Nothing
    Set OBJDATO = Nothing
    Set RSOBJp = Nothing
    Set Data2 = Nothing
    Set rsobjp1 = Nothing


End Sub

Sub PL_CERRAR_CONEXIONES1()
    On Error Resume Next

    Set dtg_valor_det.DataSource = Nothing
    Set dbg_consulta.DataSource = Nothing
    Set dtg_suma.DataSource = Nothing
    Set dtg_detalle2.DataSource = Nothing
    Set dbg_conta.DataSource = Nothing
    Set dbg_ifrs.DataSource = Nothing
    Set dtg_detalle3.DataSource = Nothing
    Set dtg_detalle4.DataSource = Nothing
    Set dtg_detalle5.DataSource = Nothing
    Set DBG_DETALLE.DataSource = Nothing
    Set dbg_valor.DataSource = Nothing
    Set dtg_valorConc.DataSource = Nothing
    Set dtg_valor_det.DataSource = Nothing


End Sub
Sub PL_CERRAR_CONEXIONES2()
    On Error Resume Next

    Set dtg_valor_det.DataSource = Nothing
    Set dtg_suma.DataSource = Nothing
    Set dtg_detalle2.DataSource = Nothing
    Set dbg_conta.DataSource = Nothing
    Set dbg_ifrs.DataSource = Nothing
    Set dtg_detalle3.DataSource = Nothing
    Set dtg_detalle4.DataSource = Nothing
    Set dtg_detalle5.DataSource = Nothing
    Set DBG_DETALLE.DataSource = Nothing
    Set dbg_valor.DataSource = Nothing
    Set dtg_valorConc.DataSource = Nothing
    Set dtg_valor_det.DataSource = Nothing


End Sub
Function filtrarTexto(texto As String, numeroSegmentos As Integer, tamanoSegmento As Integer, caracterRelleno As String) As String
    Dim I As Integer
    filtrarTexto = ""
    For I = 0 To numeroSegmentos - 1
        If (Trim(Mid(texto, tamanoSegmento * I + I + 1, tamanoSegmento)) <> "") Then
            filtrarTexto = filtrarTexto + rellenar(Trim(Mid(texto, tamanoSegmento * I + I + 1, tamanoSegmento)), tamanoSegmento, caracterRelleno, "izquierda")
        End If
    Next I
    Exit Function
End Function

Sub pl_desc_transaccion()
    Dim cod_trans As Integer

    cod_trans = Val(Me.dbg_consulta.Columns(2).Text)
End Sub
'Requerimiento CVAPD00223966.
'ABOCANE. Mayo 2015
'Modificacion cargue libro auxiliar según libro afectado (COLGAAP  IFRS)
Sub pl_carga_Conta(libro As String)

    MousePointer = 13
    Set RSOBJ3 = Nothing
    On Error Resume Next
    

    If Val(Mid(Me.cmb_tipo, 1, 2)) = 5 Or Val(Mid(Me.cmb_tipo, 1, 2)) = 3 Then

        If ING_BD_CONSULTAR = True Then
            sqlTablas = "  TBL_REGISTRO_BANA "
        Else
            sqlTablas = " USRLARC.TBL_REGISTRO_BANA@ARCHIST "
        End If

        sentencia = "SELECT  APLICACION_FUENTE,COD_TRANSACCION,FECHA_CONTABLE,FECHA_SISTEMA,CENTRO_ORIGEN,CENTRO_DESTINO,COD_TIPO_ASIENTO,CTA_AUXILIAR,CTA_FMS,VALOR_ASIENTO,ID_CAMPO" & _
                  " FROM " & sqlTablas & " WHERE  COD_ENTIDAD = 1" & _
                  " AND LIBRO_AUXILIAR = '" & libro & "'" & _
                  " AND APLICACION_FUENTE = '" & dbg_consulta.Columns(0) & "'" & _
                  " AND COD_TRANSACCION = '" & Format(dbg_consulta.Columns(1), "0###") & "'" & _
                  " AND  FECHA_CONTABLE = TO_DATE('" & Format(dbg_consulta.Columns(2), "DD/MM/YYYY") & "','DD/MM/YYYY') " & _
                  " AND ((CENTRO_ORIGEN = " & Val(dbg_consulta.Columns(4)) & " AND CENTRO_DESTINO = " & Val(dbg_consulta.Columns(5)) & ") OR " & _
                  " (CENTRO_DESTINO  = " & Val(dbg_consulta.Columns(4)) & " AND CENTRO_ORIGEN  = " & Val(dbg_consulta.Columns(5)) & "))"

        'MsgBox sentencia
        Set RSOBJ3 = cargarRecordSet(sentencia)



        tipo_configuracion = 8
    End If
    If Val(Mid(Me.cmb_tipo, 1, 2)) = 2 Then
        If ING_BD_CONSULTAR = True Then
            sqlTablas = "  TBL_REGISTRO_BANA "
        Else
            sqlTablas = " USRLARC.TBL_REGISTRO_BANA@ARCHIST "
        End If

        sentencia = "SELECT  APLICACION_FUENTE,COD_TRANSACCION,FECHA_CONTABLE,FECHA_SISTEMA,CENTRO_ORIGEN,CENTRO_DESTINO,COD_TIPO_ASIENTO,CTA_AUXILIAR,CTA_FMS,VALOR_ASIENTO,ID_CAMPO" & _
                  " FROM " & sqlTablas & "  WHERE  COD_ENTIDAD = 1" & _
                  " AND LIBRO_AUXILIAR = '" & libro & "'" & _
                  " AND APLICACION_FUENTE = '" & dbg_consulta.Columns(0) & "'" & _
                  " AND COD_TRANSACCION = '" & Format(dbg_consulta.Columns(1), "0###") & "'" & _
                  " AND  FECHA_CONTABLE = TO_DATE('" & Format(dbg_consulta.Columns(4), "DD/MM/YYYY") & "','DD/MM/YYYY') " & _
                  " AND ((CENTRO_ORIGEN = " & Val(dbg_consulta.Columns(6)) & " AND CENTRO_DESTINO = " & Val(dbg_consulta.Columns(7)) & ") OR " & _
                  " (CENTRO_DESTINO  = " & Val(dbg_consulta.Columns(6)) & " AND CENTRO_ORIGEN  = " & Val(dbg_consulta.Columns(7)) & "))"

        'MsgBox sentencia
        Set RSOBJ3 = cargarRecordSet(sentencia)
        tipo_configuracion = 8
    End If

    If RSOBJ3.EOF Then
        
        If libro = 1 Then
            MsgBox "No existe información para COLGAAP "
            carga_cont = 3
            Me.tb_detalle.Tab = 0
        Else
             MsgBox "No existe información para IFRS"
             Me.tb_detalle.Tab = 0
        End If
    Else
         If libro = 1 Then
             Set Me.dbg_conta.DataSource = RSOBJ3
            Configurar_grilla_Dat_consulta Me.dbg_conta
            carga_cont = 1
        Else
            Set Me.dbg_ifrs.DataSource = RSOBJ3
            Configurar_grilla_Dat_consulta Me.dbg_ifrs
            carga_cont_1 = 1
        End If
    End If
    MousePointer = 0
End Sub


Sub pl_tipo_consulta()
    Dim tipo_archivo As Integer
    tipo_archivo = Val(Mid(Me.cmb_tipo, 1, 2))

    'Valores
    Label17.Top = 3480
    msk_valor_min.Top = 3480
    Label18.Top = 3840
    msk_valor_max.Top = 3840

    'Tipo Cuenta
    Label9.Visible = True
    cmb_tipo_cuenta.Visible = True

    'Cuenta
    labelCuenta.Visible = True
    msk_cuenta.Visible = True

    'Codigo campo
    Label6.Visible = True
    msk_secuencia.Visible = True

    'Flag de No-Cruce
    Label8.Visible = True
    chkFlagAjuste.Visible = True
    
    'Combo Libro Auxiliar
    'REQ CVAPD00223966 ABOCANE Mayo 2016
    Label11.Visible = False
    cmb_libro_auxiliar.Visible = False

    Me.cmb_aplicacion.Enabled = True
    'cmb_aplicacion.ListIndex = -1
    Select Case tipo_archivo

    Case 1

        'Reg. no cruzan
        Label8.Visible = True
        chkFlagAjuste.Visible = True

        'Tipo Asiento
        LabelTipoAsiento.Visible = False
        cmb_tipo_asiento.Visible = False

        'Cuenta auxiliar
        Label3.Visible = False
        msk_cta_aux.Visible = False

        'Altura del cuadro de parametros
        Me.frameParametros.Height = 4695
    Case 2

        'Reg. no cruzan
        Label8.Visible = False
        chkFlagAjuste.Visible = False

        'Tipo Asiento
        LabelTipoAsiento.Visible = False
        cmb_tipo_asiento.Visible = False

        'Cuenta auxiliar
        Label3.Visible = False
        msk_cta_aux.Visible = False

       'Altura del cuadro de parametros
        Me.frameParametros.Height = 4335
    Case 3


        'Reg. no cruzan
        Label8.Visible = True
        chkFlagAjuste.Visible = True

        'Tipo Asiento
        LabelTipoAsiento.Visible = False
        cmb_tipo_asiento.Visible = False

        'Cuenta auxiliar
        Label3.Visible = False
        msk_cta_aux.Visible = False

       'Altura del cuadro de parametros
        Me.frameParametros.Height = 4695
    Case 4


        'Reg. no cruzan
        Label8.Visible = True
        chkFlagAjuste.Visible = True

        'Tipo Asiento
        LabelTipoAsiento.Visible = False
        cmb_tipo_asiento.Visible = False

        'Cuenta auxiliar
        Label3.Visible = False
        msk_cta_aux.Visible = False

       'Altura del cuadro de parametros
        Me.frameParametros.Height = 4695

    Case 5
        'Filtro
        'Me.chkFiltro.Value = 1
        'Me.chkFiltro.Enabled = False

        cmb_aplicacion.Text = "CONC"


        'Reg. no cruzan
        Label8.Visible = False
        chkFlagAjuste.Visible = False

        'Tipo Asiento
        LabelTipoAsiento.Visible = False
        cmb_tipo_asiento.Visible = False

        'Cuenta auxiliar
        Label3.Visible = False
        msk_cta_aux.Visible = False

       'Altura del cuadro de parametros
        Me.frameParametros.Height = 4335

        Dim I As Integer
        For I = 0 To Me.cmb_aplicacion.ListCount - 1
            If (Mid(cmb_aplicacion.List(I), 1, 4) = "CONC") Then
                cmb_aplicacion.ListIndex = I
                Exit For
            End If
        Next I
        cmb_aplicacion.Enabled = False
    Case 6

        'Tipo Asiento
        LabelTipoAsiento.Visible = True
        cmb_tipo_asiento.Visible = True
        LabelTipoAsiento.Top = 2400
        cmb_tipo_asiento.Top = 2400


        'Cuenta auxiliar
        Label3.Visible = True
        msk_cta_aux.Visible = True
        Label3.Top = 2760
        msk_cta_aux.Top = 2760

             
        'Combo Libro Auxiliar
        'REQ CVAPD00223966 ABOCANE Mayo 2016
         Label11.Visible = True
         cmb_libro_auxiliar.Visible = True
         Label11.Top = 3900
         cmb_libro_auxiliar.Top = 3900
        
        'Valores
        Label17.Top = 3100
        msk_valor_min.Top = 3100
        Label18.Top = 3500
        msk_valor_max.Top = 3500


        'Tipo Cuenta
        Label9.Visible = False
        cmb_tipo_cuenta.Visible = False

        'Cuenta
        labelCuenta.Visible = False
        msk_cuenta.Visible = False

        'Codigo campo
        Label6.Visible = False
        msk_secuencia.Visible = False

        'CRUZA
        Label8.Visible = False
        chkFlagAjuste.Visible = False

        'Altura del cuadro de parametros
        Me.frameParametros.Height = 4400
    End Select
End Sub

Function Valida_fecha_consulta(fecha As String, camposet As MaskEdBox, TIPO As Integer) As Boolean

    On Error GoTo error

    If fecha <> "  /  /    " Then
        If IsDate(fecha) Then
            If TIPO = 1 Then
                If Dtg_fecha_sistema >= fecha And fecha >= msk_Fecha_ini.Text Then
                Else
                    MsgBox "Fecha Invalida"
                    Valida_fecha_consulta = False
                    camposet.Text = "  /  /    "
                    camposet.SetFocus
                End If
            Else
                Valida_fecha_consulta = True
            End If

        Else
            MsgBox "Fecha Invalida"
            Valida_fecha_consulta = False
            camposet.Text = "  /  /    "
            camposet.SetFocus
        End If
    End If

    Exit Function
error:
    MsgBox Err.Number & " " & Err.Description

End Function
Sub pl_carga_grilla_vacia()

    Set rsobj = Nothing
    Set Me.dbg_consulta.DataSource = rsobj

End Sub

Sub Configurar_grilla_Dat_consulta(DATAG As DataGrid)

    Dim I As Integer
    Dim indice As Integer
    Dim desde As Integer
    Dim hasta As Integer

    desde = 13
    hasta = 72
    indice = 1
    On Error Resume Next

    Me.dbg_consulta.Width = Screen.Width / 2
    Me.dbg_consulta.ScrollBars = dbgAutomatic

    DATAG.Row = 0

    'Aplica para: Transacciones por Normalizar
    If tipo_configuracion = 1 Then
        DATAG.Columns(0).Caption = "Aplicación"
        DATAG.Columns(1).Caption = "Transacción"
        DATAG.Columns(2).Caption = "F.Contable"
        DATAG.Columns(3).Caption = "F.Proceso"
        DATAG.Columns(4).Caption = "Origen"
        DATAG.Columns(5).Caption = "Destino"
        DATAG.Columns(6).Caption = "Tipo Cta"
        DATAG.Columns(7).Caption = "Cuenta"
        DATAG.Columns(8).Caption = "Descripción"
        DATAG.Columns(9).Caption = "Tipo registro"
        DATAG.Columns(10).Caption = "Secuencia"
        'Configuracion formatos
        DATAG.Columns(2).NumberFormat = ("DD/MM/YYYY")
        DATAG.Columns(3).NumberFormat = ("DD/MM/YYYY")
        DATAG.Columns(4).NumberFormat = ("0###")
        DATAG.Columns(5).NumberFormat = ("0###")
        DATAG.Columns(7).NumberFormat = ("0###############")

        'Bloquea y vuelve visible todas las columnas
        For I = 0 To 10
            DATAG.Columns(I).Visible = True
            DATAG.Columns(I).Locked = True
        Next I

        'Oculta ciertas columnas
        DATAG.Columns(9).Visible = False
        DATAG.Columns(10).Visible = False

        'Configuracion ancho de las Columnas
        DATAG.Columns(0).Width = 1000
        DATAG.Columns(1).Width = 1200
        DATAG.Columns(2).Width = 1000
        DATAG.Columns(3).Width = 1000
        DATAG.Columns(4).Width = 650
        DATAG.Columns(5).Width = 700
        DATAG.Columns(6).Width = 800
        DATAG.Columns(7).Width = 1700
        DATAG.Columns(8).Width = 9000
        DATAG.Columns(9).Width = 1000
        DATAG.Columns(10).Width = 1000

        'Oculta el boton para el error
        DATAG.Columns(1).Button = False
    End If
    
    If tipo_configuracion = 3 Then
        DATAG.Columns(0).Caption = "Aplicación"
        DATAG.Columns(1).Caption = "Transacción"
        DATAG.Columns(2).Caption = "F.Contable"
        DATAG.Columns(3).Caption = "F.Proceso"
        DATAG.Columns(4).Caption = "Origen"
        DATAG.Columns(5).Caption = "Destino"
        DATAG.Columns(6).Caption = "Tipo Cta"
        DATAG.Columns(7).Caption = "Cuenta"
        DATAG.Columns(8).Caption = "Descripción"
        DATAG.Columns(9).Caption = "Tipo registro"
        DATAG.Columns(10).Caption = "Secuencia"
        'Configuracion formatos
        DATAG.Columns(2).NumberFormat = ("DD/MM/YYYY")
        DATAG.Columns(3).NumberFormat = ("DD/MM/YYYY")
        DATAG.Columns(4).NumberFormat = ("0###")
        DATAG.Columns(5).NumberFormat = ("0###")
        DATAG.Columns(7).NumberFormat = ("0###############")

        'Bloquea y vuelve visible todas las columnas
        For I = 0 To 10
            DATAG.Columns(I).Visible = True
            DATAG.Columns(I).Locked = True
        Next I

        'Oculta ciertas columnas
        DATAG.Columns(9).Visible = False
        DATAG.Columns(10).Visible = False

        'Configuracion ancho de las Columnas
        DATAG.Columns(0).Width = 1000
        DATAG.Columns(1).Width = 1300
        DATAG.Columns(2).Width = 1000
        DATAG.Columns(3).Width = 1000
        DATAG.Columns(4).Width = 650
        DATAG.Columns(5).Width = 750
        DATAG.Columns(6).Width = 500
        DATAG.Columns(7).Width = 1700
        DATAG.Columns(8).Width = 9800
        DATAG.Columns(9).Width = 1000
        DATAG.Columns(10).Width = 1000

        'Oculta el boton para el error
        DATAG.Columns(1).Button = False
    End If
    
    'Aplica para: Detalle de INCONSISTENTES
    If tipo_configuracion = 5 Then
        DATAG.Columns(1).Caption = "Aplicación"
        DATAG.Columns(2).Caption = "Transacción"
        DATAG.Columns(3).Caption = "F.Contable"
        DATAG.Columns(12).Caption = "F.Proceso"
        DATAG.Columns(4).Caption = "Origen"
        DATAG.Columns(5).Caption = "Destino"
        DATAG.Columns(6).Caption = "Tipo Cta"
        DATAG.Columns(7).Caption = "Cuenta"
        DATAG.Columns(8).Caption = "Num Documento"
        DATAG.Columns(13).Caption = "Secuencia"
        DATAG.Columns(14).Caption = "Tipo registro"
        DATAG.Columns(15).Caption = "Descripcion"
        'Configuracion formatos
        DATAG.Columns(3).NumberFormat = ("DD/MM/YYYY")
        DATAG.Columns(12).NumberFormat = ("DD/MM/YYYY")
        DATAG.Columns(4).NumberFormat = ("0###")
        DATAG.Columns(5).NumberFormat = ("0###")
        DATAG.Columns(7).NumberFormat = ("0###############")

        'Bloquea y vuelve visible todas las columnas
        For I = 0 To 15
            DATAG.Columns(I).Visible = True
            DATAG.Columns(I).Locked = True
        Next I

        'Oculta ciertas columnas
        DATAG.Columns(0).Visible = False
        DATAG.Columns(9).Visible = False
        DATAG.Columns(10).Visible = False
        DATAG.Columns(11).Visible = False

        'Configuracion ancho de las Columnas
        DATAG.Columns(0).Width = 1000
        DATAG.Columns(1).Width = 1100
        DATAG.Columns(2).Width = 1100
        DATAG.Columns(3).Width = 1000
        DATAG.Columns(4).Width = 680
        DATAG.Columns(5).Width = 700
        DATAG.Columns(6).Width = 600
        DATAG.Columns(7).Width = 1700
        DATAG.Columns(8).Width = 1700
        DATAG.Columns(9).Width = 1000
        DATAG.Columns(10).Width = 1000
        DATAG.Columns(15).Width = 4000

        'Oculta el boton para el error
        DATAG.Columns(1).Button = False
    End If


    'Aplica para Conciliaciones
    If tipo_configuracion = 2 Then
        On Error Resume Next
        'Asigna nombres
        DATAG.Columns(0).Caption = "Aplicación Conc."
        DATAG.Columns(1).Caption = "Transacción Conc."
        DATAG.Columns(2).Caption = "Fuente"
        DATAG.Columns(3).Caption = "Transacción"
        DATAG.Columns(4).Caption = "Fecha Contable"
        DATAG.Columns(5).Caption = "Fecha Proceso"
        DATAG.Columns(6).Caption = "Origen"
        DATAG.Columns(7).Caption = "Destino"
        DATAG.Columns(8).Caption = "Tipo Cta"
        DATAG.Columns(9).Caption = "Cuenta"
        DATAG.Columns(10).Caption = "Secuencia"
        DATAG.Columns(11).Caption = "Descripción"
        indice = 1

        For I = 0 To 11
            DATAG.Columns(I).Visible = True
            DATAG.Columns(I).Locked = True
        Next I

        DATAG.Columns(10).Visible = False

        'Ancho de Las Columnas
        DATAG.Columns(0).Width = 1500
        DATAG.Columns(1).Width = 1650
        DATAG.Columns(2).Width = 750
        DATAG.Columns(3).Width = 1300
        DATAG.Columns(4).Width = 1500
        DATAG.Columns(5).Width = 1500
        DATAG.Columns(6).Width = 680
        DATAG.Columns(7).Width = 750
        DATAG.Columns(8).Width = 850
        DATAG.Columns(9).Width = 1600
        DATAG.Columns(10).Width = 900
        DATAG.Columns(11).Width = 10000


        ' DA FORMATO A LAS COLUMNAS

        DATAG.Columns(4).NumberFormat = ("DD/MM/YYYY")
        DATAG.Columns(5).NumberFormat = ("DD/MM/YYYY")
        DATAG.Columns(6).NumberFormat = ("0###")
        DATAG.Columns(7).NumberFormat = ("0###")
        DATAG.Columns(8).NumberFormat = ("0#")
        DATAG.Columns(9).NumberFormat = ("0###############")

        tipo_configuracion = 0

        'Oculta el boton para el error
        DATAG.Columns(1).Button = False
        Exit Sub
    End If

    If tipo_configuracion = 4 Then
        DATAG.Columns(0).Caption = "Aplicación"
        DATAG.Columns(1).Caption = "Transacción"
        DATAG.Columns(2).Caption = "F.Contable"
        DATAG.Columns(3).Caption = "F. Proceso"
        DATAG.Columns(4).Caption = "Origen"
        DATAG.Columns(5).Caption = "Destino"
        DATAG.Columns(6).Caption = "Tipo Cta"
        DATAG.Columns(7).Caption = "Cuenta"
        DATAG.Columns(8).Caption = "Descripción"
        DATAG.Columns(9).Caption = "Tipo registro"
        DATAG.Columns(10).Caption = "Secuencia"
        'DATAG.Columns(11).Caption = "Error"
        'DATAG.Columns(12).Caption = "Acción correctiva"
        'Configuracion formatos
        DATAG.Columns(2).NumberFormat = ("DD/MM/YYYY")
        DATAG.Columns(3).NumberFormat = ("DD/MM/YYYY")
        DATAG.Columns(4).NumberFormat = ("0###")
        DATAG.Columns(5).NumberFormat = ("0###")
        DATAG.Columns(7).NumberFormat = ("0###############")
        
        'Bloquea y vuelve visible todas las columnas
        For I = 0 To DATAG.Columns.Count - 1
            DATAG.Columns(I).Visible = True
            DATAG.Columns(I).Locked = True
        Next I

        'Oculta ciertas columnas
        DATAG.Columns(9).Visible = False
        DATAG.Columns(10).Visible = False

        'Configuracion ancho de las Columnas
        DATAG.Columns(0).Width = 1000
        DATAG.Columns(1).Width = 1200
        DATAG.Columns(2).Width = 1000
        DATAG.Columns(3).Width = 1000
        DATAG.Columns(4).Width = 650
        DATAG.Columns(5).Width = 750
        DATAG.Columns(6).Width = 970
        DATAG.Columns(7).Width = 1700
        DATAG.Columns(8).Width = 8000


        'Muestra el boton para el error
        DATAG.Columns(1).Button = True
    End If
    
    'Aplica para Movimiento Auxiliar
    If tipo_configuracion = 6 Then
        'Configuracion grilla
        DATAG.Width = 11800
        'Configuracion etiquetas
        DATAG.Columns(0).Caption = "Aplicación"
        DATAG.Columns(1).Caption = "Transacción"
        DATAG.Columns(2).Caption = "Fecha Contable"
        DATAG.Columns(3).Caption = "Fecha Proceso"
        DATAG.Columns(4).Caption = "Origen "
        DATAG.Columns(5).Caption = "Destino "
        DATAG.Columns(6).Caption = "Asiento"
        DATAG.Columns(7).Caption = "Campo"
        DATAG.Columns(8).Caption = "Cuenta Auxiliar"
        DATAG.Columns(9).Caption = "Valor Asiento"
        
        'Configuracion formatos
        DATAG.Columns(2).NumberFormat = ("DD/MM/YYYY")
        DATAG.Columns(3).NumberFormat = ("DD/MM/YYYY")
        DATAG.Columns(4).NumberFormat = ("0###")
        DATAG.Columns(5).NumberFormat = ("0###")
        DATAG.Columns(9).NumberFormat = ("###,###,###,###.00")
        'Configuracion alineacion
        DATAG.Columns(9).Alignment = dbgRight
        'Configuracion bloqueo y visibilidad
        For I = 0 To 9
            DATAG.Columns(I).Visible = True
            DATAG.Columns(I).Locked = True
        Next I
        'Configuracion ancho de las Columnas
        DATAG.Columns(0).Width = 1000
        DATAG.Columns(1).Width = 1150
        DATAG.Columns(2).Width = 1420
        DATAG.Columns(3).Width = 1350
        DATAG.Columns(4).Width = 650
        DATAG.Columns(5).Width = 750
        DATAG.Columns(6).Width = 800
        DATAG.Columns(7).Width = 730
        DATAG.Columns(8).Width = 1450
        DATAG.Columns(9).Width = 1800
     
        'Oculta el boton para el error
        DATAG.Columns(1).Button = False
        Exit Sub
    End If
    
    'Aplica para el tab_detalle cuando se hace clic en Asientos Contables
    If tipo_configuracion = 8 Then
        'Configuracion etiquetas
        DATAG.Columns(0).Caption = "Aplicación"
        DATAG.Columns(1).Caption = "Transacción"
        DATAG.Columns(2).Caption = "Fecha Contable"
        DATAG.Columns(3).Caption = "Fecha Proceso"
        DATAG.Columns(4).Caption = "Origen"
        DATAG.Columns(5).Caption = "Destino"
        DATAG.Columns(6).Caption = "Tipo Asiento"
        DATAG.Columns(7).Caption = "Cuenta Auxiliar"
        DATAG.Columns(8).Caption = "Cuenta FMS"
        DATAG.Columns(9).Caption = "Valor Asiento"
        DATAG.Columns(10).Caption = "Campo"
        'Configuracion formatos
        DATAG.Columns(2).NumberFormat = ("DD/MM/YYYY")
        DATAG.Columns(3).NumberFormat = ("DD/MM/YYYY")
        DATAG.Columns(4).NumberFormat = ("0###")
        DATAG.Columns(5).NumberFormat = ("0###")
        DATAG.Columns(9).NumberFormat = ("###,###,###,###.00")
        'Configuracion alineacion
        DATAG.Columns(9).Alignment = dbgRight
        'Configuracion bloqueo y visibilidad
        For I = 0 To 10
            DATAG.Columns(I).Visible = True
            DATAG.Columns(I).Locked = True
        Next I
        'Configuracion ancho de las Columnas
        DATAG.Columns(0).Width = 820
        DATAG.Columns(1).Width = 970
        DATAG.Columns(2).Width = 1210
        DATAG.Columns(3).Width = 1160
        DATAG.Columns(4).Width = 600
        DATAG.Columns(5).Width = 620
        DATAG.Columns(6).Width = 960
        DATAG.Columns(7).Width = 1140
        DATAG.Columns(8).Width = 1000
        DATAG.Columns(9).Width = 1800
        DATAG.Columns(10).Width = 590
        DATAG.Columns(8).Visible = False
        
        
        Exit Sub
    End If

    'Aplica para detalle de Transacciones por codificar y tipoSeleccionado = 9 (no se sabe aun para que es ese tipo)
    If tipo_configuracion = 11 Then
        'Asigna nombres

        DATAG.Columns(0).Caption = "Entidad"
        DATAG.Columns(1).Caption = "Aplicación"
        DATAG.Columns(2).Caption = "Transacción"
        DATAG.Columns(3).Caption = "Fecha Contable"
        DATAG.Columns(4).Caption = "Origen"
        DATAG.Columns(5).Caption = "Destino"
        DATAG.Columns(6).Caption = "Tipo Cuenta"
        DATAG.Columns(7).Caption = "Cuenta"
        DATAG.Columns(8).Caption = "Documento"
        DATAG.Columns(9).Caption = "Filler"
        DATAG.Columns(10).Caption = "Numero de campos"
        DATAG.Columns(11).Caption = "Tamaño de los campos"
        DATAG.Columns(12).Caption = "Fecha Proceso"
        DATAG.Columns(13).Caption = "Secuencia"
        DATAG.Columns(14).Caption = "Descripción"

        DATAG.Columns(1).Width = 850
        DATAG.Columns(2).Width = 1000
        DATAG.Columns(3).Width = 1250
        DATAG.Columns(4).Width = 650
        DATAG.Columns(5).Width = 700
        DATAG.Columns(6).Width = 1000
        DATAG.Columns(7).Width = 1550
        DATAG.Columns(10).Width = 1500
        DATAG.Columns(11).Width = 1800
        DATAG.Columns(12).Width = 1200
        DATAG.Columns(14).Width = 8000


        'Configuracion bloqueo y visibilidad
        For I = 1 To 14
            DATAG.Columns(I).Visible = True
            DATAG.Columns(I).Locked = True
        Next I
        DATAG.Columns(0).Visible = False
        DATAG.Columns(8).Visible = False
        DATAG.Columns(9).Visible = False
        DATAG.Columns(13).Visible = False


    End If

    'Aplica para el tab_detalle cuando se hace clic en Transacciones Detalladas
    If tipo_configuracion = 12 Then
        'Asigna nombres

        DATAG.Columns(0).Caption = "Entidad"
        DATAG.Columns(1).Caption = "Aplicación"
        DATAG.Columns(2).Caption = "Transacción"
        DATAG.Columns(3).Caption = "Fecha Contable"
        DATAG.Columns(4).Caption = "Origen"
        DATAG.Columns(5).Caption = "Destino"
        DATAG.Columns(6).Caption = "Tipo Cuenta"
        DATAG.Columns(7).Caption = "Cuenta"
        DATAG.Columns(8).Caption = "Documento"
        DATAG.Columns(9).Caption = "Filler"
        DATAG.Columns(10).Caption = "Numero de campos"
        DATAG.Columns(11).Caption = "Tamaño de los campos"
        DATAG.Columns(12).Caption = "Tipo de registro"
        DATAG.Columns(13).Caption = "Fecha Proceso"
        DATAG.Columns(14).Caption = "Secuencia"
        DATAG.Columns(15).Caption = "Descripción"


        DATAG.Columns(1).Width = 850
        DATAG.Columns(2).Width = 1000
        DATAG.Columns(3).Width = 1200
        DATAG.Columns(4).Width = 700
        DATAG.Columns(5).Width = 700
        DATAG.Columns(6).Width = 1000
        DATAG.Columns(7).Width = 1500
        DATAG.Columns(10).Width = 1500
        DATAG.Columns(11).Width = 1800
        DATAG.Columns(13).Width = 1200
        DATAG.Columns(15).Width = 8000


        'Configuracion bloqueo y visibilidad
        For I = 1 To 15
            DATAG.Columns(I).Visible = True
            DATAG.Columns(I).Locked = True
        Next I
        DATAG.Columns(0).Visible = False
        DATAG.Columns(8).Visible = False
        DATAG.Columns(9).Visible = False
        DATAG.Columns(12).Visible = False
        DATAG.Columns(14).Visible = False
    End If

    'Aplica para Detalle Horizontal y Detalle Diferencias del tb_detalle
    If tipo_configuracion = 7 Then
        'Asigna nombres
        DATAG.Columns(0).Caption = "Entidad"
        DATAG.Columns(1).Caption = "Aplicación"
        DATAG.Columns(2).Caption = "Transacción"
        DATAG.Columns(3).Caption = "Fecha Contable"
        DATAG.Columns(4).Caption = "Origen"
        DATAG.Columns(5).Caption = "Destino"
        DATAG.Columns(6).Caption = "Tipo Cuenta"
        DATAG.Columns(7).Caption = "Cuenta"
        DATAG.Columns(8).Caption = "Referencia"
        DATAG.Columns(9).Caption = "Filler"
        DATAG.Columns(10).Caption = "Número de campos"
        DATAG.Columns(11).Caption = "Tamaño de los campos"
        DATAG.Columns(12).Caption = "Tipo de registro"

        DATAG.Columns(0).Visible = False
        DATAG.Columns(8).Visible = False
        DATAG.Columns(9).Visible = False
        DATAG.Columns(12).Visible = False

        I = 13
        While I <= 32

            DATAG.Columns(I).NumberFormat = ("0#")
            DATAG.Columns(I).Width = 1200
            DATAG.Columns(I + 1).NumberFormat = ("###,###,###,###.00")
            DATAG.Columns(I + 1).Width = 1500
            DATAG.Columns(I + 1).Alignment = dbgRight
            I = I + 2
        Wend

        DATAG.Columns(0).Width = 400
        DATAG.Columns(1).Width = 850
        DATAG.Columns(2).Width = 1000
        DATAG.Columns(3).Width = 1200
        DATAG.Columns(4).Width = 700
        DATAG.Columns(5).Width = 700
        DATAG.Columns(6).Width = 1000
        DATAG.Columns(7).Width = 1500
        DATAG.Columns(8).Width = 2000
        DATAG.Columns(9).Width = 50
        DATAG.Columns(10).Width = 1500
        DATAG.Columns(11).Width = 1800

    End If
    
    'Aplica para Configuración Conciliación del tb_detalle
    If tipo_configuracion = 20 Then
        'Asigna nombres

        DATAG.Columns(0).Caption = "Aplicación conciliación"
        DATAG.Columns(1).Caption = "Transacción conciliación"
        DATAG.Columns(2).Caption = "Secuencia"
        DATAG.Columns(3).Caption = "Aplicación"
        DATAG.Columns(4).Caption = "Transacción"
        DATAG.Columns(5).Caption = "Identificador"
        DATAG.Columns(6).Caption = "Aplicación"
        DATAG.Columns(7).Caption = "Transacción"
        DATAG.Columns(8).Caption = "Identificador"
        DATAG.Columns(9).Caption = "Valor aplicado 1"
        DATAG.Columns(10).Caption = "Valor aplicado 2"
        DATAG.Columns(11).Caption = "Valor diferencia 1"
        DATAG.Columns(12).Caption = "Valor diferencia 2"

        DATAG.Columns(0).Width = 850
        DATAG.Columns(1).Width = 1000
        DATAG.Columns(2).Width = 900
        DATAG.Columns(3).Width = 850
        DATAG.Columns(4).Width = 1000
        DATAG.Columns(5).Width = 1000
        DATAG.Columns(6).Width = 850
        DATAG.Columns(7).Width = 1000
        DATAG.Columns(8).Width = 1000
        DATAG.Columns(9).Width = 1250
        DATAG.Columns(10).Width = 1250
        DATAG.Columns(11).Width = 1300
        DATAG.Columns(12).Width = 1300


        'Configuracion bloqueo y visibilidad
        For I = 0 To 12
            DATAG.Columns(I).Visible = True
            DATAG.Columns(I).Locked = True
        Next I

    End If

End Sub




Function PL_Carga_Grilla(tipoMovimiento As Integer, campos As String, TABLAS As String, CONDICIONES As String)
    On Error GoTo error
    'Dependiento del tipo de movimiento realiza la consulta y ordena los datos
    Select Case Val(Mid(Me.cmb_tipo, 1, 2))
    Case 1, 3, 4, 5:
        sentencia = "SELECT  " & campos & " FROM " & TABLAS & " WHERE " & CONDICIONES & " GROUP BY " & campos & " ORDER BY " & campos
        'MsgBox sentencia
        Set rsobj = cargarRecordSet(sentencia)

    Case 2:
        sentencia = "SELECT  " & campos & " FROM " & TABLAS & " WHERE " & CONDICIONES & " "
        'MsgBox sentencia
        Set rsobj = cargarRecordSet(sentencia)
    Case 6:
        sqlText.Text = "SELECT  " & campos & " FROM " & TABLAS & " WHERE " & CONDICIONES & " ORDER BY APLICACION_FUENTE,COD_TRANSACCION,FECHA_CONTABLE,FECHA_SISTEMA,CENTRO_ORIGEN,CENTRO_DESTINO"
        Set rsobj = cargarRecordSet("SELECT  " & campos & " FROM " & TABLAS & " WHERE " & CONDICIONES & " ORDER BY APLICACION_FUENTE,COD_TRANSACCION,FECHA_CONTABLE,FECHA_SISTEMA,CENTRO_ORIGEN,CENTRO_DESTINO")
    End Select

    'Si no hay movimiento genera el mensaje
    If rsobj.EOF Then
        sst_vlr.Visible = False
        dbg_consulta.ClearFields
        dbg_consulta.Refresh
        'If Me.chkFiltro.Value = 0 Then
        If Me.Cbo_filtro.ListIndex = 0 Then
            MsgBox "No existe información que cumpla los criterios de búsqueda, o su asignación de responsabilidades no le permite tener acceso a esta información. Revise los parámetros de búsqueda o su asignación de responsabilidades.", vbInformation
        Else
            MsgBox "No existe información que cumpla los criterios de búsqueda.  Revise los parámetros de búsqueda.", vbInformation
        End If

        Me.dbg_consulta.Height = 500
        Me.dbg_valor.Height = 4200
        Me.dtg_suma.Height = 4200
        Me.dtg_error.Height = 4200
        sst_vlr.Height = 5000
        Me.cmb_tipo.SetFocus
        'Si hay movimiento se muestra en la grilla
    Else
        'El movimiento tipo 6 (BANA) no tiene campos de valor
        If Val(Mid(Me.cmb_tipo, 1, 2)) = 6 Then
            sst_vlr.Visible = False
        Else
            sst_vlr.Visible = True
        End If

        'Muestra la grilla con el resultado de la consulta
        dbg_consulta.Visible = True
        Set dbg_consulta.DataSource = rsobj

        'Oculta el panel de parametros
        frameParametros.Visible = False
        'Activa el boton de impresión
        cmd_Imprimir.Enabled = True
        'Actualiza la caja de texto que muestra el número de registros en la consulta
        Me.msk_cantidad.Text = rsobj.RecordCount

        'Configura la grilla
        Configurar_grilla_Dat_consulta Me.dbg_consulta
        Set rsobj = Nothing
        'Me.sst_vlr.Enabled = False

        dbg_consulta_Click
    End If



    Exit Function

error:
    MsgBox Err.Number & " - " & Err.Description & " llegó al error"


End Function
'Funcion que carga los tipos de movimiento a los que va tener acceso el usuario
Sub pl_cargar_tipo()
    PL_Conexion_Oracle
    'si el perfil es consulta parcial, filtramos los movimientos usando la tabla TBL_USUARIO_MOVIMIENTO
    If Stg_Perfil_Usuario_Acceso = 8 Then
        sentencia = "SELECT A.COD_TIPO_REGISTRO CODIGO, A.DESC_TIPO_REGISTRO DESCRIPCION FROM TBL_TIPO_REGISTRO A, TBL_USUARIO_MOVIMIENTO B " & _
                    "where A.COD_TIPO_REGISTRO = B.COD_TIPO_REGISTRO And COD_USR = " & Stg_cod_Usuario
    Else
        'si el perfil no es consulta parcial, entonces es consulta total, entonces no filtramos los tipos de movimiento.
        sentencia = "SELECT COD_TIPO_REGISTRO CODIGO, DESC_TIPO_REGISTRO DESCRIPCION FROM TBL_TIPO_REGISTRO"
    End If

    Set rsobj = cnObj1.Execute(sentencia)

    'Llenamos el combo de tipos de movimientos
    PG_Llenar_Combo_Lista2 rsobj, Me.cmb_tipo, "CODIGO", "DESCRIPCION", 2
    'Seleccionamos por defecto el indice 0 (el primero)
    Me.cmb_tipo.ListIndex = 0
    Set rsobj = Nothing
    Set cnObj1 = Nothing

End Sub
Sub pl_cargar_lista()
'Dim i As Integer
'Llena combo tipo registro
'Me.lst_configuracion.Clear
'If Val(Mid(Me.cmb_tipo, 1, 2)) = 2 Then
'
'    Me.lst_configuracion.AddItem "Código Entidad Concilia"
'    Me.lst_configuracion.AddItem "Aplicación Fuente Concilia"
'    Me.lst_configuracion.AddItem "Código de la Transacción Concilia"

'End If

'Me.lst_configuracion.AddItem "Código Entidad"
'Me.lst_configuracion.AddItem "Aplicación Fuente"
'Me.lst_configuracion.AddItem "Código de la Transacción"
'Me.lst_configuracion.AddItem "Fecha Contable"
'Me.lst_configuracion.AddItem "Centro Origen"
'Me.lst_configuracion.AddItem "Centro Destino"
'Me.lst_configuracion.AddItem "Tipo de Cuenta"
'Me.lst_configuracion.AddItem "Cuenta"

'i = 0
'While i < Me.lst_configuracion.ListCount
'     Me.lst_configuracion.ItemData(i) = 1
'     i = i + 1

'Wend

End Sub
Sub pl_cargar_lista_imp()
    Dim I As Integer

    'Me.lst_imprimir.Clear


    If Val(Mid(Me.cmb_tipo, 1, 2)) = 3 Or Val(Mid(Me.cmb_tipo, 1, 2)) = 5 Then

        Me.lst_imprimir.AddItem "Consulta"
        Me.lst_imprimir.AddItem "Transacciones Detalladas"


    End If

    If Val(Mid(Me.cmb_tipo, 1, 2)) = 2 Or Val(Mid(Me.cmb_tipo, 1, 2)) = 4 Then

        Me.lst_imprimir.AddItem "Consulta"
        Me.lst_imprimir.AddItem "Transacciones Consolidadas"
        Me.lst_imprimir.AddItem "Transacciones Detalladas"
        Me.lst_imprimir.AddItem "Asientos Contables"

    End If

    I = 0
    While I < Me.lst_imprimir.ListCount
        Me.lst_imprimir.ItemData(I) = 1
        I = I + 1
    Wend

End Sub

Sub pl_cargar_aplicacion(Filtrado As Boolean)
    On Error Resume Next
    PL_Conexion_Oracle
    ' REQ:266003. Ajustes por administración de usuarios por IDM.
    ' ABOCANE. DICIEMBRE 2016
    If Filtrado Then
        sentencia = "SELECT DISTINCT A.COD_APLICACION_FUENTE CODIGO, A.NOM_APLICACION_FUENTE NOMBRE" & _
                  " FROM TBL_APLICACION_FUENTE A, VTA_USUARIO_TRANSACCION B WHERE " & _
                  " A.COD_APLICACION_FUENTE = B.COD_APLICACION_FUENTE And B.COD_USR = " & Stg_cod_Usuario & _
                  " AND A.COD_APLICACION_FUENTE  NOT LIKE '@%' ORDER BY A.COD_APLICACION_FUENTE"
    Else
        sentencia = "SELECT COD_APLICACION_FUENTE CODIGO, NOM_APLICACION_FUENTE NOMBRE FROM TBL_APLICACION_FUENTE WHERE COD_APLICACION_FUENTE  NOT LIKE '@%' ORDER BY COD_APLICACION_FUENTE"

    End If
    Set rsobj = cnObj1.Execute(sentencia)
    PG_Llenar_Combo_Lista2 rsobj, Me.cmb_aplicacion, "CODIGO", "NOMBRE", 2

    Me.cmb_aplicacion.ListIndex = -1
    Set rsobj = Nothing
    Set cnObj1 = Nothing
End Sub
Sub pl_cargar_tipo_cuenta()
    Me.cmb_tipo_cuenta.AddItem "00-NO APLICA"
    Me.cmb_tipo_cuenta.AddItem "01-CUENTA CORRIENTE"
    Me.cmb_tipo_cuenta.AddItem "02-CUENTA DE AHORROS"
    Me.cmb_tipo_cuenta.AddItem "03-CDT"
    Me.cmb_tipo_cuenta.AddItem "04-TARJETA CREDITO"
    Me.cmb_tipo_cuenta.AddItem "05-CARTERA CREDITO"
    Me.cmb_tipo_cuenta.AddItem "06-SISCOI"
    Me.cmb_tipo_cuenta.AddItem "07-INTERBANCARIO"
    Me.cmb_tipo_cuenta.AddItem "08-CUENTA CONTABLE"
    Me.cmb_tipo_cuenta.AddItem "09-FORWARD"
    Me.cmb_tipo_cuenta.AddItem "10-TITULOS VALORES"
    Me.cmb_tipo_cuenta.AddItem "11-SERVICIOS PUBLICOS"
    Me.cmb_tipo_cuenta.AddItem "12-TRANSFERENCIAS"
End Sub

Function FL_Crea_Consulta_Impresion(where As String, campos As String, from As String, order As String) As Integer
    Dim datos As String
    Dim CAMPOS1 As String


    If ING_BD_CONSULTAR = True Then
        INL_BD_CONSULTAR = 1
        sqlTablas = " USRLARC.tbl_campos_valor_trad"
    Else
        INL_BD_CONSULTAR = 0
        sqlTablas = " USRLARC.tbl_campos_valor_trad@ARCHIST"
    End If

    If Me.DBG_DETALLE.Visible = True Then
        where = where & " AND  A.COD_ENTIDAD = " & Val(dbg_consulta.Columns(0)) & _
              " AND A.APLICACION_FUENTE = '" & dbg_consulta.Columns(1) & "'" & _
              " AND A.COD_TRANSACCION = '" & Format(dbg_consulta.Columns(2), "0###") & "'" & _
              " AND A.FECHA_MOVIMIENTO = TO_DATE('" & Format(dbg_consulta.Columns(3), "DD/MM/YYYY") & "','DD/MM/YYYY') " & _
              " AND A.CENTRO_ORIGEN = " & Val(dbg_consulta.Columns(4)) & _
              " AND A.CENTRO_DESTINO = " & Val(dbg_consulta.Columns(5))
    End If


    If Val(Mid(Me.cmb_tipo, 1, 2)) = 4 Then
        campos = "A.COD_ENTIDAD,A.APLICACION_FUENTE,A.COD_TRANSACCION,FECHA_MOVIMIENTO,CENTRO_ORIGEN," & _
                 "CENTRO_DESTINO,TIPO_CUENTA,COD_CUENTA,num_documento,FILLER," & _
                 "NUM_CAMPOS_MONETARIOS,TAM_CAMPOS_MONETARIOS, A.FECHA_SISTEMA , A.SEC_REGISTRO, COD_TIPO_REGISTRO"

        CAMPOS1 = "COD_ENTIDAD,APLICACION_FUENTE,COD_TRANSACCION,FECHA_MOVIMIENTO,CENTRO_ORIGEN," & _
                  "CENTRO_DESTINO,TIPO_CUENTA,COD_CUENTA,num_documento,FILLER," & _
                  "NUM_CAMPOS_MONETARIOS,TAM_CAMPOS_MONETARIOS,FECHA_SISTEMA , SEC_REGISTRO, COD_TIPO_REGISTRO"

    Else
        If Val(Mid(Me.cmb_tipo, 1, 2)) = 2 Then
            campos = "COD_ENTIDAD_CONCIL,APLICACION_FUENTE_CONCIL,COD_TRANSACCION_CONCIL,COD_ENTIDAD,APLICACION_FUENTE,COD_TRANSACCION,FECHA_MOVIMIENTO,CENTRO_ORIGEN," & _
                     "CENTRO_DESTINO,TIPO_CUENTA,COD_CUENTA,num_documento,FILLER," & _
                     "NUM_CAMPOS_MONETARIOS,TAM_CAMPOS_MONETARIOS," & _
                     "FECHA_SISTEMA , SEC_REGISTRO "

            CAMPOS1 = campos
        Else
            If Val(Mid(Me.cmb_tipo, 1, 2)) = 6 Then
                CAMPOS_CONSULTA = " APLICACION_FUENTE,COD_TRANSACCION,FECHA_CONTABLE,FECHA_SISTEMA,CENTRO_ORIGEN,CENTRO_DESTINO,COD_TIPO_ASIENTO, " & _
                                " CTA_AUXILIAR,CTA_FMS,VALOR_ASIENTO,ID_CAMPO"
                CAMPOS1 = campos
            Else
                campos = "COD_ENTIDAD,APLICACION_FUENTE,COD_TRANSACCION,FECHA_MOVIMIENTO,CENTRO_ORIGEN," & _
                         "CENTRO_DESTINO,TIPO_CUENTA,COD_CUENTA,num_documento,FILLER," & _
                         "NUM_CAMPOS_MONETARIOS,TAM_CAMPOS_MONETARIOS," & _
                         "FECHA_SISTEMA , SEC_REGISTRO, COD_TIPO_REGISTRO"
                CAMPOS1 = campos


            End If
        End If
    End If
    If Val(Mid(Me.cmb_tipo, 1, 2)) = 6 Then
        datos = " SELECT '" & Stg_cod_Usuario & "'," & campos & " FROM " & from & " Where " & where
    End If
    If Val(Mid(Me.cmb_tipo, 1, 2)) = 4 Then
        datos = " SELECT '" & Stg_cod_Usuario & "'," & campos & " FROM " & from & " Where " & where & " GROUP BY " & campos
    Else
        datos = " SELECT '" & Stg_cod_Usuario & "'," & campos & " FROM " & from & " Where " & where & " GROUP BY " & campos
    End If

    If Val(Mid(Me.cmb_tipo, 1, 2)) <> 2 And Val(Mid(Me.cmb_tipo, 1, 2)) <> 6 Then

        'Conexion a la base de datos
        PL_Conexion_Oracle

        'inserta fecha de la nueva contraseña
        sentencia = "DELETE tbl_registro_traductor_rpt WHERE COD_USR_SOLICITO = '" & Stg_cod_Usuario & "'"
        cnObj1.Execute (sentencia)
        sentencia = "DELETE tbl_registro_traductor_cons WHERE COD_USR_SOLICITO = '" & Stg_cod_Usuario & "'"
        cnObj1.Execute (sentencia)
        sentencia = "DELETE USRLARC.tbl_campos_valor_trad_rpt WHERE COD_USR_SOLICITO = '" & Stg_cod_Usuario & "'"
        cnObj1.Execute (sentencia)
    Else
        'Conexion a la base de datos
        PL_Conexion_Oracle

        If Val(Mid(Me.cmb_tipo, 1, 2)) = 6 Then
            sentencia = "DELETE tbl_registro_bana_rpt WHERE COD_USR_SOLICITO = '" & Stg_cod_Usuario & "'"
            cnObj1.Execute (sentencia)
        Else
            sentencia = "DELETE tbl_conciliacion_rpt WHERE COD_USR_SOLICITO = '" & Stg_cod_Usuario & "'"
            cnObj1.Execute (sentencia)
            sentencia = "DELETE USRLARC.tbl_campos_valor_conc_rpt WHERE COD_USR_SOLICITO = '" & Stg_cod_Usuario & "'"
            cnObj1.Execute (sentencia)
        End If
    End If



    If Val(Mid(Me.cmb_tipo, 1, 2)) <> 2 And Val(Mid(Me.cmb_tipo, 1, 2)) <> 6 Then
        If Val(Mid(Me.cmb_tipo, 1, 2)) = 4 Then
            sentencia = "INSERT INTO tbl_registro_traductor_rpt(Cod_usr_solicito, " & CAMPOS1 & ")(" & datos & ")"
            cnObj1.Execute (sentencia)

        Else


            If Val(Mid(Me.cmb_tipo, 1, 2)) = 1 Then
                If Flag_nuevo = 1 Then
                    sentencia = "INSERT INTO tbl_registro_traductor_cons(Cod_usr_solicito, " & CAMPOS1 & ")(" & datos & ")"
                    cnObj1.Execute (sentencia)
                    Set conex = New ADODB.Connection
                    'conex.ConnectionString = "provider = MSDAORA;Data Source =" & STG_NOMBRE_BD_HOST & ";User ID=" & STG_USR_BASE_HOST & ";Password=" & STG_CLAVE_BASE_HOST & ";"
                    conex.ConnectionString = "provider = " & STG_PROVIDER_HOST & ";Data Source =" & STG_NOMBRE_BD_HOST & ";User ID=" & STG_USR_BASE_HOST & ";Password=" & STG_CLAVE_BASE_HOST & ";"
                    conex.Open
                    conex.Execute "USRLARC.PL_CAMPOS_VALOR_TR('" & Stg_cod_Usuario & "'," & INL_BD_CONSULTAR & ") "
                Else
                    sentencia = "INSERT INTO tbl_registro_traductor_rpt(Cod_usr_solicito, " & CAMPOS1 & ")(" & datos & ")"
                    'MsgBox sentencia
                    cnObj1.Execute (sentencia)

                    sentencia = "insert into usrlarc.tbl_campos_valor_trad_rpt " & _
                              " select " & Stg_cod_Usuario & ", FECHA_SISTEMA, COD_TIPO_REGISTRO, SEC_REGISTRO, COD_CAMPO, VALOR_CAMPO from " & sqlTablas & _
                              " where (fecha_sistema,cod_tipo_registro,sec_registro) in (" & _
                              " select fecha_sistema,cod_tipo_registro,sec_registro FROM tbl_registro_traductor_rpt" & _
                              " where cod_usr_solicito = " & Stg_cod_Usuario & ")"

                    'MsgBox sentencia
                    cnObj1.Execute (sentencia)

                End If
            Else

                If ING_BD_CONSULTAR = True Then
                    INL_BD_CONSULTAR = 1
                    sqlTablas = " USRLARC.tbl_campos_valor_conc"
                Else
                    INL_BD_CONSULTAR = 0
                    sqlTablas = " USRLARC.tbl_campos_valor_conc@ARCHIST"
                End If
                sentencia = "INSERT INTO tbl_registro_traductor_cons(Cod_usr_solicito, " & CAMPOS1 & ")(" & datos & ")"
                cnObj1.Execute (sentencia)

                Set conex = New ADODB.Connection
                conex.ConnectionString = "provider = " & STG_PROVIDER_HOST & " ;Data Source =" & STG_NOMBRE_BD_HOST & ";User ID=" & STG_USR_BASE_HOST & ";Password=" & STG_CLAVE_BASE_HOST & ";"
                'conex.ConnectionString = "provider = MSDAORA ;Data Source =" & STG_NOMBRE_BD_HOST & ";User ID=" & STG_USR_BASE_HOST & ";Password=" & STG_CLAVE_BASE_HOST & ";"
                conex.Open
                conex.Execute "USRLARC.PL_CAMPOS_VALOR_TR('" & Stg_cod_Usuario & "'," & INL_BD_CONSULTAR & ") "

                sentencia = "insert into usrlarc.tbl_campos_valor_conc_rpt " & _
                          " select " & Stg_cod_Usuario & ", FECHA_SISTEMA,  SEC_REGISTRO, COD_CAMPO, VALOR_CAMPO from " & sqlTablas & _
                          " where (fecha_sistema,sec_registro) in ( " & _
                          " select fecha_sistema,sec_registro FROM tbl_conciliacion_rpt " & _
                          " where cod_usr_solicito = " & Stg_cod_Usuario & ")"
                cnObj1.Execute (sentencia)



            End If

        End If
    Else

        If ING_BD_CONSULTAR = True Then
            INL_BD_CONSULTAR = 1
            sqlTablas = " USRLARC.tbl_campos_valor_conc"
        Else
            INL_BD_CONSULTAR = 0
            sqlTablas = " USRLARC.tbl_campos_valor_conc@ARCHIST"
        End If

        If Val(Mid(Me.cmb_tipo, 1, 2)) = 6 Then
            sentencia = "INSERT INTO tbl_registro_bana_rpt(Cod_usr_solicito, " & CAMPOS1 & ")(" & datos & ")"
            cnObj1.Execute (sentencia)
        Else
            sentencia = "INSERT INTO tbl_conciliacion_rpt(Cod_usr_solicito, " & CAMPOS1 & ")(" & datos & ")"
            cnObj1.Execute (sentencia)

            sentencia = "insert into usrlarc.tbl_campos_valor_conc_rpt " & _
                      " select " & Stg_cod_Usuario & ", FECHA_SISTEMA,  SEC_REGISTRO, COD_CAMPO, VALOR_CAMPO from " & sqlTablas & _
                      " where (fecha_sistema,sec_registro) in ( " & _
                      " select fecha_sistema,sec_registro FROM tbl_conciliacion_rpt " & _
                      " where cod_usr_solicito = " & Stg_cod_Usuario & ")"
            cnObj1.Execute (sentencia)
        End If
    End If





    Exit Function

Error_Consulta_convenio:

    If Err = 7005 Then
        MsgBox "Error"
        Exit Function
    End If

End Function

Sub Fl_Crea_Consulta_Conciliacion()
    Dim sqlTablas As String
    Dim sqlEntidad As String
    Dim sqlTransaccionConciliacion As String
    Dim sqlAplicacionFuente As String
    Dim sqlTransaccion As String
    Dim sqlCentroOrigen As String
    Dim sqlCentroDestino As String
    Dim sqlFechaMovimiento As String
    Dim sqlFechaProceso As String
    Dim sqlTipoCuenta As String
    Dim sqlNumeroCuenta As String
    Dim sqlSecuencia As String
    Dim sqlValorMinimo As String
    Dim sqlValorMaximo As String
    Dim sqlUsuario As String

    Dim textoAuxiliar As String


    'Se selecciona la vista adecuada dependiendo de si esta o no seleccionado el filtro
    'If (Me.chkFiltro.Value = 0) Then
    If (Me.Cbo_filtro.ListIndex = 0) Then
        If Trim(Me.msk_secuencia.Text) <> "" Or Val(Me.msk_valor_min) <> 0 Or Val(Me.msk_valor_max) <> 0 Then
            If ING_BD_CONSULTAR = True Then
                sqlTablas = " VTA_REG_CONC A"
            Else
                sqlTablas = " USRLARC.VTA_REG_CONC A@ARCHIST  A "
            End If

        Else
            If ING_BD_CONSULTAR = True Then
                sqlTablas = " VTA_REG_CONC_USR A"
            Else
                sqlTablas = " USRLARC.VTA_REG_CONC_USR@ARCHIST  A "
            End If

            sqlUsuario = " AND COD_USR = " & Stg_cod_Usuario
        End If
    Else
        If Trim(Me.msk_secuencia.Text) <> "" Or Val(Me.msk_valor_min) <> 0 Or Val(Me.msk_valor_max) <> 0 Then
            If ING_BD_CONSULTAR = True Then
                sqlTablas = " VTA_REG_CONC A"
            Else
                sqlTablas = " USRLARC.VTA_REG_CONC@ARCHIST  A "
            End If
        Else
            If ING_BD_CONSULTAR = True Then
                sqlTablas = " VTA_REGISTRO_CONC A"
            Else
                sqlTablas = " USRLARC.VTA_REGISTRO_CONC@ARCHIST  A "
            End If

        End If
    End If



    'Condicion forzada para la estructura del query
    sqlEntidad = " A.COD_ENTIDAD=1 "


    'Condición para la aplicación fuente
    If Me.cmb_aplicacion.ListIndex <> -1 Or Me.cmb_aplicacion.Text <> "" Then
        sqlAplicacionFuente = " AND A.APLICACION_FUENTE = '" & Mid$(Me.cmb_aplicacion, 1, 4) & "'"
    End If


    If (Mid$(Me.cmb_aplicacion, 1, 4) = "CONC") Then
        sqlAplicacionFuente = ""
        'Condición para la transacción si se seleccionó la aplicacion fuente CONC
        sqlTransaccionConciliacion = ""
        sqlTransaccion = ""
        textoAuxiliar = filtrarTexto(Me.msk_transaccion.Text, 4, 4, "0")
        If Trim(textoAuxiliar) <> "" Then
            Select Case Len(textoAuxiliar)
            Case 4
                sqlTransaccionConciliacion = " AND A.COD_TRANSACCION_CONCIL ='" & textoAuxiliar & "'"
            Case 8
                sqlTransaccionConciliacion = " AND (A.COD_TRANSACCION_CONCIL ='" & Mid(textoAuxiliar, 1, 4) & "'"
                sqlTransaccionConciliacion = sqlTransaccionConciliacion + " OR A.COD_TRANSACCION_CONCIL ='" & Mid(textoAuxiliar, 5, 4) & "')"
            Case 12
                sqlTransaccionConciliacion = " AND (A.COD_TRANSACCION_CONCIL ='" & Mid(textoAuxiliar, 1, 4) & "'"
                sqlTransaccionConciliacion = sqlTransaccionConciliacion + " OR A.COD_TRANSACCION_CONCIL ='" & Mid(textoAuxiliar, 5, 4) & "'"
                sqlTransaccionConciliacion = sqlTransaccionConciliacion + " OR A.COD_TRANSACCION_CONCIL ='" & Mid(textoAuxiliar, 9, 4) & "')"
            Case 16
                sqlTransaccionConciliacion = " AND (A.COD_TRANSACCION_CONCIL ='" & Mid(Me.msk_transaccion.Text, 1, 4) & "'"
                sqlTransaccionConciliacion = sqlTransaccionConciliacion + " OR A.COD_TRANSACCION_CONCIL ='" & Mid(textoAuxiliar, 5, 4) & "'"
                sqlTransaccionConciliacion = sqlTransaccionConciliacion + " OR A.COD_TRANSACCION_CONCIL ='" & Mid(textoAuxiliar, 9, 4) & "'"
                sqlTransaccionConciliacion = sqlTransaccionConciliacion + " OR A.COD_TRANSACCION_CONCIL ='" & Mid(textoAuxiliar, 13, 4) & "')"
            End Select
        End If

    Else

        'Condición para la transacción si se selecciono aplicacion fuente distinta de CONC
        sqlTransaccion = ""
        sqlTransaccionConciliacion = ""
        textoAuxiliar = filtrarTexto(Me.msk_transaccion.Text, 4, 4, "0")
        If Trim(textoAuxiliar) <> "" Then
            Select Case Len(textoAuxiliar)
            Case 4
                sqlTransaccion = " AND A.COD_TRANSACCION ='" & textoAuxiliar & "'"
            Case 8
                sqlTransaccion = " AND (A.COD_TRANSACCION ='" & Mid(textoAuxiliar, 1, 4) & "'"
                sqlTransaccion = sqlTransaccion + " OR A.COD_TRANSACCION ='" & Mid(textoAuxiliar, 5, 4) & "')"
            Case 12
                sqlTransaccion = " AND (A.COD_TRANSACCION ='" & Mid(textoAuxiliar, 1, 4) & "'"
                sqlTransaccion = sqlTransaccion + " OR A.COD_TRANSACCION ='" & Mid(textoAuxiliar, 5, 4) & "'"
                sqlTransaccion = sqlTransaccion + " OR A.COD_TRANSACCION ='" & Mid(textoAuxiliar, 9, 4) & "')"
            Case 16
                sqlTransaccion = " AND (A.COD_TRANSACCION ='" & Mid(Me.msk_transaccion.Text, 1, 4) & "'"
                sqlTransaccion = sqlTransaccion + " OR A.COD_TRANSACCION ='" & Mid(textoAuxiliar, 5, 4) & "'"
                sqlTransaccion = sqlTransaccion + " OR A.COD_TRANSACCION ='" & Mid(textoAuxiliar, 9, 4) & "'"
                sqlTransaccion = sqlTransaccion + " OR A.COD_TRANSACCION ='" & Mid(textoAuxiliar, 13, 4) & "')"
            End Select
        End If
    End If
    'Condición para la unidad de negocio origen
    sqlCentroOrigen = ""
    textoAuxiliar = filtrarTexto(Me.msk_origen.Text, 4, 4, "0")
    If Trim(textoAuxiliar) <> "" Then
        Select Case Len(textoAuxiliar)
        Case 4
            sqlCentroOrigen = " AND CENTRO_ORIGEN ='" & textoAuxiliar & "'"
        Case 8
            sqlCentroOrigen = " AND (CENTRO_ORIGEN ='" & Mid(textoAuxiliar, 1, 4) & "'"
            sqlCentroOrigen = sqlCentroOrigen + " OR CENTRO_ORIGEN ='" & Mid(textoAuxiliar, 5, 4) & "')"
        Case 12
            sqlCentroOrigen = " AND (CENTRO_ORIGEN ='" & Mid(textoAuxiliar, 1, 4) & "'"
            sqlCentroOrigen = sqlCentroOrigen + " OR CENTRO_ORIGEN ='" & Mid(textoAuxiliar, 5, 4) & "'"
            sqlCentroOrigen = sqlCentroOrigen + " OR CENTRO_ORIGEN ='" & Mid(textoAuxiliar, 9, 4) & "')"
        Case 16
            sqlCentroOrigen = " AND (CENTRO_ORIGEN ='" & Mid(Me.msk_transaccion.Text, 1, 4) & "'"
            sqlCentroOrigen = sqlCentroOrigen + " OR CENTRO_ORIGEN ='" & Mid(textoAuxiliar, 5, 4) & "'"
            sqlCentroOrigen = sqlCentroOrigen + " OR CENTRO_ORIGEN ='" & Mid(textoAuxiliar, 9, 4) & "'"
            sqlCentroOrigen = sqlCentroOrigen + " OR CENTRO_ORIGEN ='" & Mid(textoAuxiliar, 13, 4) & "')"
        End Select
    End If

    'Condición para la unidad de negocio destino
    sqlCentroDestino = ""
    textoAuxiliar = filtrarTexto(Me.msk_destino.Text, 4, 4, "0")
    If Trim(textoAuxiliar) <> "" Then
        Select Case Len(textoAuxiliar)
        Case 4
            sqlCentroDestino = " AND CENTRO_DESTINO ='" & textoAuxiliar & "'"
        Case 8
            sqlCentroDestino = " AND (CENTRO_DESTINO ='" & Mid(textoAuxiliar, 1, 4) & "'"
            sqlCentroDestino = sqlCentroDestino + " OR CENTRO_DESTINO ='" & Mid(textoAuxiliar, 5, 4) & "')"
        Case 12
            sqlCentroDestino = " AND (CENTRO_DESTINO ='" & Mid(textoAuxiliar, 1, 4) & "'"
            sqlCentroDestino = sqlCentroDestino + " OR CENTRO_DESTINO ='" & Mid(textoAuxiliar, 5, 4) & "'"
            sqlCentroDestino = sqlCentroDestino + " OR CENTRO_DESTINO ='" & Mid(textoAuxiliar, 9, 4) & "')"
        Case 16
            sqlCentroDestino = " AND (CENTRO_DESTINO ='" & Mid(Me.msk_transaccion.Text, 1, 4) & "'"
            sqlCentroDestino = sqlCentroDestino + " OR CENTRO_DESTINO ='" & Mid(textoAuxiliar, 5, 4) & "'"
            sqlCentroDestino = sqlCentroDestino + " OR CENTRO_DESTINO ='" & Mid(textoAuxiliar, 9, 4) & "'"
            sqlCentroDestino = sqlCentroDestino + " OR CENTRO_DESTINO ='" & Mid(textoAuxiliar, 13, 4) & "')"
        End Select
    End If

    'Condición para la fecha de movimiento
    If msk_Fecha_ini <> "  /  /    " Then
        sqlFechaMovimiento = " AND A.FECHA_MOVIMIENTO = to_date ('" & msk_Fecha_ini & "','DD/MM/YYYY')"
    Else
        sqlFechaMovimiento = " AND A.FECHA_MOVIMIENTO = A.FECHA_MOVIMIENTO "
    End If

    'Condición para la fecha de proceso
    If msk_Fecha_Proceso <> "  /  /    " Then
        sqlFechaProceso = " AND A.FECHA_SISTEMA = to_date ('" & msk_Fecha_Proceso & "','DD/MM/YYYY')"
    Else
        sqlFechaProceso = " AND A.FECHA_SISTEMA = FECHA_SISTEMA "
    End If

    'Condición para el tipo de cuenta
    If Me.cmb_tipo_cuenta <> "" Then
        sqlTipoCuenta = " AND TIPO_CUENTA = " & Val(Mid$(Me.cmb_tipo_cuenta, 1, 2))
    End If

    'Condición para el numero de la cuenta
    If Trim(Me.msk_cuenta.Text) <> "" Then
        sqlNumeroCuenta = " AND COD_CUENTA = '" & Format(msk_cuenta, "0###############") & "'"
    End If

    'Condición para el campo de valor

    If Trim(Me.msk_secuencia.Text) <> "" Then
        sqlSecuencia = " AND  COD_CAMPO  = " & Val(Me.msk_secuencia) & " "
    End If

    'Condición para el valor minimo
    If Val(Me.msk_valor_min) <> 0 Then
        sqlValorMinimo = " AND VALOR_CAMPO " & " >= " & Val(Me.msk_valor_min)
    End If

    'Condición para el valor maximo
    If Val(Me.msk_valor_max) <> 0 Then
        sqlValorMaximo = " AND VALOR_CAMPO " & " <= " & Val(Me.msk_valor_max)
    End If

    'Se asignan las sentencias a las variables globales

    CAMPOS_CONSULTA = " APLICACION_FUENTE_CONCIL,COD_TRANSACCION_CONCIL,APLICACION_FUENTE,COD_TRANSACCION,FECHA_MOVIMIENTO,FECHA_SISTEMA,CENTRO_ORIGEN," & _
                      "CENTRO_DESTINO,TIPO_CUENTA,COD_CUENTA,SEC_REGISTRO,DESC_TRANSACCION "
    FROM_CONSULTA = sqlTablas
    WHERE_CONSULTA = sqlEntidad + sqlTransaccionConciliacion + sqlAplicacionFuente + sqlTransaccion + sqlCentroOrigen + sqlCentroDestino + sqlFechaMovimiento + sqlFechaProceso + sqlTipoCuenta + sqlNumeroCuenta + sqlSecuencia + sqlValorMinimo + sqlValorMaximo + sqlUsuario

    tipo_configuracion = 2

    Me.sst_vlr.Visible = True
    Set OBJTABLA = Nothing
    Set rsobj = Nothing
    Me.dbg_consulta.Width = Screen.Width - 50
    Exit Sub

End Sub

'Funcion dedicada a traer parametros de la base de datos para mostrar datos privilegiados
'y si el usuario actual esta dentro de los usuarios privilegiados, entonces retorna true,
'si no se encuentra el parametro o si el usuario no esta dentro de los privilegiados, retorna false
Private Function usuarioPrivilegiado() As Boolean
    
    sentencia = "SELECT * FROM TBL_PARAMETRO WHERE COD_PARAMETRO=16"
    Set rsobj = cargarRecordSet(sentencia)
    If rsobj.EOF Then
        usuarioPrivilegiado = False
        Exit Function
    Else
        Dim parametroUsuarios As String
        parametroUsuarios = rsobj("PARAMETRO_ALFANUMERICO")
        
        usuarioPrivilegiado = False
        'split de los codigo de usuario separados por comas
        Dim parametros() As String
        parametros = Split(parametroUsuarios, ",")
        'recorremos cada codigo de usuario
        For I = LBound(parametros) To UBound(parametros)
            
            'si el codigo es igual al codigo del usuario actual, retornamos true
            If Trim(parametros(I)) = Stg_cod_Usuario Then
                usuarioPrivilegiado = True
                Exit Function
            End If
  
        Next
        'si no encuentra ninguna coincidencia ya se ha asignado usuarioPrivilegiado = False, asi que retornaria false
        
    End If
End Function


Sub FL_Crea_Consulta()

    Dim sqlTablas As String
    Dim sqlTipoMovimiento As String
    Dim sqlJoin As String
    Dim sqlAplicacionFuente As String
    Dim sqlTransaccion As String
    Dim sqlCentroOrigen As String
    Dim sqlCentroDestino As String
    Dim sqlFechaMovimiento As String
    Dim sqlFechaProceso As String
    Dim sqlTipoCuenta As String
    Dim sqlNumeroCuenta As String
    Dim sqlSecuencia As String
    Dim sqlValorMinimo As String
    Dim sqlValorMaximo As String
    Dim sqlFlagAjuste As String

    Dim sqlJoinTablaError As String

    Dim sqlUsuario As String

    Dim textoAuxiliar As String
    Dim tipoMovimiento As Integer

    CAMPOS_CONSULTA = "APLICACION_FUENTE,COD_TRANSACCION,FECHA_MOVIMIENTO,FECHA_SISTEMA,CENTRO_ORIGEN,CENTRO_DESTINO,TIPO_CUENTA,COD_CUENTA,DESC_TRANSACCION,COD_TIPO_REGISTRO,SEC_REGISTRO "

    'Se selecciona la vista adecuada dependiendo de si esta o no seleccionado el filtro

    'Si esta seleccionada la opcion 0 del combo de filtro
    'Filtro por Responsabilidades.
    If (Me.Cbo_filtro.ListIndex = 0) Then
        'Si la secuencia ingresada no es vacia ó El valor minimo es distinto de cero ó el valor maximo es distinto de cero

        'Es decir si hay algun valor en alguno de los tres campos (secuencia,valor minimo o valor maximo)
        If Trim(Me.msk_secuencia.Text) <> "" Or Val(Me.msk_valor_min) <> 0 Or Val(Me.msk_valor_max) <> 0 Then
            'si la base a consultar es la base en linea, seleccionamos la vista VTA_REG_TRAD_USUARIO
            If ING_BD_CONSULTAR = True Then
                sqlTablas = " VTA_REG_TRAD_USUARIO  A "
            Else
                'sino, seleccionamos la vista USRLARC.VTA_REG_TRAD_USUARIO@ARCHIST es decir del historico
                sqlTablas = " USRLARC.VTA_REG_TRAD_USUARIO@ARCHIST  A "
            End If

        Else
            'sino se ingresaron datos de secuencia,valor minimo y maximo entonces..
            'Si la base a consultar es en linea
            If ING_BD_CONSULTAR = True Then
            
                'dsantan: si el usuario actual es privilegiado consultamos la informacion sin enmascarar
                If usuarioPrivilegiado Then
                    sqlTablas = " VTA_REGISTRO_TRAD_USR_1  A "
                'si no, consultamos la informacion enmascarada
                Else
                    'Seleccionamos la vista VTA_REGISTRO_TRAD_USR
                    sqlTablas = " VTA_REGISTRO_TRAD_USR  A "
                End If

            Else
                'sino, seleccionamos la vista historica VTA_REGISTRO_TRAD_USR@ARCHIST
                sqlTablas = " USRLARC.VTA_REGISTRO_TRAD_USR@ARCHIST  A "
            End If

        End If
        'Asignamos la condicion para filtrar por usuario
        sqlUsuario = " AND COD_USR = " & Stg_cod_Usuario
    Else
        'Todas las transacciones con Filtro
        If (Me.Cbo_filtro.ListIndex = 1) Then
            If Trim(Me.msk_secuencia.Text) <> "" Or Val(Me.msk_valor_min) <> 0 Or Val(Me.msk_valor_max) <> 0 Then
                If ING_BD_CONSULTAR = True Then
                    sqlTablas = " VTA_REG_TRAD_FILTRO A "
                Else
                    sqlTablas = " USRLARC.VTA_REG_TRAD_FILTRO@ARCHIST  A "
                End If
            Else

                If ING_BD_CONSULTAR = True Then
                    sqlTablas = " VTA_REGISTRO_TRAD_FILTRO A "
                Else
                    sqlTablas = " USRLARC.VTA_REGISTRO_TRAD_FILTRO@ARCHIST  A "
                End If
            End If

            'Este Else de donde es? que se pretende hacer aca?
        Else
            If Trim(Me.msk_secuencia.Text) <> "" Or Val(Me.msk_valor_min) <> 0 Or Val(Me.msk_valor_max) <> 0 Then
                If ING_BD_CONSULTAR = True Then
                    sqlTablas = "  VTA_REG_TRAD A "
                Else
                    sqlTablas = " USRLARC.VTA_REG_TRAD@ARCHIST  A "
                End If
                sqlUsuario = ""
            Else
                If ING_BD_CONSULTAR = True Then
                    sqlTablas = " VTA_REGISTRO_TRADUCTOR A "
                Else
                    sqlTablas = " USRLARC.VTA_REGISTRO_TRADUCTOR@ARCHIST  A "
                End If
            End If
        End If
    End If

    'Condición para la aplicación fuente
    If Me.cmb_aplicacion.ListIndex <> -1 Or Me.cmb_aplicacion.Text <> "" Then
        sqlAplicacionFuente = " AND APLICACION_FUENTE = '" & Mid$(Me.cmb_aplicacion, 1, 4) & "'"
    End If

    'Condición para la transacción
    sqlTransaccion = ""
    textoAuxiliar = filtrarTexto(Me.msk_transaccion.Text, 4, 4, "0")
    If Trim(textoAuxiliar) <> "" Then
        Select Case Len(textoAuxiliar)
        Case 4
            sqlTransaccion = " AND A.COD_TRANSACCION ='" & textoAuxiliar & "'"
        Case 8
            sqlTransaccion = " AND (A.COD_TRANSACCION ='" & Mid(textoAuxiliar, 1, 4) & "'"
            sqlTransaccion = sqlTransaccion + " OR A.COD_TRANSACCION ='" & Mid(textoAuxiliar, 5, 4) & "')"
        Case 12
            sqlTransaccion = " AND (A.COD_TRANSACCION ='" & Mid(textoAuxiliar, 1, 4) & "'"
            sqlTransaccion = sqlTransaccion + " OR A.COD_TRANSACCION ='" & Mid(textoAuxiliar, 5, 4) & "'"
            sqlTransaccion = sqlTransaccion + " OR A.COD_TRANSACCION ='" & Mid(textoAuxiliar, 9, 4) & "')"
        Case 16
            sqlTransaccion = " AND (A.COD_TRANSACCION ='" & Mid(textoAuxiliar, 1, 4) & "'"
            sqlTransaccion = sqlTransaccion + " OR A.COD_TRANSACCION ='" & Mid(textoAuxiliar, 5, 4) & "'"
            sqlTransaccion = sqlTransaccion + " OR A.COD_TRANSACCION ='" & Mid(textoAuxiliar, 9, 4) & "'"
            sqlTransaccion = sqlTransaccion + " OR A.COD_TRANSACCION ='" & Mid(textoAuxiliar, 13, 4) & "')"
        End Select
    End If

    'Condición para la unidad de negocio origen
    sqlCentroOrigen = ""
    textoAuxiliar = filtrarTexto(Me.msk_origen.Text, 4, 4, "0")
    If Trim(textoAuxiliar) <> "" Then
        Select Case Len(textoAuxiliar)
        Case 4
            sqlCentroOrigen = " AND CENTRO_ORIGEN ='" & textoAuxiliar & "'"
        Case 8
            sqlCentroOrigen = " AND (CENTRO_ORIGEN ='" & Mid(textoAuxiliar, 1, 4) & "'"
            sqlCentroOrigen = sqlCentroOrigen + " OR CENTRO_ORIGEN ='" & Mid(textoAuxiliar, 5, 4) & "')"
        Case 12
            sqlCentroOrigen = " AND (CENTRO_ORIGEN ='" & Mid(textoAuxiliar, 1, 4) & "'"
            sqlCentroOrigen = sqlCentroOrigen + " OR CENTRO_ORIGEN ='" & Mid(textoAuxiliar, 5, 4) & "'"
            sqlCentroOrigen = sqlCentroOrigen + " OR CENTRO_ORIGEN ='" & Mid(textoAuxiliar, 9, 4) & "')"
        Case 16
            sqlCentroOrigen = " AND (CENTRO_ORIGEN ='" & Mid(textoAuxiliar, 1, 4) & "'"
            sqlCentroOrigen = sqlCentroOrigen + " OR CENTRO_ORIGEN ='" & Mid(textoAuxiliar, 5, 4) & "'"
            sqlCentroOrigen = sqlCentroOrigen + " OR CENTRO_ORIGEN ='" & Mid(textoAuxiliar, 9, 4) & "'"
            sqlCentroOrigen = sqlCentroOrigen + " OR CENTRO_ORIGEN ='" & Mid(textoAuxiliar, 13, 4) & "')"
        End Select
    End If

    'Condición para la unidad de negocio destino
    sqlCentroDestino = ""
    textoAuxiliar = filtrarTexto(Me.msk_destino.Text, 4, 4, "0")
    If Trim(textoAuxiliar) <> "" Then
        Select Case Len(textoAuxiliar)
        Case 4
            sqlCentroDestino = " AND CENTRO_DESTINO ='" & textoAuxiliar & "'"
        Case 8
            sqlCentroDestino = " AND (CENTRO_DESTINO ='" & Mid(textoAuxiliar, 1, 4) & "'"
            sqlCentroDestino = sqlCentroDestino + " OR CENTRO_DESTINO ='" & Mid(textoAuxiliar, 5, 4) & "')"
        Case 12
            sqlCentroDestino = " AND (CENTRO_DESTINO ='" & Mid(textoAuxiliar, 1, 4) & "'"
            sqlCentroDestino = sqlCentroDestino + " OR CENTRO_DESTINO ='" & Mid(textoAuxiliar, 5, 4) & "'"
            sqlCentroDestino = sqlCentroDestino + " OR CENTRO_DESTINO ='" & Mid(textoAuxiliar, 9, 4) & "')"
        Case 16
            sqlCentroDestino = " AND (CENTRO_DESTINO ='" & Mid(Me.msk_destino.Text, 1, 4) & "'"
            sqlCentroDestino = sqlCentroDestino + " OR CENTRO_DESTINO ='" & Mid(textoAuxiliar, 5, 4) & "'"
            sqlCentroDestino = sqlCentroDestino + " OR CENTRO_DESTINO ='" & Mid(textoAuxiliar, 9, 4) & "'"
            sqlCentroDestino = sqlCentroDestino + " OR CENTRO_DESTINO ='" & Mid(textoAuxiliar, 13, 4) & "')"
        End Select
    End If

    'Condición para la fecha de movimiento
    If msk_Fecha_ini <> "  /  /    " Then
        sqlFechaMovimiento = " AND FECHA_MOVIMIENTO = to_date ('" & msk_Fecha_ini & "','DD/MM/YYYY')"
    Else
        sqlFechaMovimiento = " AND FECHA_MOVIMIENTO = FECHA_MOVIMIENTO "
    End If

    'Condición para la fecha de proceso
    If msk_Fecha_Proceso <> "  /  /    " Then
        sqlFechaProceso = " AND A.FECHA_SISTEMA = to_date ('" & msk_Fecha_Proceso & "','DD/MM/YYYY')"
    Else
        sqlFechaProceso = " AND A.FECHA_SISTEMA = A.FECHA_SISTEMA "
    End If

    'Condición para el tipo de cuenta
    If Me.cmb_tipo_cuenta <> "" Then
        sqlTipoCuenta = " AND TIPO_CUENTA = " & Val(Mid$(Me.cmb_tipo_cuenta, 1, 2))
    End If

    'Condición para el numero de la cuenta
    If Trim(Me.msk_cuenta.Text) <> "" Then
        sqlNumeroCuenta = " AND COD_CUENTA = '" & Format(msk_cuenta, "0###############") & "'"
    End If

    'Condición para el campo de valor
    sqlCamposGrilla = ""
    If Trim(Me.msk_secuencia.Text) <> "" Then
        sqlSecuencia = " AND  A.COD_CAMPO  = " & Val(Me.msk_secuencia) & " "
        sqlCamposGrilla = sqlSecuencia
    End If

    'Condición para el valor minimo
    If Val(Me.msk_valor_min) <> 0 Then
        sqlValorMinimo = " AND A.VALOR_CAMPO " & " >= " & Val(Me.msk_valor_min)
    End If

    'Condición para el valor maximo
    If Val(Me.msk_valor_max) <> 0 Then
        sqlValorMaximo = " AND A.VALOR_CAMPO " & " <= " & Val(Me.msk_valor_max)
    End If

    'Condición para flag de ajuste (Cruce)
    If Me.chkFlagAjuste.Value = 1 Then
        sqlFlagAjuste = " AND flag_ajuste = 1"
    End If




    tipoMovimiento = Val(Mid(Me.cmb_tipo, 1, 2))
    tipo_configuracion = tipoMovimiento
    'Dependiendo del tipo de movimiento se construyen las condiciones de la consulta especiales
    Select Case tipoMovimiento
    Case TR_ENTRADA
        'Condición para el tipo de movimiento
        sqlTipoMovimiento = " AND   COD_TIPO_REGISTRO = 1 "
    Case TR_CODIFICAR
        'Condición para el tipo de movimiento
        sqlTipoMovimiento = " AND  COD_TIPO_REGISTRO = 3 "
    Case TR_INCONSISTENTES

        'Se selecciona la vista adecuada dependiendo de si esta o no seleccionado el filtro
        If (Me.Cbo_filtro.ListIndex = 0) Then
            If Trim(Me.msk_secuencia.Text) <> "" Or Val(Me.msk_valor_min) <> 0 Or Val(Me.msk_valor_max) <> 0 Then
                sqlTablas = "VTA_REG_TRAD_USUARIO A"
            Else
                sqlTablas = " VTA_REGISTRO_TRAD_USR A"
            End If
        Else
            If (Me.Cbo_filtro.ListIndex = 2) Then
                If Trim(Me.msk_secuencia.Text) <> "" Or Val(Me.msk_valor_min) <> 0 Or Val(Me.msk_valor_max) <> 0 Then
                    sqlTablas = "  VTA_REG_TRAD A "
                    sqlUsuario = ""
                Else
                    sqlTablas = " VTA_REGISTRO_TRADUCTOR A "
                End If
            End If
        End If

        'Condición para el tipo de movimiento
        sqlJoin = " "    'AND  A.FECHA_SISTEMA = B.FECHA_SISTEMA AND A.SEC_REGISTRO = B.SEC_REGISTRO AND B.COD_ERROR = C.COD_ERROR "
        sqlTipoMovimiento = " AND COD_TIPO_REGISTRO=4 "

        'CAMPOS_CONSULTA = "APLICACION_FUENTE,COD_TRANSACCION,FECHA_MOVIMIENTO,A.FECHA_SISTEMA,CENTRO_ORIGEN,CENTRO_DESTINO,TIPO_CUENTA,COD_CUENTA,DESC_TRANSACCION,COD_TIPO_REGISTRO,A.SEC_REGISTRO,DESC_ERROR,ACCION_CORRECTIVA "

    Case TR_NORMALIZAR
        'Condición para el tipo de movimiento
        sqlTipoMovimiento = " AND  COD_TIPO_REGISTRO=3 "

        'Condición para la aplicación fuente
        sqlAplicacionFuente = " AND APLICACION_FUENTE = 'CONC' "

        'Condición para flag de ajuste (Cruce)
        sqlFlagAjuste = " AND FLAG_AJUSTE IS NOT NULL AND FLAG_AJUSTE <> 0 "
        
        tipo_configuracion = 1
    End Select

    'Se asignan las sentencias a las variables globales
    FROM_CONSULTA = sqlTablas
    WHERE_CONSULTA = "A.FECHA_SISTEMA= A.FECHA_SISTEMA " + sqlJoin + sqlTipoMovimiento + sqlAplicacionFuente + sqlTransaccion + sqlCentroOrigen + sqlCentroDestino + sqlFechaMovimiento + sqlFechaProceso + sqlTipoCuenta + sqlNumeroCuenta + sqlSecuencia + sqlValorMinimo + sqlValorMaximo + sqlFlagAjuste + sqlJoinTablaError + sqlCodigoError + sqlUsuario
    WHERE_CONSULTA1 = "A.FECHA_SISTEMA= A.FECHA_SISTEMA " + sqlTipoMovimiento + sqlAplicacionFuente + sqlTransaccion + sqlCentroOrigen + sqlCentroDestino + sqlFechaMovimiento + sqlFechaProceso + sqlTipoCuenta + sqlNumeroCuenta + sqlSecuencia + sqlValorMinimo + sqlValorMaximo + sqlFlagAjuste + sqlJoinTablaError + sqlCodigoError + sqlUsuario
    Set OBJTABLA = Nothing
    Set rsobj = Nothing

    Exit Sub
End Sub

Sub Fl_Crea_Consulta_Bana()

    Dim sqlTablas As String
    Dim sqlEntidad As String
    Dim sqlAplicacionFuente As String
    Dim sqlTransaccion As String
    Dim sqlCentroOrigen As String
    Dim sqlCentroDestino As String
    Dim sqlFechaMovimiento As String
    Dim sqlFechaProceso As String
    Dim sqlTipoAsiento As String
    Dim sqlCuentaAuxiliar As String
    Dim sqlCuentaFMS As String
    Dim sqlValorMinimo As String
    Dim sqlValorMaximo As String
    Dim sqlUsuario As String
    Dim textoAuxiliar As String

    'REQ CVAPD00223966. Modificación en la consulta del libro auxiliar
    'para que tenga en cuenta el libro
    'ABOCANE Mayo 2016
    
    Dim libro_auxiliar As String
    If (Me.cmb_libro_auxiliar.ListIndex = 0) Then
        libro_auxiliar = "AND LIBRO_AUXILIAR= 1 "
    Else
    libro_auxiliar = "AND LIBRO_AUXILIAR=2"
    End If


    'Se selecciona la vista adecuada dependiendo de si esta o no seleccionado el filtro
    'If (Me.chkFiltro.Value = 0) Then
    If (Me.Cbo_filtro.ListIndex = 0) Then
        If ING_BD_CONSULTAR = True Then
            sqlTablas = " VTA_REG_BANA_USUARIO "
        Else
            sqlTablas = " USRLARC.VTA_REG_BANA_USUARIO@ARCHIST "
        
        End If

        sqlUsuario = " AND COD_USR = " & Stg_cod_Usuario
    Else
        If ING_BD_CONSULTAR = True Then
            sqlTablas = " VTA_REG_BANA "
        Else
            sqlTablas = " USRLARC.VTA_REG_BANA@ARCHIST "
    
        End If


    End If

    'Condicion forzada para la estructura del query
    sqlEntidad = " COD_ENTIDAD=1 "

    'Condición para la aplicación fuente
    If Me.cmb_aplicacion.ListIndex <> -1 Then
        sqlAplicacionFuente = " AND APLICACION_FUENTE = '" & Mid$(Me.cmb_aplicacion, 1, 4) & "'"
    End If

    'Condición para la transacción
    sqlTransaccion = ""
    textoAuxiliar = filtrarTexto(Me.msk_transaccion.Text, 4, 4, "0")
    If Trim(textoAuxiliar) <> "" Then
        Select Case Len(textoAuxiliar)
        Case 4
            sqlTransaccion = " AND COD_TRANSACCION ='" & textoAuxiliar & "'"
        Case 8
            sqlTransaccion = " AND (COD_TRANSACCION ='" & Mid(textoAuxiliar, 1, 4) & "'"
            sqlTransaccion = sqlTransaccion + " OR COD_TRANSACCION ='" & Mid(textoAuxiliar, 5, 4) & "')"
        Case 12
            sqlTransaccion = " AND (COD_TRANSACCION ='" & Mid(textoAuxiliar, 1, 4) & "'"
            sqlTransaccion = sqlTransaccion + " OR COD_TRANSACCION ='" & Mid(textoAuxiliar, 5, 4) & "'"
            sqlTransaccion = sqlTransaccion + " OR COD_TRANSACCION ='" & Mid(textoAuxiliar, 9, 4) & "')"
        Case 16
            sqlTransaccion = " AND (COD_TRANSACCION ='" & Mid(Me.msk_transaccion.Text, 1, 4) & "'"
            sqlTransaccion = sqlTransaccion + " OR COD_TRANSACCION ='" & Mid(textoAuxiliar, 5, 4) & "'"
            sqlTransaccion = sqlTransaccion + " OR COD_TRANSACCION ='" & Mid(textoAuxiliar, 9, 4) & "'"
            sqlTransaccion = sqlTransaccion + " OR COD_TRANSACCION ='" & Mid(textoAuxiliar, 13, 4) & "')"
        End Select
    End If

    'Condición para la unidad de negocio origen
    sqlCentroOrigen = ""
    textoAuxiliar = filtrarTexto(Me.msk_origen.Text, 4, 4, "0")
    If Trim(textoAuxiliar) <> "" Then
        Select Case Len(textoAuxiliar)
        Case 4
            sqlCentroOrigen = " AND CENTRO_ORIGEN ='" & textoAuxiliar & "'"
        Case 8
            sqlCentroOrigen = " AND (CENTRO_ORIGEN ='" & Mid(textoAuxiliar, 1, 4) & "'"
            sqlCentroOrigen = sqlCentroOrigen + " OR CENTRO_ORIGEN ='" & Mid(textoAuxiliar, 5, 4) & "')"
        Case 12
            sqlCentroOrigen = " AND (CENTRO_ORIGEN ='" & Mid(textoAuxiliar, 1, 4) & "'"
            sqlCentroOrigen = sqlCentroOrigen + " OR CENTRO_ORIGEN ='" & Mid(textoAuxiliar, 5, 4) & "'"
            sqlCentroOrigen = sqlCentroOrigen + " OR CENTRO_ORIGEN ='" & Mid(textoAuxiliar, 9, 4) & "')"
        Case 16
            sqlCentroOrigen = " AND (CENTRO_ORIGEN ='" & Mid(Me.msk_transaccion.Text, 1, 4) & "'"
            sqlCentroOrigen = sqlCentroOrigen + " OR CENTRO_ORIGEN ='" & Mid(textoAuxiliar, 5, 4) & "'"
            sqlCentroOrigen = sqlCentroOrigen + " OR CENTRO_ORIGEN ='" & Mid(textoAuxiliar, 9, 4) & "'"
            sqlCentroOrigen = sqlCentroOrigen + " OR CENTRO_ORIGEN ='" & Mid(textoAuxiliar, 13, 4) & "')"
        End Select
    End If

    'Condición para la unidad de negocio destino
    sqlCentroDestino = ""
    textoAuxiliar = filtrarTexto(Me.msk_destino.Text, 4, 4, "0")
    If Trim(textoAuxiliar) <> "" Then
        Select Case Len(textoAuxiliar)
        Case 4
            sqlCentroDestino = " AND CENTRO_DESTINO ='" & textoAuxiliar & "'"
        Case 8
            sqlCentroDestino = " AND (CENTRO_DESTINO ='" & Mid(textoAuxiliar, 1, 4) & "'"
            sqlCentroDestino = sqlCentroDestino + " OR CENTRO_DESTINO ='" & Mid(textoAuxiliar, 5, 4) & "')"
        Case 12
            sqlCentroDestino = " AND (CENTRO_DESTINO ='" & Mid(textoAuxiliar, 1, 4) & "'"
            sqlCentroDestino = sqlCentroDestino + " OR CENTRO_DESTINO ='" & Mid(textoAuxiliar, 5, 4) & "'"
            sqlCentroDestino = sqlCentroDestino + " OR CENTRO_DESTINO ='" & Mid(textoAuxiliar, 9, 4) & "')"
        Case 16
            sqlCentroDestino = " AND (CENTRO_DESTINO ='" & Mid(Me.msk_transaccion.Text, 1, 4) & "'"
            sqlCentroDestino = sqlCentroDestino + " OR CENTRO_DESTINO ='" & Mid(textoAuxiliar, 5, 4) & "'"
            sqlCentroDestino = sqlCentroDestino + " OR CENTRO_DESTINO ='" & Mid(textoAuxiliar, 9, 4) & "'"
            sqlCentroDestino = sqlCentroDestino + " OR CENTRO_DESTINO ='" & Mid(textoAuxiliar, 13, 4) & "')"
        End Select
    End If

    'Condición para la fecha de movimiento
    If msk_Fecha_ini <> "  /  /    " Then
        sqlFechaMovimiento = " AND FECHA_CONTABLE = to_date ('" & msk_Fecha_ini & "','DD/MM/YYYY')"
    Else
        sqlFechaMovimiento = " AND FECHA_CONTABLE = FECHA_CONTABLE "
    End If

    'Condición para la fecha de proceso
    If msk_Fecha_Proceso <> "  /  /    " Then
        sqlFechaProceso = " AND FECHA_SISTEMA = to_date ('" & msk_Fecha_Proceso & "','DD/MM/YYYY')"
    Else
        sqlFechaProceso = " AND FECHA_SISTEMA = FECHA_SISTEMA "
    End If

    'Condición para el tipo de asiento

    If Trim(Me.cmb_tipo_asiento) <> "" Then
        sqlTipoAsiento = " AND COD_TIPO_ASIENTO = '" & Mid(cmb_tipo_asiento, 1, 1) & "'"
    End If

    'Condición para la cuenta auxiliar

    If Trim(Me.msk_cta_aux.Text) <> "" Then
        sqlCuentaAuxiliar = " AND CTA_AUXILIAR = " & Me.msk_cta_aux.Text & ""
    End If

    'Condición para el valor minimo

    If Me.msk_valor_min.Text <> "" Then
        sqlValorMinimo = " AND VALOR_ASIENTO >= " & Val(msk_valor_min.Text)
    End If

    'Condición para el valor maximo
    If msk_valor_max.Text <> "" Then
        sqlValorMaximo = " AND VALOR_ASIENTO <= " & Val(msk_valor_max.Text)
    End If

    'Se asignan las sentencias a las variables globales
    CAMPOS_CONSULTA = " APLICACION_FUENTE,COD_TRANSACCION,FECHA_CONTABLE,FECHA_SISTEMA,CENTRO_ORIGEN,CENTRO_DESTINO,COD_TIPO_ASIENTO, " & _
                    "ID_CAMPO,CTA_AUXILIAR,VALOR_ASIENTO"
    FROM_CONSULTA = sqlTablas
    WHERE_CONSULTA = sqlEntidad + sqlAplicacionFuente + sqlTransaccion + sqlCentroOrigen + sqlCentroDestino + sqlFechaMovimiento + sqlFechaProceso + sqlTipoAsiento + sqlCuentaAuxiliar + sqlCuentaFMS + sqlValorMinimo + sqlValorMaximo + sqlUsuario + libro_auxiliar
    tipo_configuracion = 6

    Set OBJTABLA = Nothing
    Set rsobj = Nothing
    Me.dbg_consulta.Width = Screen.Width - 50
    Exit Sub
End Sub


Private Sub BotonSQL_Click()
    Me.sqlText.Visible = Not Me.sqlText.Visible
End Sub

Private Sub chkFiltro_Click()
'If (Me.chkFiltro.Value = 1) Then
    If (Me.Cbo_filtro.ListIndex = 0) Then
        pl_cargar_aplicacion False
    Else
        pl_cargar_aplicacion True
    End If
End Sub




Private Sub cmb_aplica_concil_GotFocus()
    cmb_aplica_concil.SelStart = 0
    cmb_aplica_concil.SelLength = Len(cmb_aplica_concil)
End Sub

Private Sub chkFlagAjuste_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        'Si la tecla es ESCAPE debe limpiar el contenido del control
    Case vbKeyEscape
        chkFlagAjuste.Value = 0
        KeyAscii = 0
        'Si la tecla es ENTER debe navegar al siguiente control
    Case vbKeyReturn
        'SendKeys "+{tab}"
        SendKeys "{tab}"
        KeyAscii = 0
    End Select
End Sub

Private Sub cmb_aplicacion_GotFocus()
    cmb_aplicacion.SelStart = 0
    cmb_aplicacion.SelLength = Len(cmb_aplicacion)
End Sub
Private Sub cmb_aplicacion_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        ' Si la tecla es ESCAPE debe limpiar el contenido del control
    Case vbKeyEscape
        cmb_aplicacion.Text = ""
        KeyAscii = 0

        ' Si la tecla es ENTER debe navegar al siguiente control
    Case vbKeyReturn
        'SendKeys "+{tab}"
        'SendKeys "{tab}"
        KeyAscii = 0
    End Select

End Sub

Private Sub cmb_tipo_asiento_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        'Si la tecla es ESCAPE debe limpiar el contenido del control
    Case vbKeyEscape
        cmb_tipo_asiento.Text = ""
        KeyAscii = 0
        'Si la tecla es ENTER debe navegar al siguiente control
    Case vbKeyReturn
        'SendKeys "+{tab}"
        SendKeys "{tab}"
        KeyAscii = 0
    End Select
End Sub

Private Sub cmb_tipo_Click()
    ORDER_CONSULTA = ""
    PL_CERRAR_CONEXIONES
    PL_CERRAR_CONEXIONES1
    pl_tipo_consulta
End Sub
Private Sub cmb_tipo_cuenta_GotFocus()
    cmb_tipo_cuenta.SelStart = 0
    cmb_tipo_cuenta.SelLength = Len(cmb_tipo_cuenta)
End Sub
Private Sub cmb_tipo_cuenta_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        'Si la tecla es ESCAPE debe limpiar el contenido del control
    Case vbKeyEscape
        cmb_tipo_cuenta.Text = "    "
        KeyAscii = 0

        'Si la tecla es ENTER debe navegar al siguiente control
    Case vbKeyReturn
        'SendKeys "+{tab}"
        SendKeys "{tab}"
        KeyAscii = 0
    End Select
End Sub
Private Sub cmb_tipo_GotFocus()
    cmb_tipo.SelStart = 0
    cmb_tipo.SelLength = Len(cmb_tipo)
End Sub

Private Sub cmb_tipo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        cmb_tipo.Text = ""
    Else

    End If
End Sub

Private Sub cmd_Cierre_Click()
    Me.dbg_consulta.Height = 5000
    Me.dbg_valor.Height = 4200
    Me.dtg_suma.Height = 4200
    Me.dtg_error.Height = 4200
    sst_vlr.Height = 5000

    PL_CERRAR_CONEXIONES
    PL_CERRAR_CONEXIONES2


    Me.tb_detalle.Visible = False
    Me.frm_procesos.Visible = False
    Me.cmd_Cierre.Enabled = False
End Sub

Private Sub cmd_salir_imp_Click()
    Me.frm_imprimir.Visible = False
End Sub



Private Sub Command1_Click()
    PL_CERRAR_CONEXIONES
    PL_CERRAR_CONEXIONES1
End Sub

Private Sub cmdParametrosConsulta_Click()
    Me.frameParametros.Visible = Not Me.frameParametros.Visible
    pl_tipo_consulta
End Sub


Private Sub dbg_valor_ButtonClick(ByVal ColIndex As Integer)
    On Error GoTo error
    Set rsobj = cargarRecordSet("SELECT   DESC_CAMPO FROM TBL_CAMPO_TRANSACCION WHERE COD_ENTIDAD = 1 AND COD_APLICACION_FUENTE =  '" & dbg_consulta.Columns(0) & "'  AND COD_TRANSACCION =  '" & dbg_consulta.Columns(1) & "' AND COD_CAMPO = " & dbg_valor.Columns(0))
    If rsobj.EOF Then
        MsgBox "No existe descripción para el campo " & dbg_valor.Columns(0) & " de la transacción " & dbg_consulta.Columns(0) & "-" & rellenar(dbg_consulta.Columns(1), 4, "0", "izquierda")
    Else
        Me.txt_desc.Text = UCase(rsobj(0))
    End If
    If Not rsobj Is Nothing Then
        If rsobj.State = adStateOpen Then rsobj.Close
    End If
    Set rsobj = Nothing
    Me.frm_desc.Visible = True
    Exit Sub
error:
    If Err.Number = 9 Then
        MsgBox "No existe descripción para el campo " & dbg_valor.Columns(0) & " de la transacción " & dbg_consulta.Columns(0) & "-" & rellenar(dbg_consulta.Columns(1), 4, "0", "izquierda")
    Else
        MsgBox Err.Number & " . " & Err.Description
    End If
End Sub


Private Sub dtg_valor_det_ButtonClick(ByVal ColIndex As Integer)

    On Error GoTo error
    Set rsobj = cargarRecordSet("SELECT   DESC_CAMPO FROM TBL_CAMPO_TRANSACCION WHERE COD_ENTIDAD = 1 AND COD_APLICACION_FUENTE =  '" & dtg_detalle2.Columns(1) & "'  AND COD_TRANSACCION =  '" & dtg_detalle2.Columns(2) & "' AND COD_CAMPO = " & dtg_valor_det.Columns(0))
    If rsobj.EOF Then
        MsgBox "No existe descripción para el campo " & dtg_valor_det.Columns(0) & " de la transacción " & dtg_detalle2.Columns(1) & "-" & rellenar(dtg_detalle2.Columns(2), 4, "0", "izquierda")
    Else
        Me.txt_desc.Text = UCase(rsobj(0))
    End If
    If Not rsobj Is Nothing Then
        If rsobj.State = adStateOpen Then rsobj.Close
    End If
    Set rsobj = Nothing
    Me.frm_desc.Visible = True
    Exit Sub
error:
    If Err.Number = 9 Then
        MsgBox "No existe descripción para el campo " & dtg_valor_det.Columns(0) & " de la transacción " & dtg_detalle2.Columns(1) & "-" & rellenar(dtg_detalle2.Columns(2), 4, "0", "izquierda")
    Else
        MsgBox Err.Number & " . " & Err.Description
    End If


End Sub


Private Sub dbg_consulta_Click()
    Me.dbg_valor.ClearFields
    dbg_consulta_KeyUp 0, 0
    Me.sst_vlr.Enabled = True
End Sub

Private Sub dbg_consulta_DblClick()
    On Error Resume Next
    
    Me.dtg_valor_det.ClearFields
    Me.dtg_valorConc.ClearFields

    carga_conc = 0
    carga_cont = 0
    carga_cont_1 = 0
    carga_det = 0
    carga_diferencias = 0
    
    
    Dim tipoSeleccionado As String
    
    'Almaceno temporalmente el tipo de movimiento seleccionado por el usuario
    tipoSeleccionado = Val(Mid(Me.cmb_tipo, 1, 2))

    'Si se ha seleccionado tipo de movimiento Entrada o Mov Auxiliar, no se hace nada
    If tipoSeleccionado <> TR_ENTRADA And tipoSeleccionado <> MOV_LIBRO_AUXILIAR Then
        Set RSOBJ3 = Nothing

        'Si la consulta actual NO es sobre una conciliacion y el tipo es para codificar
        If dbg_consulta.Columns(0) <> "CONC" And tipoSeleccionado = TR_CODIFICAR Then

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
        If Not RSOBJ3.EOF Then
            Me.dbg_consulta.Height = 2600
            Me.dbg_valor.Height = 1900
            Me.dtg_suma.Height = 1900
            Me.dtg_error.Height = 1900
            sst_vlr.Height = 2600

            If Mid(Val(Me.cmb_tipo.Text), 1, 1) = 5 Or Mid(Val(Me.cmb_tipo.Text), 1, 1) = 3 Then
                Set Me.DBG_DETALLE.DataSource = RSOBJ3
                Configurar_grilla_Dat_consulta DBG_DETALLE

                'Set Me.DBG_DETALLE.Object = Nothing

                tb_detalle_Click 0
                'Me.tb_detalle.TabVisible(0) = True
                'Me.tb_detalle.TabVisible(2) = True
                DBG_DETALLE_KeyUp 0, 0
            End If

            If Mid(Val(Me.cmb_tipo.Text), 1, 1) = 2 Or Mid(Val(Me.cmb_tipo.Text), 1, 1) = 4 Then
                Set Me.DBG_DETALLE.DataSource = RSOBJ3
                Configurar_grilla_Dat_consulta DBG_DETALLE
                'Set Me.DBG_DETALLE.Object = Nothing

                'Me.tb_detalle.TabVisible(0) = False
                tb_detalle_Click 0
                'Me.tb_detalle.TabVisible(2) = False

                Me.tb_detalle.TabIndex = 1
                DBG_DETALLE_KeyUp 0, 0
            End If
            Me.tb_detalle.Visible = True
            Me.cmd_Cierre.Enabled = True
            cmd_Imprimir.Enabled = False
        End If

        Me.tb_detalle.Tab = 0

    End If
End Sub
Private Sub Cmd_Cancelar_Click()

    PL_CERRAR_CONEXIONES
    PL_CERRAR_CONEXIONES1
    tb_detalle.Visible = False

    Me.cmb_aplicacion.Text = ""
    Me.cmb_tipo_cuenta.Text = ""
    Me.msk_cuenta.Text = "                "
    Me.msk_destino.Text = "    ,    ,    ,    "
    Me.msk_origen.Text = "    ,    ,    ,    "
    Me.msk_transaccion.Text = "    ,    ,    ,    "
    pl_carga_grilla_vacia
    Me.cmb_tipo.SetFocus

End Sub
Private Sub cmd_consultar_Click()
    Dim tipoMovimiento As Integer
    tipoMovimiento = Val(Mid(Me.cmb_tipo, 1, 2))

    'Oculta el panel de parametros
    frameParametros.Visible = False
    PL_CERRAR_CONEXIONES
    PL_CERRAR_CONEXIONES1

    'RCC - 09/10/2009 Determino cual base de datos debe consultar la historica o la en linea

    If msk_Fecha_Proceso.Text = "  /  /    " Then
        MsgBox "Por favor digite una fecha de proceso valida, se asigna la fecha actual"
        Me.msk_Fecha_Proceso.Text = Format(Dtg_Fecha_movimiento, "DD/MM/YYYY")
        Me.frameParametros.Visible = Not Me.frameParametros.Visible
        Exit Sub
    End If


    ING_BD_CONSULTAR = FG_Retorna_base_datos_consultar(msk_Fecha_Proceso)
    If ING_BD_CONSULTAR = True Then
        frm_desc.Visible = True
        txt_desc.Text = "CONSULTANDO BD EN LINEA"
    Else
        frm_desc.Visible = True
        txt_desc.Text = "CONSULTANDO BD HISTORICA"
    End If

    Dtg_Fecha_movimiento1 = msk_Fecha_Proceso.Text

    frm_desc.Refresh
    Me.dbg_consulta.Height = 5000
    Me.dbg_valor.Height = 4200
    Me.dtg_suma.Height = 4200
    Me.dtg_error.Height = 4200
    sst_vlr.Height = 5000

    'Variables de control para optimización
    cargoTotales = False
    cargoErrores = False



    secuencia_Reg = ""
    Me.sst_vlr.Tab = 0
    On Error GoTo error
    tipo_configuracion = 0
    MousePointer = 13
    If parametrosValidos(tipoMovimiento) Then

        'Se establecen las pestañas auxiliares de acuerdo al tipo de movimiento
        Select Case tipoMovimiento
        Case 4:
            Me.sst_vlr.Tabs = 3
            Me.sst_vlr.TabEnabled(2) = True
            Me.sst_vlr.Tab = 0
        Case 1, 2, 3, 5, 6:
            Me.sst_vlr.TabEnabled(2) = False
            Me.sst_vlr.Tab = 0
        End Select

        Select Case tipoMovimiento
        Case 1, 3, 4, 5:
            FL_Crea_Consulta
        Case 2:
            Fl_Crea_Consulta_Conciliacion
        Case 6:
            Fl_Crea_Consulta_Bana
        End Select

    Else
        MousePointer = 0
        Exit Sub
    End If
    pl_carga_grilla_vacia
    PL_Carga_Grilla tipoMovimiento, CAMPOS_CONSULTA, FROM_CONSULTA, WHERE_CONSULTA

    MousePointer = 0
    Me.tb_detalle.Visible = False
    Me.frm_procesos.Visible = False
    PL_CERRAR_CONEXIONES
    frm_desc.Visible = False

    Exit Sub

error:
    frm_desc.Visible = False
    MsgBox Err.Number & "-" & Err.Description
    MousePointer = 0
    Me.frm_procesos.Visible = False
End Sub
Private Sub cmd_Imprimir_Click()
    On Error GoTo error

    MousePointer = 13


    If Val(Mid(Me.cmb_tipo, 1, 2)) = 1 Then
        If MsgBox("Desea Imprimir en formato nuevo", vbYesNo) = vbYes Then
            Flag_nuevo = 1
        Else
            Flag_nuevo = 0
        End If
    End If

    CAMPOS_CONSULTA = "APLICACION_FUENTE,COD_TRANSACCION,FECHA_MOVIMIENTO,FECHA_SISTEMA,CENTRO_ORIGEN,CENTRO_DESTINO,TIPO_CUENTA,COD_CUENTA,DESC_TRANSACCION,COD_TIPO_REGISTRO,SEC_REGISTRO,DESC_ERROR,ACCION_CORRECTIVA "

    FL_Crea_Consulta_Impresion WHERE_CONSULTA, CAMPOS_CONSULTA, FROM_CONSULTA, ORDER_CONSULTA
    MDI_Administrador_contable.CrystalReport1.Connect = STG_CONEXION_REPORTE
    If Val(Mid(Me.cmb_tipo, 1, 2)) = 1 Then
        If Flag_nuevo = 1 Then
            MDI_Administrador_contable.CrystalReport1.ReportFileName = STG_PATH_REPORTES & "rpt_arc_trad.rpt"
            
            MDI_Administrador_contable.CrystalReport1.WindowState = crptMaximized
            
            MDI_Administrador_contable.CrystalReport1.SelectionFormula = "{TBL_REGISTRO_TRADUCTOR_CONS.COD_USR_SOLICITO} = " & Stg_cod_Usuario & ""
           
        Else
            MDI_Administrador_contable.CrystalReport1.ReportFileName = STG_PATH_REPORTES & "rpt_arc_traductor.rpt"
            MDI_Administrador_contable.CrystalReport1.WindowState = crptMaximized
            MDI_Administrador_contable.CrystalReport1.SelectionFormula = "{TBL_REGISTRO_TRADUCTOR_RPT.COD_USR_SOLICITO} = " & Stg_cod_Usuario & ""
        End If
    End If


    If Val(Mid(Me.cmb_tipo, 1, 2)) = 2 Then
        MDI_Administrador_contable.CrystalReport1.ReportFileName = STG_PATH_REPORTES & "rpt_arc_concilia.rpt"
        MDI_Administrador_contable.CrystalReport1.WindowState = crptMaximized
        MDI_Administrador_contable.CrystalReport1.SelectionFormula = "{TBL_CONCILIACION_RPT.COD_USR_SOLICITO} = " & Stg_cod_Usuario & ""
    End If

    If Val(Mid(Me.cmb_tipo, 1, 2)) = 3 Then
        MDI_Administrador_contable.CrystalReport1.ReportFileName = STG_PATH_REPORTES & "rpt_arc_trad.rpt"
        MDI_Administrador_contable.CrystalReport1.WindowState = crptMaximized
        MDI_Administrador_contable.CrystalReport1.SelectionFormula = "{TBL_REGISTRO_TRADUCTOR_CONS.COD_USR_SOLICITO} = " & Stg_cod_Usuario & ""
    End If


    If Val(Mid(Me.cmb_tipo, 1, 2)) = 4 Then
        MDI_Administrador_contable.CrystalReport1.ReportFileName = STG_PATH_REPORTES & "rpt_arc_rechazo.rpt"
        MDI_Administrador_contable.CrystalReport1.WindowState = crptMaximized
        MDI_Administrador_contable.CrystalReport1.SelectionFormula = "{TBL_REGISTRO_TRADUCTOR_RPT.COD_USR_SOLICITO} = " & Stg_cod_Usuario & ""
    End If

    If Val(Mid(Me.cmb_tipo, 1, 2)) = 5 Then
        MDI_Administrador_contable.CrystalReport1.ReportFileName = STG_PATH_REPORTES & "rpt_arc_trad.rpt"
        MDI_Administrador_contable.CrystalReport1.WindowState = crptMaximized
        MDI_Administrador_contable.CrystalReport1.SelectionFormula = "{TBL_REGISTRO_TRADUCTOR_CONS.COD_USR_SOLICITO}  = " & Stg_cod_Usuario & ""
    End If

    If Val(Mid(Me.cmb_tipo, 1, 2)) = 6 Then
        MDI_Administrador_contable.CrystalReport1.ReportFileName = STG_PATH_REPORTES & "rpt_arc_Bana.rpt"
        MDI_Administrador_contable.CrystalReport1.WindowState = crptMaximized
        MDI_Administrador_contable.CrystalReport1.SelectionFormula = "{TBL_REGISTRO_BANA_RPT.COD_USR_SOLICITO} = '" & Stg_cod_Usuario & "'"
    End If
    MDI_Administrador_contable.CrystalReport1.Action = 1
    MDI_Administrador_contable.CrystalReport1.Reset

    MousePointer = vbDefault
    cmd_Imprimir.Enabled = False

    Exit Sub

error:
    MsgBox Err.Number & " " & Err.Description & "  Proceso Cancelado."
    MousePointer = 0

End Sub
Private Sub cmd_salir_Click()
    On Error Resume Next
    Unload Me
End Sub


Private Sub dbg_consulta_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Set RSOBJp = Nothing

    cargoTotales = False
    cargoErrores = False


    Select Case Val(Mid(Me.cmb_tipo, 1, 2))
    Case 1, 3, 4, 5


        If ING_BD_CONSULTAR = True Then
            sqlTablas = "  VTA_TRADUCTOR_1 A,TBL_CAMPO_TRANSACCION B "
        Else
            sqlTablas = " USRLARC.VTA_TRADUCTOR_1@ARCHIST  A,TBL_CAMPO_TRANSACCION B "
        End If

        'sentencia = "SELECT  cod_campo, valor_campo FROM  TBL_CAMPOS_VALOR_TRAD WHERE  FECHA_SISTEMA = TO_DATE('" & dbg_consulta.Columns(3) & "','DD/MM/YYYY') AND COD_TIPO_REGISTRO = " & dbg_consulta.Columns(9) & "  AND VALOR_CAMPO <> 0 AND SEC_REGISTRO = " & dbg_consulta.Columns(10) & " ORDER BY COD_CAMPO "
        sentencia = "SELECT  A.cod_campo, DESC_CAMPO,valor_campo FROM   " & sqlTablas & _
                  " where A.COD_ENTIDAD   = B.COD_ENTIDAD (+) AND A.APLICACION_FUENTE = B.COD_APLICACION_FUENTE (+) AND A.COD_TRANSACCION = B.COD_TRANSACCION (+) " & _
                    "AND A.COD_CAMPO = B.COD_CAMPO (+)  AND FECHA_SISTEMA = TO_DATE('" & dbg_consulta.Columns(3) & "','DD/MM/YYYY') AND COD_TIPO_REGISTRO = " & dbg_consulta.Columns(9) & "   AND SEC_REGISTRO = " & dbg_consulta.Columns(10) & sqlCamposGrilla & " ORDER BY COD_CAMPO "
        ' AND VALOR_CAMPO <> 0
        'MsgBox sentencia
        Set RSOBJp = cargarRecordSet(sentencia)
    Case 2
        If ING_BD_CONSULTAR = True Then
            sqlTablas = "  VTA_REG_CONC A,TBL_CAMPO_TRANSACCION B "
        Else
            sqlTablas = " USRLARC.VTA_REG_CONC@ARCHIST  A,TBL_CAMPO_TRANSACCION B "
        End If

        'sentencia = "SELECT  cod_campo, valor_campo FROM  TBL_CAMPOS_VALOR_CONC WHERE  FECHA_SISTEMA = TO_DATE('" & dbg_consulta.Columns(5) & "','DD/MM/YYYY')  AND VALOR_CAMPO <> 0 AND SEC_REGISTRO = " & dbg_consulta.Columns(10) & " ORDER BY COD_CAMPO "
        sentencia = "SELECT  A.cod_campo, DESC_CAMPO,valor_campo FROM  " & sqlTablas & _
                  " where A.COD_ENTIDAD   = B.COD_ENTIDAD (+) AND A.APLICACION_FUENTE = B.COD_APLICACION_FUENTE (+) AND A.COD_TRANSACCION = B.COD_TRANSACCION (+) " & _
                    "AND A.COD_CAMPO = B.COD_CAMPO (+)  AND FECHA_SISTEMA = TO_DATE('" & dbg_consulta.Columns(5) & "','DD/MM/YYYY') AND VALOR_CAMPO <> 0 AND SEC_REGISTRO = " & dbg_consulta.Columns(10) & " ORDER BY COD_CAMPO "
        'MsgBox sentencia
        Set RSOBJp = cargarRecordSet(sentencia)
    Case 6
        Set RSOBJp = Nothing
    End Select

    If Val(Mid(Me.cmb_tipo, 1, 2)) <> 6 Then
        Me.sst_vlr.Visible = True
        Me.sst_vlr.Tab = 0
        If RSOBJp.EOF = True Then
            MsgBox "No existen valores asociados al registro"
            Set RSOBJp = Nothing
            Set Me.dbg_valor.DataSource = RSOBJp
        Else
            Set Me.dbg_valor.DataSource = RSOBJp
            dbg_valor.Columns(0).Width = 650
            dbg_valor.Columns(1).Width = 6000
            dbg_valor.Columns(2).Width = 2000

            dbg_valor.Columns(0).Visible = True
            dbg_valor.Columns(1).Visible = True
            dbg_valor.Columns(2).Visible = True

            dbg_valor.Columns(0).NumberFormat = ("0#")
            dbg_valor.Columns(2).NumberFormat = ("###,###,###,###.00")

            dbg_valor.Columns(0).Caption = "Campo"
            dbg_valor.Columns(1).Caption = "Descipción"
            dbg_valor.Columns(2).Caption = "Valor"

            dbg_valor.Columns(0).Locked = True
            dbg_valor.Columns(1).Locked = True
            dbg_valor.Columns(2).Locked = True
            'dbg_valor.Columns(0).Button = True

            dbg_valor.Columns(0).Alignment = dbgCenter
            dbg_valor.Columns(1).Alignment = dbgLeft
            dbg_valor.Columns(2).Alignment = dbgRight
        End If

    End If
End Sub





Private Sub DBG_DETALLE_ButtonClick(ByVal ColIndex As Integer)
    On Error GoTo error
    cmd_Imprimir.Enabled = False
    If ColIndex = 2 Then
        Me.txt_desc.Text = Me.DBG_DETALLE.Columns(14).Text
        Me.frm_desc.Visible = True
    End If
    Exit Sub
error:
    If Err.Number = 9 Then
        MsgBox "No Existe Descripción para esta Transacción"
    Else
        MsgBox Err.Number & " . " & Err.Description
    End If
End Sub

Private Sub DBG_DETALLE_Click()
    Me.dtg_valorConc.ClearFields
    DBG_DETALLE_KeyUp 0, 0
End Sub

Sub pl_carga_conc()
    On Error Resume Next
    Dim num As Integer
    Dim transaccion As String

    num = 0
    transaccion = ""
    
    'Puntero de procesamiento
    MousePointer = 13

    Set RSOBJ3 = Nothing
    
    'Construimos la condicion de transaccion dependiendo de las filas de detalle que se tengan.
    While num < Me.DBG_DETALLE.ApproxCount
        
        'Seleccionamos la fila usando la variable num
        Me.DBG_DETALLE.Row = num
        
        'Si no es la primera fila, agregamos la condicion con OR, lo que quiere decir que agregamos mas transacciones a la condicion.
        If Me.DBG_DETALLE.ApproxCount > 1 And num <> 0 Then
            transaccion = transaccion & " OR (A.COD_TRANSACCION  = '" & Format(DBG_DETALLE.Columns(2), "0###") & "' AND APLICACION_FUENTE = '" & DBG_DETALLE.Columns(1) & "')"
        Else
            transaccion = " AND ( (A.COD_TRANSACCION  = '" & Format(DBG_DETALLE.Columns(2), "0###") & "' AND APLICACION_FUENTE = '" & DBG_DETALLE.Columns(1) & "')"
        End If
        num = num + 1
    Wend
    
    'Si se trata de una conciliacion, llenamos el detalle.
    If dbg_consulta.Columns(0) = "CONC" Then

        'Consultamos base datos en linea
        If ING_BD_CONSULTAR = True Then
            sqlTablas = "  TBL_REGISTRO_TRADUCTOR A, TBL_TRANSACCION_TRADUCTOR B "
        Else
        'Consultamos base de datos historica
            sqlTablas = " USRLARC.TBL_REGISTRO_TRADUCTOR@ARCHIST A, TBL_TRANSACCION_TRADUCTOR B"
        End If


        Set RSOBJ3 = cargarRecordSet("SELECT  A.COD_ENTIDAD,A.APLICACION_FUENTE,A.COD_TRANSACCION,A.FECHA_MOVIMIENTO,A.CENTRO_ORIGEN,A.CENTRO_DESTINO,A.TIPO_CUENTA,A.COD_CUENTA,A.num_documento,A.FILLER,A.NUM_CAMPOS_MONETARIOS,A.TAM_CAMPOS_MONETARIOS, A.COD_TIPO_REGISTRO ,A.FECHA_SISTEMA, A.SEC_REGISTRO , B.DESC_TRANSACCION " & _
                                     "FROM  " & sqlTablas & _
                                   " WHERE  A.COD_ENTIDAD = B.COD_ENTIDAD AND A.APLICACION_FUENTE = B.COD_APLICACION_FUENTE  " & _
                                   " AND A.COD_TRANSACCION = B.COD_TRANSACCION  " & _
                                   " AND A.COD_ENTIDAD = " & Val(Me.DBG_DETALLE.Columns(0)) & transaccion & ")" & _
                                   " AND  FECHA_SISTEMA = TO_DATE('" & Format(DBG_DETALLE.Columns(12), "DD/MM/YYYY") & "','DD/MM/YYYY') " & _
                                   " AND  FECHA_MOVIMIENTO = TO_DATE('" & Format(DBG_DETALLE.Columns(3), "DD/MM/YYYY") & "','DD/MM/YYYY') " & _
                                   " AND CENTRO_ORIGEN = " & Val(DBG_DETALLE.Columns(4)) & _
                                   " AND CENTRO_DESTINO = " & Val(DBG_DETALLE.Columns(5)) & " AND A.COD_TIPO_REGISTRO = 1  " & _
                                   " ORDER BY COD_CUENTA ")

        Set rsdato = Nothing

        tipo_configuracion = 12

        'Si la consulta retorna resultados, llenamos el detalle.
        If Not RSOBJ3.EOF Then
            Set Me.dtg_detalle2.DataSource = RSOBJ3
            Configurar_grilla_Dat_consulta Me.dtg_detalle2
            dtg_detalle2_Click
        End If
    End If

    MousePointer = 0
End Sub

Private Sub DBG_DETALLE_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Set RSOBJp = Nothing
    Dim SQL As String


    If Val(Mid(Me.cmb_tipo, 1, 2)) = 2 Or Val(Mid(Me.cmb_tipo, 1, 2)) = 3 Or Val(Mid(Me.cmb_tipo, 1, 2)) = 5 Then

        If ING_BD_CONSULTAR = True Then
            sqlTablas = " TBL_CAMPOS_VALOR_CONC "
        Else
            sqlTablas = " USRLARC.TBL_CAMPOS_VALOR_CONC@ARCHIST "
        End If
        SQL = ""
    Else
        If ING_BD_CONSULTAR = True Then
            sqlTablas = " TBL_CAMPOS_VALOR_TRAD "
        Else
            sqlTablas = " USRLARC.TBL_CAMPOS_VALOR_TRAD@ARCHIST "
        End If
        SQL = " AND COD_TIPO_REGISTRO = 1 "

    End If

    sentencia = "SELECT  cod_campo, valor_campo FROM " & sqlTablas & " WHERE FECHA_SISTEMA = TO_DATE('" & DBG_DETALLE.Columns(12) & "','DD/MM/YYYY') " & SQL & "  AND SEC_REGISTRO = " & DBG_DETALLE.Columns(13)
    'MsgBox sentencia
    Set RSOBJp = cargarRecordSet(sentencia)
    Set rsdato1 = Nothing

    If RSOBJp.EOF = True Then
        MsgBox "No existen Valores asociados al registro"
    Else
        Set Me.dtg_valorConc.DataSource = RSOBJp

        dtg_valorConc.Columns(0).Width = 650
        dtg_valorConc.Columns(1).Width = 1950

        dtg_valorConc.Columns(0).NumberFormat = ("0#")
        dtg_valorConc.Columns(1).NumberFormat = ("###,###,###,###.00")

        dtg_valorConc.Columns(0).Caption = "Id.campo"
        dtg_valorConc.Columns(1).Caption = "Valor"

        dtg_valorConc.Columns(0).Locked = True
        dtg_valorConc.Columns(1).Locked = True

        dtg_valorConc.Columns(0).Alignment = dbgCenter
        dtg_valorConc.Columns(1).Alignment = dbgRight
        dtg_valorConc.Columns(0).Button = True
    End If
End Sub

Private Sub DBG_DETALLE_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.frm_desc.Visible = False
End Sub

Private Sub dtg_detalle2_ButtonClick(ByVal ColIndex As Integer)

    On Error Resume Next

    If ColIndex = 2 Then
        Me.txt_desc.Text = Me.dtg_detalle2.Columns(15).Text
        Me.frm_desc.Visible = True
    End If

End Sub

Private Sub dtg_detalle2_Click()
    Me.dtg_valor_det.ClearFields
    dtg_detalle2_KeyUp 0, 0
End Sub

Private Sub dtg_detalle2_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Set RSOBJp = Nothing


    If Val(Mid(Me.cmb_tipo, 1, 2)) = 2 And Me.tb_detalle.Visible = False Then

        If ING_BD_CONSULTAR = True Then
            sqlTablas = " TBL_CAMPOS_VALOR_CONC "
        Else
            sqlTablas = " USRLARC.TBL_CAMPOS_VALOR_CONC@ARCHIST "
        End If

        sentencia = "SELECT   cod_campo, valor_campo FROM " & sqlTablas & " WHERE FECHA_SISTEMA = TO_DATE('" & dtg_detalle2.Columns(13) & "','DD/MM/YYYY')  AND VALOR_CAMPO <> 0  AND SEC_REGISTRO = " & dtg_detalle2.Columns(14)

    Else
        If ING_BD_CONSULTAR = True Then
            sqlTablas = " TBL_CAMPOS_VALOR_TRAD "
        Else
            sqlTablas = " USRLARC.TBL_CAMPOS_VALOR_TRAD@ARCHIST "
        End If

        sentencia = "SELECT   cod_campo, valor_campo FROM " & sqlTablas & " WHERE FECHA_SISTEMA = TO_DATE('" & dtg_detalle2.Columns(13) & "','DD/MM/YYYY') AND COD_TIPO_REGISTRO = 1  AND VALOR_CAMPO <> 0  AND SEC_REGISTRO = " & dtg_detalle2.Columns(14)

    End If
    'MsgBox sentencia

    Set RSOBJp = cargarRecordSet(sentencia)
    If RSOBJp.EOF = True Then
        MsgBox "No existen Valores asociados al registtro"
    Else
        Set Me.dtg_valor_det.DataSource = RSOBJp

        dtg_valor_det.Columns(0).Width = 700
        dtg_valor_det.Columns(1).Width = 1900

        dtg_valor_det.Columns(0).NumberFormat = ("0#")
        dtg_valor_det.Columns(1).NumberFormat = ("###,###,###,###.00")

        dtg_valor_det.Columns(0).Caption = "Id.campo"
        dtg_valor_det.Columns(1).Caption = "Valor"

        dtg_valor_det.Columns(0).Locked = True
        dtg_valor_det.Columns(1).Locked = True

        dtg_valor_det.Columns(0).Alignment = dbgCenter
        dtg_valor_det.Columns(1).Alignment = dbgRight
        dtg_valor_det.Columns(0).Button = True

    End If

    Set RSOBJ3 = Nothing


End Sub

Private Sub dtg_detalle2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.frm_desc.Visible = False
End Sub

Private Sub dtg_valorConc_ButtonClick(ByVal ColIndex As Integer)
    On Error GoTo error

    If Val(Mid(Me.cmb_tipo, 1, 2)) = 2 Then

        If ING_BD_CONSULTAR = True Then
            sqlTablas = " TBL_CAMPO_TRANSACCION "
        Else
            sqlTablas = " USRLARC.TBL_CAMPO_TRANSACCION@ARCHIST "
        End If

        sentencia = "SELECT   DESC_CAMPO FROM " & sqlTablas & " WHERE COD_ENTIDAD = 1 AND COD_APLICACION_FUENTE =  '" & DBG_DETALLE.Columns(1) & "'  AND COD_TRANSACCION =  '" & DBG_DETALLE.Columns(2) & "' AND COD_CAMPO = " & dtg_valorConc.Columns(0)

    Else
        If ING_BD_CONSULTAR = True Then
            sqlTablas = " TBL_CAMPO_TRANSACCION "
        Else
            sqlTablas = " USRLARC.TBL_TBL_CAMPO_TRANSACCION@ARCHIST "
        End If

        sentencia = "SELECT   DESC_CAMPO FROM " & sqlTablas & " WHERE COD_ENTIDAD = 1 AND COD_APLICACION_FUENTE =  '" & DBG_DETALLE.Columns(1) & "'  AND COD_TRANSACCION =  '" & DBG_DETALLE.Columns(2) & "' AND COD_CAMPO = " & dtg_valorConc.Columns(0)

    End If

    'MsgBox sentencia

    Set rsobj = cargarRecordSet(sentencia)
    If rsobj.EOF Then
        MsgBox "No existe descripción para el campo " & dtg_valorConc.Columns(0) & " de la transacción " & DBG_DETALLE.Columns(1) & "-" & rellenar(DBG_DETALLE.Columns(2), 4, "0", "izquierda")
    Else
        Me.txt_desc.Text = UCase(rsobj(0))
    End If
    If Not rsobj Is Nothing Then
        If rsobj.State = adStateOpen Then rsobj.Close
    End If
    Set rsobj = Nothing
    Me.frm_desc.Visible = True
    Exit Sub
error:
    If Err.Number = 9 Then
        MsgBox "No existe descripción para el campo " & dtg_valor.Columns(0) & " de la transacción " & dbg_consulta.Columns(0) & "-" & rellenar(dbg_consulta.Columns(1), 4, "0", "izquierda")
    Else
        MsgBox Err.Number & " . " & Err.Description
    End If
End Sub


Private Sub Form_Load()
    Set rsobj = Nothing
    'Carga el combo de tipo de movimiento
    pl_cargar_tipo
    'Carga el combo de aplicacion fuente
    If Stg_Perfil_Usuario_Acceso = 8 Then
        pl_cargar_aplicacion (True)
    Else
        pl_cargar_aplicacion (False)
    End If
    'Carga el combo de tipo de cuenta
    pl_cargar_tipo_cuenta

    'Visualizar el valor de la fecha contable por defecto
    Me.msk_Fecha_Proceso.Text = Format(Dtg_Fecha_movimiento, "DD/MM/YYYY")

    'Se configura la posicion y tamaño de los paneles
    configurarPaneles
    frameParametros.Visible = True


    'Activacion de la caja de chequeo del filtro de transacciones
    'Me.chkFiltro.Value = 1
    Me.Cbo_filtro.ListIndex = 0
    If (Stg_Perfil_Usuario_Acceso = 8) Then
        'Me.chkFiltro.Enabled = False
        Me.Cbo_filtro.Enabled = False
    Else
        'Me.chkFiltro.Enabled = True
        Me.Cbo_filtro.Enabled = True
    End If
    
   'Requerimiento CVAPD00223966.
   'ABOCANE. Mayo 2015
   'Asignacion de valores para el libro auxiliar
    Me.cmb_libro_auxiliar.Clear
    Me.cmb_libro_auxiliar.AddItem "COLGAAP"
    Me.cmb_libro_auxiliar.AddItem "IFRS"
    Me.cmb_libro_auxiliar.ListIndex = 1
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PL_CERRAR_CONEXIONES
    PL_CERRAR_CONEXIONES1
End Sub

Private Sub Image1_Click()
    PL_CERRAR_CONEXIONES
    PL_CERRAR_CONEXIONES1
End Sub
Private Sub Image2_Click()
    pl_tipo_consulta
End Sub



Private Sub lbl_consultar_Click()
    PL_CERRAR_CONEXIONES
    PL_CERRAR_CONEXIONES1
    pl_tipo_consulta
    Cmd_Consultar.Enabled = True
End Sub





Private Sub msk_cta_aux_GotFocus()
    msk_cta_aux.SelStart = 0
    msk_cta_aux.SelLength = Len(msk_cta_aux)
End Sub

Private Sub msk_cta_aux_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        'Si la tecla es ESCAPE debe limpiar el contenido del control
    Case vbKeyEscape
        msk_cta_aux.Text = ""
        KeyAscii = 0
        'Si la tecla es ENTER debe navegar al siguiente control
    Case vbKeyReturn
        'SendKeys "+{tab}"
        SendKeys "{tab}"
        KeyAscii = 0
    End Select
End Sub
Private Sub msk_cuenta_GotFocus()
    msk_cuenta.SelStart = 0
    msk_cuenta.SelLength = Len(msk_cuenta)
End Sub
Private Sub msk_cuenta_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        'Si la tecla es ESCAPE debe limpiar el contenido del control
    Case vbKeyEscape
        msk_cuenta.Text = "                "
        KeyAscii = 0

        'Si la tecla es ENTER debe navegar al siguiente control
    Case vbKeyReturn
        'SendKeys "+{tab}"
        SendKeys "{tab}"
        KeyAscii = 0
    End Select
End Sub




Private Sub msk_destino_GotFocus()
    msk_destino.SelStart = 0
    msk_destino.SelLength = Len(msk_destino)
End Sub
Private Sub msk_destino_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        'Si la tecla es ESCAPE debe limpiar el contenido del control
    Case vbKeyEscape
        msk_destino.Text = "    ,    ,    ,    "
        KeyAscii = 0

        'Si la tecla es ENTER debe navegar al siguiente control
    Case vbKeyReturn
        'SendKeys "+{tab}"
        SendKeys "{tab}"
        KeyAscii = 0
    End Select
End Sub

Private Sub msk_destino1_GotFocus()
    msk_destino1.SelStart = 0
    msk_destino1.SelLength = Len(msk_destino1)
End Sub


Private Sub msk_entidad_GotFocus()
    msk_entidad.SelStart = 0
    msk_entidad.SelLength = Len(msk_entidad)
End Sub
Private Sub msk_entidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        msk_entidad.Text = "  "
    Else
        If KeyAscii = vbKeyReturn Then
            Me.msk_transaccion.SetFocus
        End If
    End If
End Sub

Private Sub msk_fecha_fin_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        msk_fecha_fin.Text = "  /  /    "
    Else
        If KeyAscii = vbKeyReturn Then
            Me.cmb_tipo_cuenta.SetFocus
        End If
    End If
End Sub

Private Sub msk_Fecha_ini_GotFocus()
    msk_Fecha_ini.SelStart = 0
    msk_Fecha_ini.SelLength = Len(msk_Fecha_ini)
End Sub
Private Sub msk_fecha_ini_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        'Si la tecla es ESCAPE debe limpiar el contenido del control
    Case vbKeyEscape
        msk_Fecha_ini.Text = "  /  /    "
        KeyAscii = 0

        'Si la tecla es ENTER debe navegar al siguiente control
    Case vbKeyReturn
        'SendKeys "+{tab}"
        SendKeys "{tab}"
        KeyAscii = 0
    End Select
End Sub


Private Sub msk_Fecha_Proceso_GotFocus()
    msk_Fecha_Proceso.SelStart = 0
    msk_Fecha_Proceso.SelLength = Len(msk_Fecha_Proceso)
End Sub

Private Sub msk_Fecha_Proceso_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        'Si la tecla es ESCAPE debe limpiar el contenido del control
    Case vbKeyEscape
        msk_Fecha_Proceso.Text = "  /  /    "
        KeyAscii = 0

        'Si la tecla es ENTER debe navegar al siguiente control
    Case vbKeyReturn
        'SendKeys "+{tab}"
        'SendKeys "{tab}"
        KeyAscii = 0
    End Select
End Sub

Private Sub msk_origen_GotFocus()
    msk_origen.SelStart = 0
    msk_origen.SelLength = Len(msk_origen)
End Sub
Private Sub msk_origen_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        'Si la tecla es ESCAPE debe limpiar el contenido del control
    Case vbKeyEscape
        msk_origen.Text = "    ,    ,    ,    "
        KeyAscii = 0

        'Si la tecla es ENTER debe navegar al siguiente control
    Case vbKeyReturn
        'SendKeys "+{tab}"
        SendKeys "{tab}"
        KeyAscii = 0
    End Select
End Sub





Private Sub msk_origen1_GotFocus()
    msk_origen1.SelStart = 0
    msk_origen1.SelLength = Len(msk_origen1)
End Sub

Private Sub msk_referencia_GotFocus()
    msk_referencia.SelStart = 0
    msk_referencia.SelLength = Len(msk_referencia)
End Sub
Private Sub msk_referencia_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        msk_referencia.Text = "    "
    Else
        If KeyAscii = vbKeyReturn Then
            Me.msk_secuencia.SetFocus
        End If
    End If
End Sub
Private Sub msk_secuencia_GotFocus()
    msk_secuencia.SelStart = 0
    msk_secuencia.SelLength = Len(msk_secuencia)
End Sub
Private Sub msk_secuencia_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        'Si la tecla es ESCAPE debe limpiar el contenido del control
    Case vbKeyEscape
        msk_secuencia.Text = "  "
        KeyAscii = 0

        'Si la tecla es ENTER debe navegar al siguiente control
    Case vbKeyReturn
        'SendKeys "+{tab}"
        SendKeys "{tab}"
        KeyAscii = 0
    End Select
End Sub


Private Sub msk_transaccion_concil_GotFocus()
    msk_transaccion_concil.SelStart = 0
    msk_transaccion_concil.SelLength = Len(msk_transaccion_concil)
End Sub
Private Sub msk_transaccion_concil_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        msk_transaccion_concil.Text = "  "
    Else
        If KeyAscii = vbKeyReturn Then
            Me.cmb_aplicacion.SetFocus
        End If
    End If
End Sub
Private Sub msk_transaccion_GotFocus()
    msk_transaccion.SelStart = 0
    msk_transaccion.SelLength = Len(msk_transaccion)
End Sub
Private Sub msk_transaccion_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))    ' Mayuscula
    Select Case KeyAscii
        ' Si la tecla es ESCAPE debe limpiar el contenido del control
    Case vbKeyEscape
        msk_transaccion.Text = "    ,    ,    ,    "
        KeyAscii = 0

        ' Si la tecla es ENTER debe navegar al siguiente control
    Case vbKeyReturn
        'SendKeys "+{tab}"
        SendKeys "{tab}"
        KeyAscii = 0
    End Select


End Sub
Private Sub msk_valor_max_GotFocus()
    msk_valor_max.Text = Me.msk_valor_min
    msk_valor_max.SelStart = 0
    msk_valor_max.SelLength = Len(msk_valor_max)
End Sub
Private Sub msk_valor_max_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        ' Si la tecla es ESCAPE debe limpiar el contenido del control
    Case vbKeyEscape
        msk_valor_max.Text = ""
        KeyAscii = 0

        ' Si la tecla es ENTER debe navegar al siguiente control
    Case vbKeyReturn
        'SendKeys "+{tab}"
        SendKeys "{tab}"
        KeyAscii = 0
    End Select
End Sub
Private Sub msk_valor_min_GotFocus()
    msk_valor_min.SelStart = 0
    msk_valor_min.SelLength = Len(msk_valor_min)
End Sub
Private Sub msk_valor_min_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        'Si la tecla es ESCAPE debe limpiar el contenido del control
    Case vbKeyEscape
        msk_valor_min.Text = ""
        KeyAscii = 0

        'Si la tecla es ENTER debe navegar al siguiente control
    Case vbKeyReturn
        'SendKeys "+{tab}"
        SendKeys "{tab}"
        KeyAscii = 0
    End Select
End Sub

Private Sub sst_vlr_Click(PreviousTab As Integer)
    Dim rsobjp1 As ADODB.Recordset



    'Cambia a cursor en espera    Me.MousePointer = vbHourglass

    'Evalua si el ususario hizo clic en el tab "Totales"
    Select Case Me.sst_vlr.Tab
    Case 0
        dbg_valor.Visible = True
        dtg_suma.Visible = False
        dtg_error.Visible = False
    Case 1
        dbg_valor.Visible = False
        dtg_suma.Visible = True
        dtg_error.Visible = False
    Case 2
        dbg_valor.Visible = False
        dtg_suma.Visible = False
        dtg_error.Visible = True
    End Select


    'Evalua si el ususario hizo clic en el tab "Totales"
    If Me.sst_vlr.Tab = 1 And cargoTotales = False Then

        Set rsobjp1 = Nothing
        'Construye la sentencia SQL necesaria para extraer los totales
        If (Me.Cbo_filtro.ListIndex = 0) Then
            Select Case Val(Mid(Me.cmb_tipo, 1, 2))
            Case 1, 3, 5
                'FROM_CONSULTA = "VTA_REG_TRAD_USUARIO A"

                If ING_BD_CONSULTAR = True Then
                    sqlTablas = "  VTA_REG_TRAD_USUARIO A,TBL_CAMPO_TRANSACCION B  "
                Else
                    sqlTablas = " USRLARC.VTA_REG_TRAD_USUARIO@ARCHIST  A,TBL_CAMPO_TRANSACCION B "
                End If
                sentencia = " SELECT  A.aplicacion_fuente, A.cod_transaccion, A.cod_campo,desc_campo, SUM(valor_campo) suma from  " & sqlTablas & _
                          " WHERE " & WHERE_CONSULTA1 & " AND  A.COD_ENTIDAD   = B.COD_ENTIDAD (+)  AND A.APLICACION_FUENTE  =  B.COD_APLICACION_FUENTE  (+) " & _
                          " AND A.COD_TRANSACCION  =  B.COD_TRANSACCION (+)  AND A.COD_CAMPO = B.COD_CAMPO (+) " & _
                          " AND fecha_sistema = TO_DATE('" & dbg_consulta.Columns(3) & "','DD/MM/YYYY') " & _
                          " GROUP BY a.aplicacion_fuente, a.cod_transaccion, a.cod_campo,desc_campo ORDER BY a.aplicacion_fuente, a.cod_transaccion, a.cod_campo"
                Set rsobjp1 = cargarRecordSet(sentencia)

            Case 2

                If ING_BD_CONSULTAR = True Then
                    'sqlTablas = "  VTA_REG_TRAD_USUARIO A,TBL_CAMPO_TRANSACCION B "
                    sqlTablas = "  VTA_REG_CONC_USUARIO A,TBL_CAMPO_TRANSACCION B "
                Else
                    'sqlTablas = " USRLARC.VTA_REG_TRAD_USUARIO@ARCHIST  A,TBL_CAMPO_TRANSACCION B "
                    sqlTablas = " USRLARC.VTA_REG_CONC_USUARIO@ARCHIST  A,TBL_CAMPO_TRANSACCION B "
                End If

                sentencia = " SELECT  A.aplicacion_fuente, A.cod_transaccion, A.cod_campo,desc_campo, SUM(valor_campo) suma from " & sqlTablas & _
                          " WHERE " & WHERE_CONSULTA & " AND  A.COD_ENTIDAD   = B.COD_ENTIDAD (+)  AND A.APLICACION_FUENTE  =  B.COD_APLICACION_FUENTE  (+) " & _
                          " AND A.COD_TRANSACCION  =  B.COD_TRANSACCION (+)  AND A.COD_CAMPO = B.COD_CAMPO " & _
                          " AND fecha_sistema = TO_DATE('" & dbg_consulta.Columns(5) & "','DD/MM/YYYY') " & _
                          " GROUP BY a.aplicacion_fuente, a.cod_transaccion, a.cod_campo,desc_campo "
                ' MsgBox sentencia

                Set rsobjp1 = cargarRecordSet(sentencia)
            Case 4
                'FROM_CONSULTA = "VTA_REG_TRAD_USUARIO A"
                If ING_BD_CONSULTAR = True Then
                    sqlTablas = "  VTA_REG_TRAD_USUARIO  A,TBL_CAMPO_TRANSACCION B "
                Else
                    sqlTablas = " USRLARC.VTA_REG_TRAD_USUARIO@ARCHIST  A,TBL_CAMPO_TRANSACCION B "
                End If

                sentencia = " SELECT  A.aplicacion_fuente, A.cod_transaccion, A.cod_campo,desc_campo, SUM(valor_campo) suma from  " & sqlTablas & _
                          " WHERE " & WHERE_CONSULTA & " AND  A.COD_ENTIDAD   = B.COD_ENTIDAD (+)  AND A.APLICACION_FUENTE  =  B.COD_APLICACION_FUENTE  (+) " & _
                          " AND A.COD_TRANSACCION  =  B.COD_TRANSACCION (+)  AND A.COD_CAMPO = B.COD_CAMPO (+) " & _
                          " AND fecha_sistema = TO_DATE('" & dbg_consulta.Columns(3) & "','DD/MM/YYYY') " & _
                          " GROUP BY a.aplicacion_fuente, a.cod_transaccion, a.cod_campo,desc_campo ORDER BY a.aplicacion_fuente, a.cod_transaccion, a.cod_campo"
                Set rsobjp1 = cargarRecordSet(sentencia)
            End Select
        Else
            Select Case Val(Mid(Me.cmb_tipo, 1, 2))
            Case 1, 3, 5

                If ING_BD_CONSULTAR = True Then
                    sqlTablas = "  VTA_TRADUCTOR_1 A,TBL_CAMPO_TRANSACCION B "
                Else
                    sqlTablas = " USRLARC.VTA_TRADUCTOR_1@ARCHIST  A,TBL_CAMPO_TRANSACCION B"
                End If

                sentencia = " SELECT  A.aplicacion_fuente, A.cod_transaccion, A.cod_campo,desc_campo, SUM(valor_campo) suma from   " & sqlTablas & _
                          " WHERE " & WHERE_CONSULTA & " AND  A.COD_ENTIDAD   = B.COD_ENTIDAD (+)  AND A.APLICACION_FUENTE  =  B.COD_APLICACION_FUENTE  (+) " & _
                          " AND A.COD_TRANSACCION  =  B.COD_TRANSACCION (+)  AND A.COD_CAMPO = B.COD_CAMPO (+) " & _
                          " AND fecha_sistema = TO_DATE('" & dbg_consulta.Columns(3) & "','DD/MM/YYYY') " & _
                          " GROUP BY a.aplicacion_fuente, a.cod_transaccion, a.cod_campo,desc_campo ORDER BY a.aplicacion_fuente, a.cod_transaccion, a.cod_campo"
                Set rsobjp1 = cargarRecordSet(sentencia)

            Case 2
                If ING_BD_CONSULTAR = True Then
                    sqlTablas = "  VTA_TRADUCTOR_1 A,TBL_CAMPO_TRANSACCION B "
                Else
                    sqlTablas = " USRLARC.VTA_TRADUCTOR_1@ARCHIST  A,TBL_CAMPO_TRANSACCION B"
                End If

                sentencia = " SELECT  A.aplicacion_fuente, A.cod_transaccion, A.cod_campo,desc_campo, SUM(valor_campo) suma from   " & sqlTablas & _
                          " WHERE " & WHERE_CONSULTA1 & " AND  A.COD_ENTIDAD   = B.COD_ENTIDAD (+)  AND A.APLICACION_FUENTE  =  B.COD_APLICACION_FUENTE  (+) " & _
                          " AND A.COD_TRANSACCION  =  B.COD_TRANSACCION (+)  AND A.COD_CAMPO = B.COD_CAMPO " & _
                          " AND fecha_sistema = TO_DATE('" & dbg_consulta.Columns(5) & "','DD/MM/YYYY') " & _
                          " GROUP BY a.aplicacion_fuente, a.cod_transaccion, a.cod_campo,desc_campo "

                Set rsobjp1 = cargarRecordSet(sentencia)
            Case 4
                'FROM_CONSULTA = "VTA_REG_TRAD_USUARIO A"
                If ING_BD_CONSULTAR = True Then
                    sqlTablas = "  VTA_TRADUCTOR_1  A,TBL_CAMPO_TRANSACCION B "
                Else
                    sqlTablas = " USRLARC.VTA_TRADUCTOR_1@ARCHIST  A,TBL_CAMPO_TRANSACCION B"
                End If
                sentencia = " SELECT  A.aplicacion_fuente, A.cod_transaccion, A.cod_campo,desc_campo, SUM(valor_campo) suma from   " & sqlTablas & _
                          " WHERE " & WHERE_CONSULTA1 & " AND  A.COD_ENTIDAD   = B.COD_ENTIDAD (+)  AND A.APLICACION_FUENTE  =  B.COD_APLICACION_FUENTE  (+) " & _
                          " AND A.COD_TRANSACCION  =  B.COD_TRANSACCION (+)  AND A.COD_CAMPO = B.COD_CAMPO (+) " & _
                          " AND fecha_sistema = TO_DATE('" & dbg_consulta.Columns(3) & "','DD/MM/YYYY') " & _
                          " GROUP BY a.aplicacion_fuente, a.cod_transaccion, a.cod_campo,desc_campo ORDER BY a.aplicacion_fuente, a.cod_transaccion, a.cod_campo"
                Set rsobjp1 = cargarRecordSet(sentencia)
            End Select

        End If

        'Configura la grilla de respuesta dependiendo del contenido de la consulta
        If rsobjp1.EOF = True Then
            MsgBox "No existen Valores asociados al registro"
        Else
            Set Me.dtg_suma.DataSource = rsobjp1
            cargoTotales = True
            'carga_conc = 1

            dtg_suma.Columns(0).Width = 1100
            dtg_suma.Columns(1).Width = 1200
            dtg_suma.Columns(2).Width = 700
            dtg_suma.Columns(3).Width = 4300
            dtg_suma.Columns(4).Width = 1900    'dtg_suma.Width - (dtg_suma.Columns(0).Width + dtg_suma.Columns(1).Width + dtg_suma.Columns(2).Width + 600)

            dtg_suma.Columns(2).NumberFormat = ("0#")
            dtg_suma.Columns(4).NumberFormat = ("###,###,###,###.00")

            dtg_suma.Columns(0).Caption = "Aplicación"
            dtg_suma.Columns(1).Caption = "Transacción"
            dtg_suma.Columns(2).Caption = "Campo"
            dtg_suma.Columns(3).Caption = "Descripcion"
            dtg_suma.Columns(4).Caption = "Valor"

            dtg_suma.Columns(0).Locked = True
            dtg_suma.Columns(1).Locked = True
            dtg_suma.Columns(2).Locked = True
            dtg_suma.Columns(3).Locked = True

            dtg_suma.Columns(0).Alignment = dbgCenter
            dtg_suma.Columns(1).Alignment = dbgCenter
            dtg_suma.Columns(2).Alignment = dbgCenter
            dtg_suma.Columns(3).Alignment = dbgLeft
            dtg_suma.Columns(4).Alignment = dbgRight
            dtg_suma.Refresh
        End If
    End If
    'Evalua si el ususario hizo clic en el tab "Errores"
    If Me.sst_vlr.Tab = 2 And cargoErrores = False Then
        Set rsobjp1 = Nothing
        'Construye la sentencia SQL necesaria para extraer los errores
        Select Case Val(Mid(Me.cmb_tipo, 1, 2))
        Case 4
            If ING_BD_CONSULTAR = True Then
                sqlTablas = "  TBL_ERROR E, TBL_ERROR_REGISTRO R "
            Else
                sqlTablas = " TBL_ERROR E, USRLARC.TBL_ERROR_REGISTRO@ARCHIST  R "
            End If

            Set rsobjp1 = cargarRecordSet(" SELECT  e.DESC_ERROR, e.ACCION_CORRECTIVA FROM  " & sqlTablas & _
                                        " WHERE e.cod_error = r.cod_error and fecha_sistema = TO_DATE('" & dbg_consulta.Columns(3) & "','DD/MM/YYYY') " & _
                                        " and SEC_REGISTRO= " & dbg_consulta.Columns(10) & _
                                        " Order by R.cod_error ")
        End Select

        'Configura la grilla de respuesta dependiendo del contenido de la consulta
        If rsobjp1.EOF = True Then
            MsgBox "No existen errores en el registro"
        Else
            cargoErrores = True
            Set Me.dtg_error.DataSource = rsobjp1

            dtg_error.Columns(0).Width = 10000
            dtg_error.Columns(1).Width = 30000

            dtg_error.Columns(0).Caption = "Error"
            dtg_error.Columns(1).Caption = "Acción correctiva"


            dtg_error.Columns(0).Locked = True
            dtg_error.Columns(1).Locked = True

            dtg_error.Columns(0).Alignment = dbgLeft
            dtg_error.Columns(1).Alignment = dbgLeft
            dtg_error.Refresh
        End If
    End If

    Me.MousePointer = 0
End Sub

' REQ CVAPD00223966
' ABOCANE Mayo 2016
' Creación de nuevo tab para información IFRS
Private Sub tb_detalle_Click(PreviousTab As Integer)
    
    Dim tipoSeleccionado As String
    
    'Almaceno temporalmente el tipo de movimiento seleccionado por el usuario
    tipoSeleccionado = Val(Mid(Me.cmb_tipo, 1, 2))
    
    'Si se trata de movimiento inconsistente, no se muestra nada en el detalle
    If tipoSeleccionado = TR_INCONSISTENTES Then
        Exit Sub
    End If
    
    If Me.tb_detalle.Caption = "Transacciones Detalladas" And carga_conc = 0 Then
        Set Me.dtg_detalle2.DataSource = Nothing
        pl_carga_conc
        carga_conc = 1
    End If
    If Me.tb_detalle.Caption = "Asientos COLGAAP" And carga_cont = 0 Then
        Set Me.dbg_conta.DataSource = Nothing
        pl_carga_Conta ("1")
        carga_cont = 1
    End If
    
     If Me.tb_detalle.Caption = "Asientos IFRS" And carga_cont_1 = 0 Then
        Set Me.dbg_ifrs.DataSource = Nothing
        pl_carga_Conta ("2")
        carga_cont_1 = 1
    End If

    If Me.tb_detalle.Caption = "Detalle Horizontal" And carga_det = 0 Then
        Set Me.dtg_detalle3.DataSource = Nothing
        pl_carga_det
        carga_det = 1

    End If
    If Me.tb_detalle.Caption = "Detalle Diferencias" And carga_diferencias = 0 Then
        Set Me.dtg_detalle4.DataSource = Nothing
        pl_carga_Diferencias
        carga_diferencias = 1

    End If

    If Me.tb_detalle.Caption = "Configuración Conciliación" Then
        Set Me.dtg_detalle5.DataSource = Nothing
        pl_carga_Config
    End If

End Sub
Sub pl_carga_det()
    On Error Resume Next
    Dim num As Integer
    Dim transaccion As String
    num = 0
    transaccion = ""

    MousePointer = 13

    Set RSOBJ3 = Nothing

    While num < Me.DBG_DETALLE.ApproxCount
        Me.DBG_DETALLE.Row = num
        If Me.DBG_DETALLE.ApproxCount > 1 And num <> 0 Then
            transaccion = transaccion & " OR (COD_TRANSACCION  = '" & Format(DBG_DETALLE.Columns(2), "0###") & "' AND APLICACION_FUENTE = '" & DBG_DETALLE.Columns(1) & "')"
        Else
            transaccion = " AND ( (COD_TRANSACCION  = '" & Format(DBG_DETALLE.Columns(2), "0###") & "' AND APLICACION_FUENTE = '" & DBG_DETALLE.Columns(1) & "')"
        End If
        num = num + 1
    Wend


    If dbg_consulta.Columns(0) = "CONC" Then
        PL_Conexion_Oracle
        If ING_BD_CONSULTAR = True Then
            sqlTablas = "  TBL_REGISTRO_TRADUCTOR "
        Else
            sqlTablas = " USRLARC.TBL_REGISTRO_TRADUCTOR@ARCHIST"
        End If

        sentencia = "DELETE  TBL_REGISTRO_TRADUCTOR_CONS WHERE COD_USR_SOLICITO = '" & Stg_cod_Usuario & "'"
        cnObj1.Execute (sentencia)

        sentencia = "INSERT INTO TBL_REGISTRO_TRADUCTOR_CONS" & _
                  " (FECHA_SISTEMA, SEC_REGISTRO,COD_USR_SOLICITO, COD_ENTIDAD,APLICACION_FUENTE,COD_TRANSACCION,FECHA_MOVIMIENTO,CENTRO_ORIGEN," & _
                  " CENTRO_DESTINO,TIPO_CUENTA,COD_CUENTA,num_documento,FILLER," & _
                  " NUM_CAMPOS_MONETARIOS,TAM_CAMPOS_MONETARIOS, COD_TIPO_REGISTRO , FLAG_AJUSTE)" & _
                  " SELECT FECHA_SISTEMA, SEC_REGISTRO," & Stg_cod_Usuario & ", COD_ENTIDAD,APLICACION_FUENTE,COD_TRANSACCION,FECHA_MOVIMIENTO,CENTRO_ORIGEN," & _
                  " CENTRO_DESTINO,TIPO_CUENTA,COD_CUENTA,num_documento,FILLER," & _
                  " NUM_CAMPOS_MONETARIOS,TAM_CAMPOS_MONETARIOS, COD_TIPO_REGISTRO, 0 " & _
                  " FROM " & sqlTablas & _
                  " WHERE COD_ENTIDAD = 1" & transaccion & ")" & _
                  " AND FECHA_SISTEMA = TO_DATE('" & Dtg_Fecha_movimiento1 & "','DD/MM/YYYY') " & _
                  " AND  FECHA_MOVIMIENTO = TO_DATE('" & Format(DBG_DETALLE.Columns(3), "DD/MM/YYYY") & "','DD/MM/YYYY') " & _
                  " AND CENTRO_ORIGEN = " & Val(DBG_DETALLE.Columns(4)) & _
                  " AND CENTRO_DESTINO = " & Val(DBG_DETALLE.Columns(5)) & " AND COD_TIPO_REGISTRO = 1  "
        cnObj1.Execute (sentencia)
        'MsgBox sentencia

        If ING_BD_CONSULTAR = True Then
            INL_BD_CONSULTAR = 1
        Else
            INL_BD_CONSULTAR = 0
        End If

        sentencia = "USRLARC.PL_CAMPOS_VALOR_TR('" & Stg_cod_Usuario & "'," & INL_BD_CONSULTAR & ") "

        'MsgBox sentencia
        cnObj1.Execute sentencia


        Set RSOBJ3 = cargarRecordSet("SELECT  COD_ENTIDAD,APLICACION_FUENTE,COD_TRANSACCION,FECHA_MOVIMIENTO,CENTRO_ORIGEN," & _
                                   " CENTRO_DESTINO,TIPO_CUENTA,COD_CUENTA,num_documento,FILLER," & _
                                   " NUM_CAMPOS_MONETARIOS,TAM_CAMPOS_MONETARIOS, COD_TIPO_REGISTRO,  " & _
                                   " COD_CAMPO1,VALOR1,COD_CAMPO2,VALOR2,COD_CAMPO3,VALOR3,COD_CAMPO4,VALOR4,COD_CAMPO5,VALOR5, " & _
                                   " COD_CAMPO6,VALOR6,COD_CAMPO7,VALOR7,COD_CAMPO8,VALOR8,COD_CAMPO9,VALOR9,COD_CAMPO10,VALOR10 " & _
                                   " FROM TBL_REGISTRO_TRADUCTOR_CONS WHERE COD_USR_SOLICITO = '" & Stg_cod_Usuario & "'" & _
                                   " order by COD_CAMPO1,VALOR1,COD_CAMPO2,VALOR2,COD_CAMPO3,VALOR3,COD_CAMPO4,VALOR4,COD_CUENTA,APLICACION_FUENTE  ")

        Set Me.dtg_detalle3.DataSource = RSOBJ3
        Set rsdato = Nothing
        tipo_configuracion = 7

        If RSOBJ3.EOF Then
            Set RSOBJ3 = Nothing
        Else
            Configurar_grilla_Dat_consulta Me.dtg_detalle3
        End If
    End If
    Set RSOBJ3 = Nothing
    MousePointer = 0
End Sub

Sub pl_carga_Config()
    MousePointer = 13
    Set RSOBJ3 = Nothing
    On Error Resume Next

    If dbg_consulta.Columns(0) = "CONC" Then
        Set RSOBJ3 = cargarRecordSet("SELECT  COD_APL_FUENTE_CONCILIACION,COD_TRANSACCION_CONCILIACION,SECUENCIA_CONCILIACION," & _
                                   " COD_APLICACION_FUENTE,COD_TRANSACCION,IDENTIFICADOR_CAMPO,COD_APLICACION_FUENTE_RELACION," & _
                                   " COD_TRANSACCION_RELACION,IDENTIFICADOR_CAMPO_RELACION,  " & _
                                   " VALOR_APLICADO_UNO,VALOR_APLICADO_DOS,VALOR_DIFERENCIA_UNO,VALOR_DIFERENCIA_DOS  " & _
                                   " FROM TBL_CONCILIACION_TRADUCTOR " & _
                                   " WHERE COD_APL_FUENTE_CONCILIACION   = 'CONC' AND COD_TRANSACCION_CONCILIACION = '" & dbg_consulta.Columns(1) & "' order by SECUENCIA_CONCILIACION  ")
        If RSOBJ3.EOF Then
            Set RSOBJ3 = Nothing
        Else
            Set Me.dtg_detalle5.DataSource = RSOBJ3
            tipo_configuracion = 20
            Configurar_grilla_Dat_consulta Me.dtg_detalle5
        End If
    End If
    Set RSOBJ3 = Nothing
    MousePointer = 0
End Sub
Sub pl_carga_Diferencias()
    On Error Resume Next
    Dim num As Integer
    Dim transaccion As String
    num = 0
    transaccion = ""
    MousePointer = 13
    Set RSOBJ3 = Nothing
    'If Me.DBG_DETALLE.DataMember <> "" Or Val(Mid(Me.cmb_tipo, 1, 2)) <> 2 Then
    If dbg_consulta.Columns(0) = "CONC" And Val(Mid(Me.cmb_tipo, 1, 2)) = 5 Then
        While num < Me.DBG_DETALLE.ApproxCount
            Me.DBG_DETALLE.Row = num
            If Me.DBG_DETALLE.ApproxCount > 1 And num <> 0 Then
                transaccion = transaccion & " OR (COD_TRANSACCION  = '" & Format(DBG_DETALLE.Columns(2), "0###") & "' AND APLICACION_FUENTE = '" & DBG_DETALLE.Columns(1) & "')"
            Else
                transaccion = " AND ( (COD_TRANSACCION  = '" & Format(DBG_DETALLE.Columns(2), "0###") & "' AND APLICACION_FUENTE = '" & DBG_DETALLE.Columns(1) & "')"
            End If
            num = num + 1
        Wend
        '    Else
        '        transaccion = " AND (APLICACION_FUENTE =APLICACION_FUENTE"
        '    End If
        ' If dbg_consulta.Columns(0) = "CONC" then
        PL_Conexion_Oracle

        If ING_BD_CONSULTAR = True Then
            sqlTablas = "  TBL_REGISTRO_TRADUCTOR "
        Else
            sqlTablas = " TBL_REGISTRO_TRADUCTOR@ARCHIST"
        End If

        sentencia = "DELETE TBL_REGISTRO_TRADUCTOR_CONS WHERE COD_USR_SOLICITO = '" & Stg_cod_Usuario & "'"
        cnObj1.Execute (sentencia)

        'MsgBox sentencia

        sentencia = "INSERT INTO TBL_REGISTRO_TRADUCTOR_CONS " & _
                    "(FECHA_SISTEMA, SEC_REGISTRO,COD_USR_SOLICITO, COD_ENTIDAD,APLICACION_FUENTE,COD_TRANSACCION,FECHA_MOVIMIENTO,CENTRO_ORIGEN," & _
                  " CENTRO_DESTINO,TIPO_CUENTA,COD_CUENTA,num_documento,FILLER," & _
                  " NUM_CAMPOS_MONETARIOS,TAM_CAMPOS_MONETARIOS, COD_TIPO_REGISTRO , FLAG_AJUSTE) " & _
                  " SELECT FECHA_SISTEMA, SEC_REGISTRO," & Stg_cod_Usuario & ", COD_ENTIDAD,APLICACION_FUENTE,COD_TRANSACCION,FECHA_MOVIMIENTO,CENTRO_ORIGEN," & _
                  " CENTRO_DESTINO,TIPO_CUENTA,COD_CUENTA,num_documento,FILLER," & _
                  " NUM_CAMPOS_MONETARIOS,TAM_CAMPOS_MONETARIOS, COD_TIPO_REGISTRO, 0 " & _
                    "FROM " & sqlTablas & _
                  " WHERE COD_ENTIDAD = 1" & transaccion & ")" & _
                  " AND FECHA_SISTEMA = TO_DATE('" & Dtg_Fecha_movimiento1 & "','DD/MM/YYYY') " & _
                  " AND  FECHA_MOVIMIENTO = TO_DATE('" & Format(DBG_DETALLE.Columns(3), "DD/MM/YYYY") & "','DD/MM/YYYY') " & _
                  " AND CENTRO_ORIGEN = " & Val(DBG_DETALLE.Columns(4)) & _
                  " AND CENTRO_DESTINO = " & Val(DBG_DETALLE.Columns(5)) & " AND COD_TIPO_REGISTRO = 1 AND FLAG_AJUSTE = 1 "
        'MsgBox sentencia
        cnObj1.Execute (sentencia)

        If ING_BD_CONSULTAR = True Then
            INL_BD_CONSULTAR = 1
        Else
            INL_BD_CONSULTAR = 0
        End If

        cnObj1.Execute "USRLARC.PL_CAMPOS_VALOR_TR('" & Stg_cod_Usuario & "'," & INL_BD_CONSULTAR & ") "

        Set RSOBJ3 = cargarRecordSet("SELECT  COD_ENTIDAD,APLICACION_FUENTE,COD_TRANSACCION,FECHA_MOVIMIENTO,CENTRO_ORIGEN," & _
                                   " CENTRO_DESTINO,TIPO_CUENTA,COD_CUENTA,num_documento,FILLER," & _
                                   " NUM_CAMPOS_MONETARIOS,TAM_CAMPOS_MONETARIOS, COD_TIPO_REGISTRO,  " & _
                                   " COD_CAMPO1,VALOR1,COD_CAMPO2,VALOR2,COD_CAMPO3,VALOR3,COD_CAMPO4,VALOR4,COD_CAMPO5,VALOR5, " & _
                                   " COD_CAMPO6,VALOR6,COD_CAMPO7,VALOR7,COD_CAMPO8,VALOR8,COD_CAMPO9,VALOR9,COD_CAMPO10,VALOR10 " & _
                                   " FROM TBL_REGISTRO_TRADUCTOR_CONS WHERE COD_USR_SOLICITO = '" & Stg_cod_Usuario & "' " & _
                                   " ORDER BY COD_CUENTA,COD_CAMPO1,VALOR1,COD_CAMPO2,VALOR2,COD_CAMPO3,VALOR3,COD_CAMPO4,VALOR4,APLICACION_FUENTE  ")
        If Not RSOBJ3.EOF Then
            Set Me.dtg_detalle4.DataSource = RSOBJ3
            tipo_configuracion = 7
            Configurar_grilla_Dat_consulta Me.dtg_detalle4
        End If
    End If
    Set RSOBJ3 = Nothing
    MousePointer = 0
End Sub











