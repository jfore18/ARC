VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmConsultaResponsabilidades 
   Caption         =   "Asignación de responsabilidades"
   ClientHeight    =   7275
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7275
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Asignación de responsabilidades"
      ForeColor       =   &H00800000&
      Height          =   7215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11850
      Begin TabDlg.SSTab SSTab2 
         Height          =   6255
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   11033
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Tipos de movimiento"
         TabPicture(0)   =   "frmConsultaResponsabilidades.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "arbolMovimiento"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Transacciones y Conciliaciones"
         TabPicture(1)   =   "frmConsultaResponsabilidades.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "mskTransacciones"
         Tab(1).Control(1)=   "arbolTransaccion"
         Tab(1).Control(2)=   "Label1"
         Tab(1).ControlCount=   3
         Begin MSMask.MaskEdBox mskTransacciones 
            Height          =   255
            Left            =   -71760
            TabIndex        =   5
            Top             =   5520
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin MSComctlLib.TreeView arbolTransaccion 
            Height          =   4815
            Left            =   -74760
            TabIndex        =   3
            Top             =   480
            Width           =   10935
            _ExtentX        =   19288
            _ExtentY        =   8493
            _Version        =   393217
            Style           =   7
            Appearance      =   1
         End
         Begin MSComctlLib.TreeView arbolMovimiento 
            Height          =   5055
            Left            =   240
            TabIndex        =   2
            Top             =   600
            Width           =   10935
            _ExtentX        =   19288
            _ExtentY        =   8916
            _Version        =   393217
            Style           =   7
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label1 
            Caption         =   "Transacción (FUENTE TRANSACCION)"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   -74760
            TabIndex        =   4
            Top             =   5520
            Width           =   3015
         End
      End
   End
End
Attribute VB_Name = "frmConsultaResponsabilidades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cargarArbolMovimiento()


'Caracteristicas de arbol
    Me.arbolMovimiento.LabelEdit = tvwManual
    Me.arbolMovimiento.Nodes.Clear

    'Declaración de variables para las consultas

    Dim sentencia As String
    Dim registro As String
    Dim descripcion As String
    Dim Usuario As String
    Dim nodo As Node
    
    ' REQ:266003. Ajustes por administración de usuarios por IDM.
    ' ABOCANE. DICIEMBRE 2016
    
       ' Según el perfil del usuario construimos el arbol de transacciones asignadas
    
    If Stg_Perfil_Usuario_Acceso = 8 Then
    
    'Construccion del arbol con archivos y sus descripciones
    sentencia = " SELECT  NVL(U.COD_USR,0) USUARIO, LPAD(M.COD_TIPO_REGISTRO,2,0)REGISTRO," & _
                " M.DESC_TIPO_REGISTRO DESCRIPCION" & _
                " FROM  TBL_TIPO_REGISTRO M, TBL_USUARIO_MOVIMIENTO U" & _
                " WHERE M.COD_TIPO_REGISTRO = U.COD_TIPO_REGISTRO AND" & _
                " U.COD_USR =" & Stg_cod_Usuario & " ORDER BY 2 "
    
    End If
    
    
    If Stg_Perfil_Usuario_Acceso = 9 Then
    
    'Construccion del arbol con archivos y sus descripciones
    sentencia = " SELECT LPAD(M.COD_TIPO_REGISTRO,2,0)REGISTRO," & _
                " M.DESC_TIPO_REGISTRO DESCRIPCION" & _
                " FROM  TBL_TIPO_REGISTRO M ORDER BY 1"
               
    End If
      
    Set rsobj = cnObj1.Execute(sentencia)

    'Variable que expresa el número de archivos asignables
    Dim I As Integer
    I = 0

    'Se construye el arbol con solo hojas
    Do While Not rsobj.EOF
        registro = rsobj("REGISTRO")
        descripcion = rsobj("DESCRIPCION")

        'Asigna la clave e inserta los nodos ARCHIVO

        Set nodo = Me.arbolMovimiento.Nodes.Add(, , "R" & registro, registro & " - " & descripcion)

        'se desplaza al siguiente registro del resultSet
        rsobj.MoveNext

        I = I + 1
    Loop
End Sub
Private Sub cargarArbolTransacciones()

    Me.MousePointer = 11
    'Limpia el contenido del arbol
    Me.arbolTransaccion.Nodes.Clear

    'Caracteristicas de arbol
    Me.arbolTransaccion.LabelEdit = tvwManual

    'Declaración de variables para las consultas

    Dim sentencia As String
    Dim negocio As String
    Dim descripcionNegocio As String
    Dim Usuario As String
    Dim negocioAnterior As String
    Dim conciliacion As String
    Dim conciliacionAnterior As String
    Dim transaccion As String
    Dim descripcionTransaccion As String
    Dim clave1 As String
    Dim clave2 As String
    Dim nodo As Node

    'Conexion a la base de datos
    PL_Conexion_Oracle
    
    ' REQ:266003. Ajustes por administración de usuarios por IDM.
    ' ABOCANE. DICIEMBRE 2016
    
    ' Según el perfil del usuario construimos el arbol de transacciones asignadas
    
    If Stg_Perfil_Usuario_Acceso = 8 Then
    
       sentencia = " SELECT  NVL(LPAD(T.COD_NEGOCIO,4,'0'),'0000') NEGOCIO, NVL(N.DESC_NEGOCIO,'NINGUNO')" & _
                " DESCRIPCION_NEGOCIO, NVL(T.COD_APLICACION_CONC||T.COD_TRANSACCION_CONC,'TRANSACCIONES LIBRES')" & _
                " CONCILIACION,NVL(T.DESC_TRANSACCION,'TRANSACCION SIN NOMBRE')DESCRIPCION_TRANSACCION," & _
                " T.COD_APLICACION_FUENTE||' '||T.COD_TRANSACCION TRANSACCION ,U.COD_USR USUARIO" & _
                " FROM TBL_TRANSACCION_TRADUCTOR T, TBL_NEGOCIO N, VTA_USUARIO_TRANSACCION U " & _
                " WHERE T.COD_ENTIDAD  = N.COD_ENTIDAD (+) AND" & _
                " T.COD_NEGOCIO  = N.COD_NEGOCIO (+) AND " & _
                " T.COD_APLICACION_FUENTE =  U.COD_APLICACION_FUENTE AND " & _
                " T.COD_TRANSACCION =  U.COD_TRANSACCION AND " & _
                " U.COD_USR  =" & Stg_cod_Usuario & " AND " & _
                " T.COD_APLICACION_FUENTE <> 'CONC'" & _
                " ORDER BY 1,CONCILIACION,TRANSACCION"
    End If
    
    
       If Stg_Perfil_Usuario_Acceso = 9 Then
    
       sentencia = " SELECT  NVL(LPAD(T.COD_NEGOCIO,4,'0'),'0000') NEGOCIO, NVL(N.DESC_NEGOCIO,'NINGUNO')" & _
                " DESCRIPCION_NEGOCIO, NVL(T.COD_APLICACION_CONC||T.COD_TRANSACCION_CONC,'TRANSACCIONES LIBRES')" & _
                " CONCILIACION,NVL(T.DESC_TRANSACCION,'TRANSACCION SIN NOMBRE')DESCRIPCION_TRANSACCION," & _
                " T.COD_APLICACION_FUENTE||' '||T.COD_TRANSACCION TRANSACCION" & _
                " FROM TBL_TRANSACCION_TRADUCTOR T, TBL_NEGOCIO N" & _
                " WHERE T.COD_ENTIDAD  = N.COD_ENTIDAD (+) AND" & _
                " T.COD_NEGOCIO  = N.COD_NEGOCIO (+) AND " & _
                " T.COD_APLICACION_FUENTE <> 'CONC'" & _
                " ORDER BY 1,CONCILIACION,TRANSACCION"
    End If
       
    Set rsobj = cnObj1.Execute(sentencia)

    Set nodo = Me.arbolTransaccion.Nodes.Add(, , "raiz", "TODOS LOS NEGOCIOS")
    nodo.Expanded = True

    'Si no existen transacciones en la tabla TBL_TRANSACCION_TRADUCTOR
    If rsobj.EOF Then
        MsgBox "No existen transacciones disponibles para cargar."
        Exit Sub
    End If

    Do While Not rsobj.EOF
        negocio = rsobj("NEGOCIO")
        descripcionNegocio = rsobj("DESCRIPCION_NEGOCIO")
        conciliacion = rsobj("CONCILIACION")
        transaccion = rsobj("TRANSACCION")
        descripcionTransaccion = rsobj("DESCRIPCION_TRANSACCION")
       ' Usuario = rsobj("USUARIO")


        'Asigna la clave e inserta los nodos NEGOCIO
        clave1 = "N" & negocio
        If (negocioAnterior <> negocio) Then
            Set nodo = Me.arbolTransaccion.Nodes.Add("raiz", tvwChild, clave1, "NEGOCIO: " & descripcionNegocio)
            nodo.Expanded = True
            conciliacionAnterior = ""
        End If

        'Asigna la clave e inserta los nodos CONCILIACION
        clave2 = clave1 & "C" & conciliacion
        If (conciliacionAnterior <> conciliacion) Then
            Set nodo = Me.arbolTransaccion.Nodes.Add(clave1, tvwChild, clave2, "CONCILIACION: " & conciliacion)
            nodo.Expanded = False
        End If


        'Asigna la clave e inserta los nodos TRANSACCION
        Set nodo = Me.arbolTransaccion.Nodes.Add(clave2, tvwChild, "TRANS" & transaccion, transaccion & " " & descripcionTransaccion)


        'Se actualizan las variables que controlan la inserción en el arbol
        negocioAnterior = negocio
        conciliacionAnterior = conciliacion
        'se desplaza al sigueinte registro del resultSet
        rsobj.MoveNext
    Loop

    Me.MousePointer = 0

End Sub
Private Sub Form_Load()
'Carga los datos del arbol de transacciones
    cargarArbolTransacciones
'Carga los datos del arbol de archivos
    cargarArbolMovimiento
End Sub
Private Sub mskTransacciones_GotFocus()
    mskTransacciones.SelStart = 0
    mskTransacciones.SelLength = Len(mskTransacciones)
End Sub

Private Sub mskTransacciones_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Dim I As Integer
        For I = 1 To Me.arbolTransaccion.Nodes.Count
            arbolTransaccion.Nodes(I).Bold = False
        Next I
        For I = 1 To Me.arbolTransaccion.Nodes.Count
            If (Mid(Me.arbolTransaccion.Nodes(I).Text, 1, 9) = Me.mskTransacciones.Text) Or (Mid(Me.arbolTransaccion.Nodes(I).Text, 15, 23) = Mid(Me.mskTransacciones.Text, 1, 4) & Mid(Me.mskTransacciones.Text, 6, 4)) Then
                Me.arbolTransaccion.Nodes(I).EnsureVisible
                Set arbolTransaccion.SelectedItem = Me.arbolTransaccion.Nodes(I)
                arbolTransaccion.Nodes(I).Bold = True
                Exit Sub
            End If
        Next I
        MsgBox "No se encontró la transacción " & Me.mskTransacciones.Text
        Me.mskTransacciones.SetFocus
        For I = 1 To Me.arbolTransaccion.Nodes.Count
            arbolTransaccion.Nodes(I).Bold = False
        Next I
    End If
    If KeyAscii = vbKeyEscape Then
        mskTransacciones.Text = "         "
    End If
End Sub

