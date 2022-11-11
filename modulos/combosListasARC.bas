Attribute VB_Name = "combosListasARC"
Option Explicit
Sub PG_Llenar_Combo_Lista(Tabla As ADODB.Recordset, Lista As ListBox, campo As String)
    On Error GoTo error

    If Tabla.RecordCount <> 0 Then
        Tabla.MoveFirst
        Lista.Clear
        While Not Tabla.EOF
            Lista.AddItem Tabla(campo)
            Tabla.MoveNext
        Wend
    Else
        MsgBox "No encontro registros en la tabla"
    End If

    Exit Sub
error:
    If Err.Number = 91 Then
        MsgBox "Existen Problemas de comunicación con la Base de datos , contacte al administrador de la red.  Ext. 3291", vbCritical
    Else
        MsgBox Err.Number & " " & Err.Description
    End If

End Sub
Function PG_Llenar_Combo_Lista2(Tabla As ADODB.Recordset, Combo As ComboBox, campo As String, campo2 As String, num_digitos As Integer)
    On Error GoTo error
    Combo.Clear
    If Tabla.RecordCount <> 0 Then
        Tabla.MoveFirst
        While Not Tabla.EOF
            Combo.AddItem Format$(Tabla(campo), "0" & String(num_digitos - 1, "#")) & "-" & Tabla(campo2)
            Tabla.MoveNext
        Wend
    Else
        MsgBox "No encontró registros en la tabla"
    End If
    Exit Function
error:
    If Err.Number = 91 Then
        MsgBox "Existen Problemas de comunicación con la Base de datos , contacte al administrador de la red.  Ext. 3291", vbCritical
    Else
        MsgBox Err.Number & " " & Err.Description
0
    End If

End Function
Function PG_Llenar_Combo_Lista_Consul(Tabla As ADODB.Recordset, Combo As ListBox, campo As String, campo2 As String, num_digitos As Integer)
    On Error GoTo error

    If Tabla.RecordCount <> 0 Then
        Tabla.MoveFirst
        Combo.Clear
        While Not Tabla.EOF
            Combo.AddItem Format$(Tabla(campo), "0" & String(num_digitos - 1, "#")) & "-" & Tabla(campo2)
            Tabla.MoveNext
        Wend
    Else
        MsgBox "No encontró registros en la tabla"
    End If

    Exit Function
error:
    If Err.Number = 91 Then
        MsgBox "Existen Problemas de comunicación con la Base de datos , contacte al administrador de la red.  Ext. 3291", vbCritical
    Else
        MsgBox Err.Number & " " & Err.Description
    End If

End Function
Function PG_Llenar_ComboBox(Tabla As ADODB.Recordset, Combo As ComboBox, campo As String, num_digitos As Integer)
    On Error GoTo error

    If Tabla.RecordCount <> 0 Then
        Tabla.MoveFirst
        'Combo.Clear
        While Not Tabla.EOF
            Combo.AddItem Format$(Tabla(campo), "0" & String(num_digitos - 1, "#"))
            Tabla.MoveNext

        Wend
    Else
        MsgBox "No encontró registros en la tabla"
    End If

    Exit Function
error:
    If Err.Number = 91 Then
        MsgBox "Existen Problemas de comunicación con la Base de datos , contacte al administrador de la red.  Ext. 3291", vbCritical
    Else
        MsgBox Err.Number & " " & Err.Description
    End If

End Function
Sub PG_Llenar_Combo_Codigo(Tabla As ADODB.Recordset, Combo As ComboBox, campo As String, num_digitos As Integer)
'---------------------------------------------------------------------------
' NOMBRE:   PG_Llenar_combo
'
' DESCRIPCION:
' Recorre una tabla de la base de datos correspondiente a campo <tabla>,
' los campos de cada registro son  añadidos en forma consecutiva a un combo
'
' PARAMETROS:
'   tabla : Objeto tipo Recorset del cual se consultan los datos
'   Lista : Objeto tipo Combo en el cual se adicionan los datos consultados
'
' ---------------------------------------------------------------------------
    On Error GoTo error

    If Tabla.RecordCount > 0 Then
        Tabla.MoveFirst
        'Combo.Clear
        While Not Tabla.EOF
            Combo.AddItem Format$(Tabla(campo), "0" & String(num_digitos - 1, "#"))
            Tabla.MoveNext
        Wend
    Else
        MsgBox "No encontró registros en la tabla"
    End If

    Exit Sub
error:
    If Err.Number = 91 Then
        MsgBox "Existen Problemas de comunicación con la Base de datos , contacte al administrador de la red.  Ext. 3291", vbCritical
    Else
        MsgBox Err.Number & " " & Err.Description
    End If


End Sub
Sub PG_Llenar_Combo_Texto(Tabla As ADODB.Recordset, Combo As ComboBox, campo As String)
'---------------------------------------------------------------------------
' NOMBRE:   PG_Llenar_combo
'
' DESCRIPCION:
' Recorre una tabla de la base de datos correspondiente a campo <tabla>,
' los campos de cada registro son  añadidos en forma consecutiva a un combo
'
' PARAMETROS:
'   tabla : Objeto tipo Recorset del cual se consultan los datos
'   Lista : Objeto tipo Combo en el cual se adicionan los datos consultados
'
' ---------------------------------------------------------------------------
    On Error GoTo error

    If Tabla.RecordCount > 0 Then
        Tabla.MoveFirst
        'Combo.Clear
        While Not Tabla.EOF
            Combo.AddItem Tabla(campo)
            Tabla.MoveNext
        Wend
    Else
        MsgBox "No encontró registros en la tabla"
    End If

    Exit Sub
error:
    If Err.Number = 91 Then
        MsgBox "Existen Problemas de comunicación con la Base de datos , contacte al administrador de la red.  Ext. 3291", vbCritical
    Else
        MsgBox Err.Number & " " & Err.Description
    End If

End Sub

Sub PG_Llenar_Combo(Tabla As ADODB.Recordset, Combo As ComboBox)

    On Error GoTo error

    Dim n As Integer
    Dim I As Integer
    Dim Linea As String

    n = Tabla.Fields.Count

    If Tabla.RecordCount > 0 Then
        Tabla.MoveFirst
        'Combo.Clear
        While Not Tabla.EOF
            Linea = ""
            For I = 0 To n - 1
                Linea = Linea & Tabla.Fields(I).Value & "  "
            Next I
            Combo.AddItem Linea
            Tabla.MoveNext
        Wend
    Else
        MsgBox "No encontró registros en la tabla"
    End If

    Exit Sub
error:
    If Err.Number = 91 Then
        MsgBox "Existen Problemas de comunicación con la Base de datos , contacte al administrador de la red.  Ext. 3291", vbCritical
    Else
        MsgBox Err.Number & " " & Err.Description
    End If

End Sub

