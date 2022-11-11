Attribute VB_Name = "sistemaARC"
Option Explicit
Dim Stg_Tabla As String
Dim Stg_Campo_Retorno As String
Dim Stg_Where As String
Dim Stg_Order As String
Dim hora As String

' ==========================================================
' = Get Windows Information                                =
' ==========================================================

Public Const MAX_COMPUTERNAME_LENGTH = 31
Public Const MAX_PATH = 260
Public Const UNLEN = 256

Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

' ==========================================================
' = Shutdown Windows                                       =
' ==========================================================

Public Const EWX_FORCE = 4
Public Const EWX_LOGOFF = 0
Public Const EWX_REBOOT = 2
Public Const EWX_SHUTDOWN = 1

Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long

Sub Pg_Informacion_sistema()

    Dim llAux1 As Long
    Dim llAux2 As Long
    Dim llAux3 As String

    'Nombre del Computador

    llAux3 = String(CInt(MAX_COMPUTERNAME_LENGTH + 1), Chr(0))
    llAux2 = MAX_COMPUTERNAME_LENGTH
    llAux1 = GetComputerName(llAux3, llAux2)
    sNombre_Computador = Left(llAux3, llAux2)

    'Nombre del Usuario Windows NT

    llAux3 = String(UNLEN + 1, Chr(0))
    llAux2 = UNLEN
    llAux1 = GetUserName(llAux3, llAux2)
    sUsuario_Windows = Left(llAux3, llAux2)
    
End Sub



