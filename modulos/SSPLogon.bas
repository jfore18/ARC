Attribute VB_Name = "SSPLogon"
' Module Name:  SSPLogon.bas
Option Explicit

Private Const HEAP_ZERO_MEMORY = &H8

Private Const SEC_WINNT_AUTH_IDENTITY_ANSI = &H1

Private Const SECBUFFER_TOKEN = &H2

Private Const SECURITY_NATIVE_DREP = &H10

Private Const SECPKG_CRED_INBOUND = &H1
Private Const SECPKG_CRED_OUTBOUND = &H2

Private Const SEC_I_CONTINUE_NEEDED = &H90312
Private Const SEC_I_COMPLETE_NEEDED = &H90313
Private Const SEC_I_COMPLETE_AND_CONTINUE = &H90314

Private Const VER_PLATFORM_WIN32_NT = &H2

Type SecPkgInfo
    fCapabilities As Long
    wVersion As Integer
    wRPCID As Integer
    cbMaxToken As Long
    Name As Long
    Comment As Long
End Type

Type SecHandle
    dwLower As Long
    dwUpper As Long
End Type

Type AUTH_SEQ
    fInitialized As Boolean
    fHaveCredHandle As Boolean
    fHaveCtxtHandle As Boolean
    hcred As SecHandle
    hctxt As SecHandle
End Type

Type SEC_WINNT_AUTH_IDENTITY
    user As String
    UserLength As Long
    domain As String
    DomainLength As Long
    Password As String
    PasswordLength As Long
    Flags As Long
End Type

Type TimeStamp
    LowPart As Long
    HighPart As Long
End Type

Type SecBuffer
    cbBuffer As Long
    BufferType As Long
    pvBuffer As Long
End Type

Type SecBufferDesc
    ulVersion As Long
    cBuffers As Long
    pBuffers As Long
End Type

Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type


'This Data Structure lets define an active Directory user with all information
Public Type ADUser
    distinguishedName As String
    displayName As String
    mail As String
    givenName As String
    SN As String    'SN= SurName
    telephoneNumber As String
    otherTelephone As String
    cn As String    'cn= Complete Name
    initials As String
    physicalDeliveryOfficeName As String
    wWWHomePage As String
    url As String
    streetAddress As String
    postOfficeBox As String
    l As String    'l= City
    st As String    ' st=State/Province
    postalCode As String
    c As String    'Country abreviated
    co As String    'Country complete
    countryCode As String
    userPrincipalName As String
    sAMAccountName As String
    userAccountControl As String    'Returns a number, see more info in http://support.microsoft.com/kb/305144
    userWorkstations As String
    profilePath As String
    scriptPath As String
    homeDirectory As String
    homeDrive As String
    homePhone As String
    info As String
    title As String
    department As String
    company As String
    manager As String
End Type

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
                               (Destination As Any, Source As Any, ByVal Length As Long)

Private Declare Function NT4QuerySecurityPackageInfo Lib "security" _
                                                     Alias "QuerySecurityPackageInfoA" (ByVal PackageName As String, _
                                                                                        ByRef pPackageInfo As Long) As Long

Private Declare Function QuerySecurityPackageInfo Lib "secur32" _
                                                  Alias "QuerySecurityPackageInfoA" (ByVal PackageName As String, _
                                                                                     ByRef pPackageInfo As Long) As Long

Private Declare Function NT4FreeContextBuffer Lib "security" _
                                              Alias "FreeContextBuffer" (ByVal pvContextBuffer As Long) As Long

Private Declare Function FreeContextBuffer Lib "secur32" _
                                           (ByVal pvContextBuffer As Long) As Long

Private Declare Function NT4InitializeSecurityContext Lib "security" _
                                                      Alias "InitializeSecurityContextA" _
                                                      (ByRef phCredential As SecHandle, ByRef phContext As SecHandle, _
                                                       ByVal pszTargetName As Long, ByVal fContextReq As Long, _
                                                       ByVal Reserved1 As Long, ByVal TargetDataRep As Long, _
                                                       ByRef pInput As SecBufferDesc, ByVal Reserved2 As Long, _
                                                       ByRef phNewContext As SecHandle, ByRef pOutput As SecBufferDesc, _
                                                       ByRef pfContextAttr As Long, ByRef ptsExpiry As TimeStamp) As Long

Private Declare Function InitializeSecurityContext Lib "secur32" _
                                                   Alias "InitializeSecurityContextA" _
                                                   (ByRef phCredential As SecHandle, ByRef phContext As SecHandle, _
                                                    ByVal pszTargetName As Long, ByVal fContextReq As Long, _
                                                    ByVal Reserved1 As Long, ByVal TargetDataRep As Long, _
                                                    ByRef pInput As SecBufferDesc, ByVal Reserved2 As Long, _
                                                    ByRef phNewContext As SecHandle, ByRef pOutput As SecBufferDesc, _
                                                    ByRef pfContextAttr As Long, ByRef ptsExpiry As TimeStamp) As Long

Private Declare Function NT4InitializeSecurityContext2 Lib "security" _
                                                       Alias "InitializeSecurityContextA" _
                                                       (ByRef phCredential As SecHandle, ByVal phContext As Long, _
                                                        ByVal pszTargetName As Long, ByVal fContextReq As Long, _
                                                        ByVal Reserved1 As Long, ByVal TargetDataRep As Long, _
                                                        ByVal pInput As Long, ByVal Reserved2 As Long, _
                                                        ByRef phNewContext As SecHandle, ByRef pOutput As SecBufferDesc, _
                                                        ByRef pfContextAttr As Long, ByRef ptsExpiry As TimeStamp) As Long

Private Declare Function InitializeSecurityContext2 Lib "secur32" _
                                                    Alias "InitializeSecurityContextA" _
                                                    (ByRef phCredential As SecHandle, ByVal phContext As Long, _
                                                     ByVal pszTargetName As Long, ByVal fContextReq As Long, _
                                                     ByVal Reserved1 As Long, ByVal TargetDataRep As Long, _
                                                     ByVal pInput As Long, ByVal Reserved2 As Long, _
                                                     ByRef phNewContext As SecHandle, ByRef pOutput As SecBufferDesc, _
                                                     ByRef pfContextAttr As Long, ByRef ptsExpiry As TimeStamp) As Long

Private Declare Function NT4AcquireCredentialsHandle Lib "security" _
                                                     Alias "AcquireCredentialsHandleA" (ByVal pszPrincipal As Long, _
                                                                                        ByVal pszPackage As String, ByVal fCredentialUse As Long, _
                                                                                        ByVal pvLogonId As Long, _
                                                                                        ByRef pAuthData As SEC_WINNT_AUTH_IDENTITY, _
                                                                                        ByVal pGetKeyFn As Long, ByVal pvGetKeyArgument As Long, _
                                                                                        ByRef phCredential As SecHandle, ByRef ptsExpiry As TimeStamp) _
                                                                                        As Long

Private Declare Function AcquireCredentialsHandle Lib "secur32" _
                                                  Alias "AcquireCredentialsHandleA" (ByVal pszPrincipal As Long, _
                                                                                     ByVal pszPackage As String, ByVal fCredentialUse As Long, _
                                                                                     ByVal pvLogonId As Long, _
                                                                                     ByRef pAuthData As SEC_WINNT_AUTH_IDENTITY, _
                                                                                     ByVal pGetKeyFn As Long, ByVal pvGetKeyArgument As Long, _
                                                                                     ByRef phCredential As SecHandle, ByRef ptsExpiry As TimeStamp) _
                                                                                     As Long

Private Declare Function NT4AcquireCredentialsHandle2 Lib "security" _
                                                      Alias "AcquireCredentialsHandleA" (ByVal pszPrincipal As Long, _
                                                                                         ByVal pszPackage As String, ByVal fCredentialUse As Long, _
                                                                                         ByVal pvLogonId As Long, ByVal pAuthData As Long, _
                                                                                         ByVal pGetKeyFn As Long, ByVal pvGetKeyArgument As Long, _
                                                                                         ByRef phCredential As SecHandle, ByRef ptsExpiry As TimeStamp) _
                                                                                         As Long

Private Declare Function AcquireCredentialsHandle2 Lib "secur32" _
                                                   Alias "AcquireCredentialsHandleA" (ByVal pszPrincipal As Long, _
                                                                                      ByVal pszPackage As String, ByVal fCredentialUse As Long, _
                                                                                      ByVal pvLogonId As Long, ByVal pAuthData As Long, _
                                                                                      ByVal pGetKeyFn As Long, ByVal pvGetKeyArgument As Long, _
                                                                                      ByRef phCredential As SecHandle, ByRef ptsExpiry As TimeStamp) _
                                                                                      As Long

Private Declare Function NT4AcceptSecurityContext Lib "security" _
                                                  Alias "AcceptSecurityContext" (ByRef phCredential As SecHandle, _
                                                                                 ByRef phContext As SecHandle, ByRef pInput As SecBufferDesc, _
                                                                                 ByVal fContextReq As Long, ByVal TargetDataRep As Long, _
                                                                                 ByRef phNewContext As SecHandle, ByRef pOutput As SecBufferDesc, _
                                                                                 ByRef pfContextAttr As Long, ByRef ptsExpiry As TimeStamp) As Long

Private Declare Function AcceptSecurityContext Lib "secur32" _
                                               (ByRef phCredential As SecHandle, _
                                                ByRef phContext As SecHandle, ByRef pInput As SecBufferDesc, _
                                                ByVal fContextReq As Long, ByVal TargetDataRep As Long, _
                                                ByRef phNewContext As SecHandle, ByRef pOutput As SecBufferDesc, _
                                                ByRef pfContextAttr As Long, ByRef ptsExpiry As TimeStamp) As Long

Private Declare Function NT4AcceptSecurityContext2 Lib "security" _
                                                   Alias "AcceptSecurityContext" (ByRef phCredential As SecHandle, _
                                                                                  ByVal phContext As Long, ByRef pInput As SecBufferDesc, _
                                                                                  ByVal fContextReq As Long, ByVal TargetDataRep As Long, _
                                                                                  ByRef phNewContext As SecHandle, ByRef pOutput As SecBufferDesc, _
                                                                                  ByRef pfContextAttr As Long, ByRef ptsExpiry As TimeStamp) As Long

Private Declare Function AcceptSecurityContext2 Lib "secur32" _
                                                Alias "AcceptSecurityContext" (ByRef phCredential As SecHandle, _
                                                                               ByVal phContext As Long, ByRef pInput As SecBufferDesc, _
                                                                               ByVal fContextReq As Long, ByVal TargetDataRep As Long, _
                                                                               ByRef phNewContext As SecHandle, ByRef pOutput As SecBufferDesc, _
                                                                               ByRef pfContextAttr As Long, ByRef ptsExpiry As TimeStamp) As Long

Private Declare Function NT4CompleteAuthToken Lib "security" _
                                              Alias "CompleteAuthToken" (ByRef phContext As SecHandle, _
                                                                         ByRef pToken As SecBufferDesc) As Long

Private Declare Function CompleteAuthToken Lib "secur32" _
                                           (ByRef phContext As SecHandle, _
                                            ByRef pToken As SecBufferDesc) As Long

Private Declare Function NT4DeleteSecurityContext Lib "security" _
                                                  Alias "DeleteSecurityContext" (ByRef phContext As SecHandle) _
                                                  As Long

Private Declare Function DeleteSecurityContext Lib "secur32" _
                                               (ByRef phContext As SecHandle) _
                                               As Long

Private Declare Function NT4FreeCredentialsHandle Lib "security" _
                                                  Alias "FreeCredentialsHandle" (ByRef phContext As SecHandle) _
                                                  As Long

Private Declare Function FreeCredentialsHandle Lib "secur32" _
                                               (ByRef phContext As SecHandle) _
                                               As Long

Private Declare Function GetProcessHeap Lib "kernel32" () As Long

Private Declare Function HeapAlloc Lib "kernel32" _
                                   (ByVal hHeap As Long, ByVal dwFlags As Long, _
                                    ByVal dwBytes As Long) As Long

Private Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, _
                                                  ByVal dwFlags As Long, ByVal lpMem As Long) As Long

Private Declare Function GetVersionExA Lib "kernel32" _
                                       (lpVersionInformation As OSVERSIONINFO) As Integer

Dim g_NT4 As Boolean

Private Function GenClientContext(ByRef AuthSeq As AUTH_SEQ, _
                                  ByRef AuthIdentity As SEC_WINNT_AUTH_IDENTITY, _
                                  ByVal pIn As Long, ByVal cbIn As Long, _
                                  ByVal pOut As Long, ByRef cbOut As Long, _
                                  ByRef fDone As Boolean) As Boolean

    Dim ss As Long
    Dim tsExpiry As TimeStamp
    Dim sbdOut As SecBufferDesc
    Dim sbOut As SecBuffer
    Dim sbdIn As SecBufferDesc
    Dim sbIn As SecBuffer
    Dim fContextAttr As Long

    GenClientContext = False

    If Not AuthSeq.fInitialized Then

        If g_NT4 Then
            ss = NT4AcquireCredentialsHandle(0&, "NTLM", _
                                             SECPKG_CRED_OUTBOUND, 0&, AuthIdentity, 0&, 0&, _
                                             AuthSeq.hcred, tsExpiry)
        Else
            ss = AcquireCredentialsHandle(0&, "NTLM", _
                                          SECPKG_CRED_OUTBOUND, 0&, AuthIdentity, 0&, 0&, _
                                          AuthSeq.hcred, tsExpiry)
        End If

        If ss < 0 Then
            Exit Function
        End If

        AuthSeq.fHaveCredHandle = True

    End If

    ' Prepare output buffer
    sbdOut.ulVersion = 0
    sbdOut.cBuffers = 1
    sbdOut.pBuffers = HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, _
                                Len(sbOut))

    sbOut.cbBuffer = cbOut
    sbOut.BufferType = SECBUFFER_TOKEN
    sbOut.pvBuffer = pOut

    CopyMemory ByVal sbdOut.pBuffers, sbOut, Len(sbOut)

    ' Prepare input buffer
    If AuthSeq.fInitialized Then

        sbdIn.ulVersion = 0
        sbdIn.cBuffers = 1
        sbdIn.pBuffers = HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, _
                                   Len(sbIn))

        sbIn.cbBuffer = cbIn
        sbIn.BufferType = SECBUFFER_TOKEN
        sbIn.pvBuffer = pIn

        CopyMemory ByVal sbdIn.pBuffers, sbIn, Len(sbIn)

    End If

    If AuthSeq.fInitialized Then

        If g_NT4 Then
            ss = NT4InitializeSecurityContext(AuthSeq.hcred, _
                                              AuthSeq.hctxt, 0&, 0, 0, SECURITY_NATIVE_DREP, sbdIn, _
                                              0, AuthSeq.hctxt, sbdOut, fContextAttr, tsExpiry)
        Else
            ss = InitializeSecurityContext(AuthSeq.hcred, _
                                           AuthSeq.hctxt, 0&, 0, 0, SECURITY_NATIVE_DREP, sbdIn, _
                                           0, AuthSeq.hctxt, sbdOut, fContextAttr, tsExpiry)
        End If

    Else

        If g_NT4 Then
            ss = NT4InitializeSecurityContext2(AuthSeq.hcred, 0&, 0&, _
                                               0, 0, SECURITY_NATIVE_DREP, 0&, 0, AuthSeq.hctxt, _
                                               sbdOut, fContextAttr, tsExpiry)
        Else
            ss = InitializeSecurityContext2(AuthSeq.hcred, 0&, 0&, _
                                            0, 0, SECURITY_NATIVE_DREP, 0&, 0, AuthSeq.hctxt, _
                                            sbdOut, fContextAttr, tsExpiry)
        End If

    End If

    If ss < 0 Then
        GoTo FreeResourcesAndExit
    End If

    AuthSeq.fHaveCtxtHandle = True

    ' If necessary, complete token
    If ss = SEC_I_COMPLETE_NEEDED _
       Or ss = SEC_I_COMPLETE_AND_CONTINUE Then

        If g_NT4 Then
            ss = NT4CompleteAuthToken(AuthSeq.hctxt, sbdOut)
        Else
            ss = CompleteAuthToken(AuthSeq.hctxt, sbdOut)
        End If

        If ss < 0 Then
            GoTo FreeResourcesAndExit
        End If

    End If

    CopyMemory sbOut, ByVal sbdOut.pBuffers, Len(sbOut)
    cbOut = sbOut.cbBuffer

    If Not AuthSeq.fInitialized Then
        AuthSeq.fInitialized = True
    End If

    fDone = Not (ss = SEC_I_CONTINUE_NEEDED _
                 Or ss = SEC_I_COMPLETE_AND_CONTINUE)

    GenClientContext = True

FreeResourcesAndExit:

    If sbdOut.pBuffers <> 0 Then
        HeapFree GetProcessHeap(), 0, sbdOut.pBuffers
    End If

    If sbdIn.pBuffers <> 0 Then
        HeapFree GetProcessHeap(), 0, sbdIn.pBuffers
    End If

End Function


Private Function GenServerContext(ByRef AuthSeq As AUTH_SEQ, _
                                  ByVal pIn As Long, ByVal cbIn As Long, _
                                  ByVal pOut As Long, ByRef cbOut As Long, _
                                  ByRef fDone As Boolean) As Boolean

    Dim ss As Long
    Dim tsExpiry As TimeStamp
    Dim sbdOut As SecBufferDesc
    Dim sbOut As SecBuffer
    Dim sbdIn As SecBufferDesc
    Dim sbIn As SecBuffer
    Dim fContextAttr As Long

    GenServerContext = False

    If Not AuthSeq.fInitialized Then

        If g_NT4 Then
            ss = NT4AcquireCredentialsHandle2(0&, "NTLM", _
                                              SECPKG_CRED_INBOUND, 0&, 0&, 0&, 0&, AuthSeq.hcred, _
                                              tsExpiry)
        Else
            ss = AcquireCredentialsHandle2(0&, "NTLM", _
                                           SECPKG_CRED_INBOUND, 0&, 0&, 0&, 0&, AuthSeq.hcred, _
                                           tsExpiry)
        End If

        If ss < 0 Then
            Exit Function
        End If

        AuthSeq.fHaveCredHandle = True

    End If

    ' Prepare output buffer
    sbdOut.ulVersion = 0
    sbdOut.cBuffers = 1
    sbdOut.pBuffers = HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, _
                                Len(sbOut))

    sbOut.cbBuffer = cbOut
    sbOut.BufferType = SECBUFFER_TOKEN
    sbOut.pvBuffer = pOut

    CopyMemory ByVal sbdOut.pBuffers, sbOut, Len(sbOut)

    ' Prepare input buffer
    sbdIn.ulVersion = 0
    sbdIn.cBuffers = 1
    sbdIn.pBuffers = HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, _
                               Len(sbIn))

    sbIn.cbBuffer = cbIn
    sbIn.BufferType = SECBUFFER_TOKEN
    sbIn.pvBuffer = pIn

    CopyMemory ByVal sbdIn.pBuffers, sbIn, Len(sbIn)

    If AuthSeq.fInitialized Then

        If g_NT4 Then
            ss = NT4AcceptSecurityContext(AuthSeq.hcred, AuthSeq.hctxt, _
                                          sbdIn, 0, SECURITY_NATIVE_DREP, AuthSeq.hctxt, sbdOut, _
                                          fContextAttr, tsExpiry)
        Else
            ss = AcceptSecurityContext(AuthSeq.hcred, AuthSeq.hctxt, _
                                       sbdIn, 0, SECURITY_NATIVE_DREP, AuthSeq.hctxt, sbdOut, _
                                       fContextAttr, tsExpiry)
        End If

    Else

        If g_NT4 Then
            ss = NT4AcceptSecurityContext2(AuthSeq.hcred, 0&, sbdIn, 0, _
                                           SECURITY_NATIVE_DREP, AuthSeq.hctxt, sbdOut, _
                                           fContextAttr, tsExpiry)
        Else
            ss = AcceptSecurityContext2(AuthSeq.hcred, 0&, sbdIn, 0, _
                                        SECURITY_NATIVE_DREP, AuthSeq.hctxt, sbdOut, _
                                        fContextAttr, tsExpiry)
        End If

    End If

    If ss < 0 Then
        GoTo FreeResourcesAndExit
    End If

    AuthSeq.fHaveCtxtHandle = True

    ' If necessary, complete token
    If ss = SEC_I_COMPLETE_NEEDED _
       Or ss = SEC_I_COMPLETE_AND_CONTINUE Then

        If g_NT4 Then
            ss = NT4CompleteAuthToken(AuthSeq.hctxt, sbdOut)
        Else
            ss = CompleteAuthToken(AuthSeq.hctxt, sbdOut)
        End If

        If ss < 0 Then
            GoTo FreeResourcesAndExit
        End If

    End If

    CopyMemory sbOut, ByVal sbdOut.pBuffers, Len(sbOut)
    cbOut = sbOut.cbBuffer

    If Not AuthSeq.fInitialized Then
        AuthSeq.fInitialized = True
    End If

    fDone = Not (ss = SEC_I_CONTINUE_NEEDED _
                 Or ss = SEC_I_COMPLETE_AND_CONTINUE)

    GenServerContext = True

FreeResourcesAndExit:

    If sbdOut.pBuffers <> 0 Then
        HeapFree GetProcessHeap(), 0, sbdOut.pBuffers
    End If

    If sbdIn.pBuffers <> 0 Then
        HeapFree GetProcessHeap(), 0, sbdIn.pBuffers
    End If

End Function

'Author: David Santana
'Extracted from http://support.microsoft.com/kb/279815/es
'This function makes the user authentication given a domain and password strings
'Input: User: String representing the username in the Active Directory
'Domain: String representing the domain in wich the user 'lives'
'Password: String representing the user password without any encryption
Public Function authenticateUser(user As String, domain As String, _
                                 Password As String) As Boolean

    Dim pSPI As Long
    Dim SPI As SecPkgInfo
    Dim cbMaxToken As Long

    Dim pClientBuf As Long
    Dim pServerBuf As Long

    Dim ai As SEC_WINNT_AUTH_IDENTITY

    Dim asClient As AUTH_SEQ
    Dim asServer As AUTH_SEQ
    Dim cbIn As Long
    Dim cbOut As Long
    Dim fDone As Boolean

    Dim osinfo As OSVERSIONINFO

    authenticateUser = False

    If user = "" Or Password = "" Then
        GoTo FreeResourcesAndExit
    End If

    ' Determine if system is Windows NT (version 4.0 or earlier)
    osinfo.dwOSVersionInfoSize = Len(osinfo)
    osinfo.szCSDVersion = Space$(128)
    GetVersionExA osinfo
    g_NT4 = (osinfo.dwPlatformId = VER_PLATFORM_WIN32_NT And _
             osinfo.dwMajorVersion <= 4)

    ' Get max token size
    If g_NT4 Then
        NT4QuerySecurityPackageInfo "NTLM", pSPI
    Else
        QuerySecurityPackageInfo "NTLM", pSPI
    End If

    CopyMemory SPI, ByVal pSPI, Len(SPI)
    cbMaxToken = SPI.cbMaxToken

    If g_NT4 Then
        NT4FreeContextBuffer pSPI
    Else
        FreeContextBuffer pSPI
    End If

    ' Allocate buffers for client and server messages
    pClientBuf = HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, _
                           cbMaxToken)
    If pClientBuf = 0 Then
        GoTo FreeResourcesAndExit
    End If

    pServerBuf = HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, _
                           cbMaxToken)
    If pServerBuf = 0 Then
        GoTo FreeResourcesAndExit
    End If

    ' Initialize auth identity structure
    ai.domain = domain
    ai.DomainLength = Len(domain)
    ai.user = user
    ai.UserLength = Len(user)
    ai.Password = Password
    ai.PasswordLength = Len(Password)
    ai.Flags = SEC_WINNT_AUTH_IDENTITY_ANSI

    ' Prepare client message (negotiate) .
    cbOut = cbMaxToken
    If Not GenClientContext(asClient, ai, 0, 0, pClientBuf, cbOut, _
                            fDone) Then
        GoTo FreeResourcesAndExit
    End If

    ' Prepare server message (challenge) .
    cbIn = cbOut
    cbOut = cbMaxToken
    If Not GenServerContext(asServer, pClientBuf, cbIn, pServerBuf, _
                            cbOut, fDone) Then
        ' Most likely failure: AcceptServerContext fails with
        ' SEC_E_LOGON_DENIED in the case of bad szUser or szPassword.
        ' Unexpected Result: Logon will succeed if you pass in a bad
        ' szUser and the guest account is enabled in the specified domain.
        GoTo FreeResourcesAndExit
    End If

    ' Prepare client message (authenticate) .
    cbIn = cbOut
    cbOut = cbMaxToken
    If Not GenClientContext(asClient, ai, pServerBuf, cbIn, pClientBuf, _
                            cbOut, fDone) Then
        GoTo FreeResourcesAndExit
    End If

    ' Prepare server message (authentication) .
    cbIn = cbOut
    cbOut = cbMaxToken
    If Not GenServerContext(asServer, pClientBuf, cbIn, pServerBuf, _
                            cbOut, fDone) Then
        GoTo FreeResourcesAndExit
    End If

    authenticateUser = True

FreeResourcesAndExit:

    ' Clean up resources
    If asClient.fHaveCtxtHandle Then
        If g_NT4 Then
            NT4DeleteSecurityContext asClient.hctxt
        Else
            DeleteSecurityContext asClient.hctxt
        End If
    End If

    If asClient.fHaveCredHandle Then
        If g_NT4 Then
            NT4FreeCredentialsHandle asClient.hcred
        Else
            FreeCredentialsHandle asClient.hcred
        End If
    End If

    If asServer.fHaveCtxtHandle Then
        If g_NT4 Then
            NT4DeleteSecurityContext asServer.hctxt
        Else
            DeleteSecurityContext asServer.hctxt
        End If
    End If

    If asServer.fHaveCredHandle Then
        If g_NT4 Then
            NT4FreeCredentialsHandle asServer.hcred
        Else
            FreeCredentialsHandle asServer.hcred
        End If
    End If

    If pClientBuf <> 0 Then
        HeapFree GetProcessHeap(), 0, pClientBuf
    End If

    If pServerBuf <> 0 Then
        HeapFree GetProcessHeap(), 0, pServerBuf
    End If

End Function

'Author: David Santana (dsantan)
'This method fills the parameter aComboBox object with an available domain list
'Input: aComboBox: A ComboBox objetc reference in wich the domain list will be loaded
'       defaultDomain: An optional String representing the name of the default domain wich will be selected by default
'                       If not found leave the first item by default
'Output: No output.
'Mayo de 2015.Requerimiento 169124.Modificado por ABOCANE
'Se ajusta la funcion para obtener la estructura de directorio activo segun el dominio con el cual se ingresa al equipo
'donde esta instalada la aplicacion

Public Function getDomainListInComboBox(aComboBox As Object, Optional defaultDomain As String)
    
    On Error GoTo default
    Dim objConnection
    Dim objCommand
    Dim objRecordSet
    Dim index As Integer

    ' Requerimiento 169124. Definicion de variables
    Dim objRootDSE, strDNSDomain, strBase, strFilter, strAttributes, strQuery, sDepth
    
    ' Obtencion de la conexion
    Set objConnection = CreateObject("ADODB.Connection")
    objConnection.Open "Provider=ADsDSOObject;"
    Set objCommand = CreateObject("ADODB.Command")
    objCommand.ActiveConnection = objConnection
    
    ' Requerimiento 169124
    ' Obtenemos la información del directorio activo, de acuerdo al dominio con el cual se ingresó al equipo donde esta la aplicacion
    Set objRootDSE = GetObject("LDAP://rootDSE")
    strDNSDomain = objRootDSE.Get("configurationNamingContext")
    ' Requerimiento 169124
    ' Construimos los parametros de consulta LDAP,segun sintaxis consulta LDAP con ADO.
    strBase = "<LDAP://" & strDNSDomain & ">"
    strFilter = "(NETBIOSName=*)"
    strAttributes = "name"
    sDepth = "subTree"
    'Construimos la consulta LDAP
    strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";" & sDepth

    'Ejecucion consulta LDAP
    'Consulta anterior:objCommand.CommandText = "<LDAP://CN=Configuration,DC=organizacion,DC=net>;(NETBIOSName=*);objectClass,distinguishedName,name;subtree"
    objCommand.CommandText = strQuery
    Set objRecordSet = objCommand.Execute

    index = 0
    
    While Not objRecordSet.EOF

        aComboBox.AddItem (objRecordSet.Fields("Name"))
        'Select the item named like the defaultDomain by default
        If objRecordSet.Fields("Name") = defaultDomain Then
            index = aComboBox.NewIndex
        End If
        objRecordSet.MoveNext
    Wend
    objConnection.Close
    'Asign the found index for the defaultDomain declared constant. This select by default the item called like the defaultDomain constant in the ComboBox
   aComboBox.ListIndex = index
   Exit Function

' Requerimiento 169124
' Si falla la obtención de dominios, asignamos el default
default:
aComboBox.Clear
aComboBox.AddItem (defaultDomain)

End Function

'Author: David Santana (dsantan)
'This method returns a String array containing the domain list
'Input:No input
'Output: domainArray: A String array in wich the domains will be returned
Public Function getDomainListAsStringArray() As String()

    Dim objConnection
    Dim objCommand
    Dim objRecordSet

    Set objConnection = CreateObject("ADODB.Connection")

    objConnection.Open "Provider=ADsDSOObject;"

    Set objCommand = CreateObject("ADODB.Command")

    objCommand.ActiveConnection = objConnection

    objCommand.CommandText = "<LDAP://CN=Configuration,DC=organizacion,DC=net>;(NETBIOSName=*);objectClass,distinguishedName,name;subtree"

    Set objRecordSet = objCommand.Execute

    Dim domainArray() As String

    ReDim domainArray(1 To objRecordSet.RecordCount) As String

    Dim counter As Integer

    counter = 1

    While Not objRecordSet.EOF

        domainArray(counter) = objRecordSet.Fields("Name")

        objRecordSet.MoveNext

        counter = counter + 1

    Wend

    objConnection.Close
    'Returns the domain list
    getDomainListAsStringArray = domainArray

End Function



'Author: David Santana (dsantan)
'Extracted from http://www.tek-tips.com/viewthread.cfm?qid=766263&page=1
'and http://www.rlmueller.net/UserAttributes.htm
'This function gets the user information given a username
'it uses the RootDSE in wich the computer lives for determine the domain
'Inputs: ntUsername: The user who we want to know its info
'Outputs:SSPLogon.ADUser: A structure represeting the Active Directory User and its info

Function getBasicUserInfoAsADUser(ntUsername As String, domain As String) As SSPLogon.ADUser

    Dim objRootDSE, strDNSDomain, objCommand, objConnection, strQuery
    Dim objRecordSet, strName, strDN
    Dim strBase, strFilter, strAttributes

    Dim user As ADUser

    ' Determine DNS domain name from RootDSE object.
    
    strDNSDomain = getBBTADomainStr(domain)

    ' Use ADO to search Active Directory.
    Set objCommand = CreateObject("ADODB.Command")
    Set objConnection = CreateObject("ADODB.Connection")
    objConnection.Provider = "ADsDSOObject"
    'objConnection.Properties("User ID") = "DOMAIN\account"
    'objConnection.Properties("Password") = "xxxxxx"
    objConnection.Open "Active Directory Provider"
    objCommand.ActiveConnection = objConnection

    ' Search for all user objects. Sort recordset by DisplayName.
    strBase = "<LDAP://" & strDNSDomain & ">"
    strFilter = "(&(objectCategory=person)(objectClass=user)(sAMAccountName=" & ntUsername & "*))"
    strAttributes = "distinguishedName,displayName,mail,givenName,SN,telephoneNumber,otherTelephone" & _
                    ",cn,initials,physicalDeliveryOfficeName,wWWHomePage,url,streetAddress,postOfficeBox,l,st,postalCode," & _
                    "c,co,countryCode,userPrincipalName,sAMAccountName,userAccountControl,userWorkstations,profilePath,scriptPath," & _
                    "homeDirectory,homeDrive,homePhone,info,title,department,company,manager,description"
    'strAttributes = "*"
    strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"

    objCommand.CommandText = strQuery
    objCommand.Properties("Page Size") = 100
    objCommand.Properties("Timeout") = 30
    objCommand.Properties("Cache Results") = False
    objCommand.Properties("Sort On") = "displayName"
    Set objRecordSet = objCommand.Execute

    If objRecordSet.EOF Then
        'User not found
        user.displayName = "User not found"
        GoTo CleanUp
        Exit Function
    End If


    ' Loop through results
    Do Until objRecordSet.EOF
        'We add an empty string for avoid Null values in the assignation statement (& "")
        user.displayName = objRecordSet.Fields("displayName") & ""
        user.mail = objRecordSet.Fields("mail") & ""
        user.distinguishedName = objRecordSet.Fields("distinguishedName") & ""
        user.givenName = objRecordSet.Fields("givenName") & ""
        user.SN = objRecordSet.Fields("SN") & ""
        user.telephoneNumber = objRecordSet.Fields("telephoneNumber") & ""
        user.otherTelephone = objRecordSet.Fields("otherTelephone") & ""
        user.cn = objRecordSet.Fields("cn") & ""
        user.initials = objRecordSet.Fields("initials") & ""
        user.physicalDeliveryOfficeName = objRecordSet.Fields("physicalDeliveryOfficeName") & ""
        user.wWWHomePage = objRecordSet.Fields("wWWHomePage") & ""
        user.url = objRecordSet.Fields("url") & ""
        user.streetAddress = objRecordSet.Fields("streetAddress") & ""
        user.postOfficeBox = objRecordSet.Fields("postOfficeBox") & ""
        user.l = objRecordSet.Fields("l") & ""
        user.st = objRecordSet.Fields("st") & ""
        user.postalCode = objRecordSet.Fields("postalCode") & ""
        user.c = objRecordSet.Fields("c") & ""
        user.co = objRecordSet.Fields("co") & ""
        user.countryCode = objRecordSet.Fields("countryCode") & ""
        user.userPrincipalName = objRecordSet.Fields("userPrincipalName") & ""
        user.sAMAccountName = objRecordSet.Fields("sAMAccountName") & ""
        user.userAccountControl = objRecordSet.Fields("userAccountControl") & ""
        user.userWorkstations = objRecordSet.Fields("userWorkstations") & ""
        user.profilePath = objRecordSet.Fields("profilePath") & ""
        user.scriptPath = objRecordSet.Fields("scriptPath") & ""
        user.homeDirectory = objRecordSet.Fields("homeDirectory") & ""
        user.homeDrive = objRecordSet.Fields("homeDrive") & ""
        user.homePhone = objRecordSet.Fields("homePhone") & ""
        user.info = objRecordSet.Fields("info") & ""
        user.title = objRecordSet.Fields("title") & ""
        user.department = objRecordSet.Fields("department") & ""
        user.company = objRecordSet.Fields("company") & ""
        user.manager = objRecordSet.Fields("manager") & ""

        objRecordSet.MoveNext
    Loop



CleanUp:
    getBasicUserInfoAsADUser = user
    objConnection.Close
    Set objRootDSE = Nothing
    Set objCommand = Nothing
    Set objConnection = Nothing
    Set objRecordSet = Nothing


End Function

'Author: David Santana (dsantan)
'Extracted from http://www.tek-tips.com/viewthread.cfm?qid=766263&page=1
'and http://www.rlmueller.net/UserAttributes.htm
'This function gets the user information given a username
'it uses the RootDSE in wich the computer lives for determine the domain
'Inputs: ntUsername: The user who we want to know its info
'Outputs:String represeting the Active Directory User an its info

Function getBasicUserInfoAsString(ntUsername As String, domain As String) As String

    Dim objRootDSE, strDNSDomain, objCommand, objConnection, strQuery
    Dim objRecordSet, strName, strDN
    Dim strBase, strFilter, strAttributes

    ' Determine DNS domain name from RootDSE object.
    
    strDNSDomain = getBBTADomainStr(domain)

    ' Use ADO to search Active Directory.
    Set objCommand = CreateObject("ADODB.Command")
    Set objConnection = CreateObject("ADODB.Connection")
    objConnection.Provider = "ADsDSOObject"
    'objConnection.Properties("User ID") = "DOMAIN\account"
    'objConnection.Properties("Password") = "xxxxxx"
    objConnection.Open "Active Directory Provider"
    objCommand.ActiveConnection = objConnection

    ' Search for all user objects. Sort recordset by DisplayName.
    strBase = "<LDAP://" & strDNSDomain & ">"
    strFilter = "(&(objectCategory=person)(objectClass=user)(sAMAccountName=" & ntUsername & "*))"
    strAttributes = "distinguishedName,displayName,mail,givenName,SN,telephoneNumber,otherTelephone" & _
                    ",cn,initials,physicalDeliveryOfficeName,wWWHomePage,url,streetAddress,postOfficeBox,l,st,postalCode," & _
                    "c,co,countryCode,userPrincipalName,sAMAccountName,userAccountControl,userWorkstations,profilePath,scriptPath," & _
                    "homeDirectory,homeDrive,homePhone,info,title,department,company,manager,description"
    'strAttributes = "*"
    strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"

    objCommand.CommandText = strQuery
    objCommand.Properties("Page Size") = 100
    objCommand.Properties("Timeout") = 30
    objCommand.Properties("Cache Results") = False
    objCommand.Properties("Sort On") = "displayName"
    Set objRecordSet = objCommand.Execute

    If objRecordSet.EOF Then
        getBasicUserInfoAsString = "User not found"
        GoTo CleanUp
        Exit Function
    End If


    ' Loop through results
    Do Until objRecordSet.EOF
        getBasicUserInfoAsString = "displayName: " & objRecordSet.Fields("displayName") & vbNewLine & _
                                   "mail: " & objRecordSet.Fields("mail") & vbNewLine & _
                                   "distinguishedName: " & objRecordSet.Fields("distinguishedName") & vbNewLine & _
                                   "givenName: " & objRecordSet.Fields("givenName") & vbNewLine & _
                                   "SN: " & objRecordSet.Fields("SN") & vbNewLine & _
                                   "telephoneNumber: " & objRecordSet.Fields("telephoneNumber") & vbNewLine & _
                                   "otherTelephone: " & objRecordSet.Fields("otherTelephone") & vbNewLine & _
                                   "cn: " & objRecordSet.Fields("cn") & vbNewLine & _
                                   "initials: " & objRecordSet.Fields("initials") & vbNewLine & _
                                   "physicalDeliveryOfficeName: " & objRecordSet.Fields("physicalDeliveryOfficeName") & vbNewLine & _
                                   "wWWHomePage: " & objRecordSet.Fields("wWWHomePage") & vbNewLine & _
                                   "url: " & objRecordSet.Fields("url") & vbNewLine & _
                                   "streetAddress: " & objRecordSet.Fields("streetAddress") & vbNewLine & _
                                   "postOfficeBox: " & objRecordSet.Fields("postOfficeBox") & vbNewLine & _
                                   "l: " & objRecordSet.Fields("l") & vbNewLine & _
                                   "st: " & objRecordSet.Fields("st") & vbNewLine & _
                                   "postalCode: " & objRecordSet.Fields("postalCode") & vbNewLine & _
                                   "c: " & objRecordSet.Fields("c") & vbNewLine & _
                                   "co: " & objRecordSet.Fields("co") & vbNewLine & _
                                   "countryCode: " & objRecordSet.Fields("countryCode") & vbNewLine & _
                                   "userPrincipalName: " & objRecordSet.Fields("userPrincipalName") & vbNewLine & _
                                   "sAMAccountName: " & objRecordSet.Fields("sAMAccountName") & vbNewLine & _
                                   "userAccountControl: " & objRecordSet.Fields("userAccountControl") & vbNewLine & _
                                   "userWorkstations: " & objRecordSet.Fields("userWorkstations") & vbNewLine & _
                                   "profilePath: " & objRecordSet.Fields("profilePath")


        getBasicUserInfoAsString = getBasicUserInfoAsString & "scriptPath: " & objRecordSet.Fields("scriptPath") & vbNewLine & _
                                   "homeDirectory: " & objRecordSet.Fields("homeDirectory") & vbNewLine & _
                                   "homeDrive: " & objRecordSet.Fields("homeDrive") & vbNewLine & _
                                   "homePhone: " & objRecordSet.Fields("homePhone") & vbNewLine & _
                                   "info: " & objRecordSet.Fields("info") & vbNewLine & _
                                   "title: " & objRecordSet.Fields("title") & vbNewLine & _
                                   "department: " & objRecordSet.Fields("department") & vbNewLine & _
                                   "company: " & objRecordSet.Fields("company") & vbNewLine & _
                                   "manager: " & objRecordSet.Fields("manager")    '& vbNewLine &

        objRecordSet.MoveNext
    Loop

CleanUp:
    objConnection.Close
    Set objRootDSE = Nothing
    Set objCommand = Nothing
    Set objConnection = Nothing
    Set objRecordSet = Nothing

End Function

Function isRegisteredUser(ntUsername As String, domain As String) As Boolean

    Dim objRootDSE, strDNSDomain, objCommand, objConnection, strQuery
    Dim objRecordSet, strName, strDN
    Dim strBase, strFilter, strAttributes

    ' Determine DNS domain name from RootDSE object.
    
    strDNSDomain = getBBTADomainStr(domain)

    If strDNSDomain = "" Then
        isRegisteredUser = False
        Exit Function
    End If
    ' Use ADO to search Active Directory.
    Set objCommand = CreateObject("ADODB.Command")
    Set objConnection = CreateObject("ADODB.Connection")
    objConnection.Provider = "ADsDSOObject"
    'objConnection.Properties("User ID") = "DOMAIN\account"
    'objConnection.Properties("Password") = "xxxxxx"
    objConnection.Open "Active Directory Provider"
    objCommand.ActiveConnection = objConnection

    ' Search for all user objects. Sort recordset by DisplayName.
    strBase = "<LDAP://" & strDNSDomain & ">"
    strFilter = "(&(objectCategory=person)(objectClass=user)(sAMAccountName=" & ntUsername & "*))"
    strAttributes = "displayName"
    strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"

    objCommand.CommandText = strQuery
    objCommand.Properties("Page Size") = 100
    objCommand.Properties("Timeout") = 30
    objCommand.Properties("Cache Results") = False
    objCommand.Properties("Sort On") = "displayName"
    Set objRecordSet = objCommand.Execute

    If objRecordSet.EOF Then
        isRegisteredUser = False
        GoTo CleanUp
        Exit Function
    Else
        isRegisteredUser = True
    End If

CleanUp:
    objConnection.Close
    Set objRootDSE = Nothing
    Set objCommand = Nothing
    Set objConnection = Nothing
    Set objRecordSet = Nothing
End Function

Function getUserState(ntUsername As String, domain As String) As String
    Dim objRootDSE, strDNSDomain, objCommand, objConnection, strQuery
    Dim objRecordSet, strName, strDN
    Dim strBase, strFilter, strAttributes

    ' Determine DNS domain name from RootDSE object.
    
    strDNSDomain = getBBTADomainStr(domain)

    ' Use ADO to search Active Directory.
    Set objCommand = CreateObject("ADODB.Command")
    Set objConnection = CreateObject("ADODB.Connection")
    objConnection.Provider = "ADsDSOObject"
    'objConnection.Properties("User ID") = "DOMAIN\account"
    'objConnection.Properties("Password") = "xxxxxx"
    objConnection.Open "Active Directory Provider"
    objCommand.ActiveConnection = objConnection

    ' Search for all user objects. Sort recordset by DisplayName.
    strBase = "<LDAP://" & strDNSDomain & ">"
    strFilter = "(&(objectCategory=person)(objectClass=user)(sAMAccountName=" & ntUsername & "*))"
    strAttributes = "displayName,useraccountcontrol,msds-user-account-control-computed"
    strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"

    objCommand.CommandText = strQuery
    objCommand.Properties("Page Size") = 100
    objCommand.Properties("Timeout") = 30
    objCommand.Properties("Cache Results") = False
    objCommand.Properties("Sort On") = "displayName"
    Set objRecordSet = objCommand.Execute

    If objRecordSet.EOF Then
        getUserState = "User not found"
        GoTo CleanUp
        Exit Function
    End If

    ' Loop through results
    Do Until objRecordSet.EOF
        If objRecordSet.Fields("useraccountcontrol") = 512 Then
            getUserState = "NORMAL_ACCOUNT"
        End If
        If objRecordSet.Fields("useraccountcontrol") = 514 Then
            getUserState = "ACCOUNTDISABLE"
        End If
        If objRecordSet.Fields("msds-user-account-control-computed") = 8388608 Then
            getUserState = "PASSWORD_EXPIRED"
        End If
        If objRecordSet.Fields("msds-user-account-control-computed") = 16 Then
            getUserState = "LOCKOUT"
        End If
        'getUserState = "useraccountcontrol: " & objRecordSet.Fields("useraccountcontrol") & "  " & "msds-user-account-control-computed: " & objRecordSet.Fields("msds-user-account-control-computed")
        objRecordSet.MoveNext

    Loop

CleanUp:
    objConnection.Close
    Set objRootDSE = Nothing
    Set objCommand = Nothing
    Set objConnection = Nothing
    Set objRecordSet = Nothing
End Function

'Author: David Santana
'Recibe un string con el nombre del dominio tal y como aparece en la lista de dominios
'Devuelve un string con la ruta correcta del dominio para dialecto LDAP del banco de bogota
Function getBBTADomainStr(dominio) As String
    Dim objRootDSE
    
    dominio = Trim(dominio)
    
    'Para cada dominio, establecemos la ruta adecuada para la consulta LDAP
    If dominio = "BQUILLA" Then
        getBBTADomainStr = "bq.bancodebogota.net/DC=bq,DC=bancodebogota,DC=net"
        Exit Function
    End If
    
    
    If dominio = "CALI" Then
        getBBTADomainStr = "ca.bancodebogota.net/DC=ca,DC=bancodebogota,DC=net"
        Exit Function
    End If
    
    If dominio = "BMANGA" Then
        getBBTADomainStr = "bu.bancodebogota.net/DC=bu,DC=bancodebogota,DC=net"
        Exit Function
    End If
    
    If dominio = "MEDELLIN" Then
        getBBTADomainStr = "md.bancodebogota.net/DC=md,DC=bancodebogota,DC=net"
        Exit Function
    End If
    
    If dominio = "BOGOTA" Then
        getBBTADomainStr = "bo.bancodebogota.net/DC=bo,DC=bancodebogota,DC=net"
        Exit Function
    End If
    
    If dominio = "BANCODEBOGOTA" Then
        getBBTADomainStr = "DC=bancodebogota,DC=net"
        Exit Function
    End If
    
    'en cualquier otro caso, retornamos el dominio en el que se encuentre la maquina.
    'Set objRootDSE = GetObject("LDAP://RootDSE")
    'getBBTADomainStr = objRootDSE.Get("defaultNamingContext")
    
End Function
