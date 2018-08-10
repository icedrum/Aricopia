Attribute VB_Name = "bus"
'NOTA: en este mòdul, ademés, n'hi han funcions generals que no siguen de formularis (molt bé)
Option Explicit

'Definicion Conexión a BASE DE DATOS
'---------------------------------------------------
'Conexión a la BD Avnics de la empresa
Public conn As ADODB.Connection

'Que conexion a base de datos se va a utilizar
'Public Const cPTours As Byte = 1 'trabajaremos con conn (conexion a BD Avnics)
Public Const cConta As Byte = 2 'trabajaremos con conn (conexion a BD Conta)


Public cerrar_ As Boolean 'Si hay que cerrar el programa pq no hay nada para actualizar


'Definicion de clases de la aplicación
'-----------------------------------------------------
'Public vEmpresa As Cempresa  'Los datos de la empresa
'Public vParamAplic As CParamAplic   'parametros de la aplicacion
Public vSesion As CSesion   'Los datos del usuario que hizo login
Public vConfig As Configuracion



'Definicion de FORMATOS
'---------------------------------------------------
Public FormatoFecha As String
Public FormatoHora As String
Public FormatoImporte As String 'Decimal(12,2)
Public FormatoPrecio As String 'Decimal(8,3)
'Public FormatoCantidad As String 'Decimal(10,2)
Public FormatoPorcen As String 'Decimal(5,2) 'Porcentajes
Public FormatoExp As String  'Expedientes

'Public FormatoDec10d2 As String 'Decimal(10,2)
'Public FormatoDec10d3 As String 'Decimal(10,3)
'Public FormatoDec5d4 As String 'Decimal(5,4)

'Public FIni As String
'Public FFin As String

'Public FIniSeg As String 'fecha de inicio de ejercicio de la contabilidad de Seguros
'Public FFinSeg As String 'fecha de fin de ejercicio de la contabilidad de Seguros

'Public FIniTel As String 'fecha de inicio de ejercicio de la contabilidad de Telefonia
'Public FFinTel As String 'fecha de fin de ejercicio de la contabilidad de Telefonia

'Public FormatoKms As String 'Decimal(8,4)


Public teclaBuscar As Integer 'llamada desde prismaticos

Public CadenaDesdeOtroForm As String

'Global para nº de registro eliminado
Public NumRegElim  As Long

'publica para almacenar control cambios en registros de formularios
'se utiliza en InsertarCambios
Public CadenaCambio As String
Public ValorAnterior As String

Public MensError As String

'Para algunos campos de texto suletos controlarlos
'Public miTag As CTag

'Variable para saber si se ha actualizado algun asiento
'Public AlgunAsientoActualizado As Boolean
'Public TieneIntegracionesPendientes As Boolean

'Public miRsAux As ADODB.Recordset

Public AnchoLogin As String  'Para fijar los anchos de columna

Public Aplicaciones As String



'Inicio Aplicación
Public Sub Main()

Dim NomPc As String
Dim Servidor As String
Dim CadenaParametros As String
Dim cad As String, Cad1 As String
Dim Mens As String
Dim b As Boolean

    If App.PrevInstance Then
        MsgBox "Aricopia ya se esta ejecutando", vbExclamation
        End
    End If

    Set vConfig = New Configuracion
    If vConfig.Leer = 1 Then

         MsgBox "MAL CONFIGURADO", vbCritical
         End
         Exit Sub
    End If

     'obric la conexio
    If AbrirConexion() = False Then
        MsgBox "La aplicación no puede continuar sin acceso a los datos. ", vbCritical
        End
    End If
    
    'recoge los parametros del .exe
    CadenaParametros = Command
    
    If CadenaParametros = "/A" Then
        ' ejecutamos la aplicacion desde el servidor
        Load frmIdentifica
        'CadenaDesdeOtroForm = ""
        
        'Necesitaremos el archivo login.dat
        frmIdentifica.Show
    Else
        'frmActualizarVers.Show vbModal
        Load frmNuevoActualizar
        If cerrar_ Then
            Unload frmNuevoActualizar
            Set conn = Nothing
            End
        Else
            frmNuevoActualizar.Show
        End If
    
        
    End If
End Sub


'espera els segon que li digam
Public Function espera(Segundos As Single)
    Dim T1
    T1 = Timer
    Do
    Loop Until Timer - T1 > Segundos
End Function


Public Function AbrirConexion() As Boolean
Dim cad As String
On Error GoTo EAbrirConexion
    
    AbrirConexion = False
    Set conn = Nothing
    Set conn = New Connection
    'Conn.CursorLocation = adUseClient
    conn.CursorLocation = adUseServer
'    cad = "DSN=plannertours;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=plannertours;SERVER=srvcentral;UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"

'    cad = "DSN=arigasol;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=arigasol;UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"
'--monica
'    cad = "DSN=vAriagroutil;DESC=MySQL ODBC 3.51 Driver DSN;UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"
    
'++ de david
'    cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATA SOURCE= vCopiav"
'    cad = cad & ";UID=" & Usuario
'    cad = cad & ";PWD=" & Pass
'    cad = cad & ";Persist Security Info=true"
    
    cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATABASE=aricopia;SERVER=" & vConfig.SERVER
    cad = cad & ";UID=" & vConfig.User
    cad = cad & ";PWD=" & vConfig.password
    '---- Laura: 29/09/2006
    cad = cad & ";PORT=" & vConfig.puerto & ";OPTION=3;STMT=;"
    cad = cad & ";Persist Security Info=true"
    
    
         conn.ConnectionString = cad
    conn.Open
  
    
    
    
    
    AbrirConexion = True
    Exit Function
    
EAbrirConexion:
    MuestraError Err.Number, "Abrir conexión.", Err.Description
End Function



'
'Public Sub LeerDatosEmpresa()
''Crea instancia de la clase Cempresa con los valores en
''Tabla: Empresas
''BDatos: PTours y Conta
'
'    Set vEmpresa = New Cempresa
'    If vEmpresa.LeerDatos(1) = False Then  'De Avnics
'        MsgBox "No se han podido cargar los datos de la empresa. Debe configurar la aplicación.", vbExclamation
'        Set vEmpresa = Nothing
'       ' Set vSesion = Nothing
'       ' Set conn = Nothing
'        Exit Sub
'    End If
'
'    ' ### [Monica] 06/09/2006
'    ' añadido
'    Set vParamAplic = New CParamAplic
'    If vParamAplic.Leer = 1 Then
'        MsgBox "No se han podido cargar los parámetros de contabilidad. Debe configurar la aplicación.", vbExclamation
'
'        Set vParamAplic = Nothing
'        Exit Sub
'    End If
'
''    Set vParam = New Cparametros
''    If vParam.Leer = False Then   'De AriGasol
''        MsgBox "No se han podido cargar los parámetros de la empresa. Debe configurar la aplicación.", vbExclamation
''        Set vEmpresa = Nothing
''        Set vSesion = Nothing
''        Set Conn = Nothing
''        End
''    End If
'End Sub


Public Function PonerDatosPpal()
    'If Not vEmpresa Is Nothing Then
        MDIppal.Caption = "AriVersion" & " ver. " & App.Major & "." & App.Minor & "." & App.Revision '& "   -  Empresa: " & vEmpresa.nomEmpre
    'End If
    'If Err.Number <> 0 Then MuestraError Err.Description, "Poniendo datos de la pantalla principal", Err.Description
End Function

    

Public Sub MuestraError(numero As Long, Optional cadena As String, Optional Desc As String)
    Dim cad As String
    Dim Aux As String
    
    'Con este sub pretendemos unificar el msgbox para todos los errores
    'que se produzcan
    On Error Resume Next
    cad = "Se ha producido un error: " & vbCrLf
    If cadena <> "" Then
        cad = cad & vbCrLf & cadena & vbCrLf & vbCrLf
    End If
    'Numeros de errores que contolamos
    If conn.Errors.Count > 0 Then
        ControlamosError Aux
        conn.Errors.Clear
    Else
        Aux = ""
    End If
    If Aux <> "" Then Desc = Aux
    If Desc <> "" Then cad = cad & vbCrLf & Desc & vbCrLf & vbCrLf
    If Aux = "" Then cad = cad & "Número: " & numero & vbCrLf & "Descripción: " & Error(numero)
    MsgBox cad, vbExclamation
End Sub

Public Function DBSet(vData As Variant, Tipo As String, Optional EsNulo As String) As Variant
'Establece el valor del dato correcto antes de Insertar en la BD
Dim cad As String

        If IsNull(vData) Then
            DBSet = ValorNulo
            Exit Function
        End If

        If Tipo <> "" Then
            Select Case Tipo
                Case "T"    'Texto
                    If vData = "" Then
                        If EsNulo = "N" Then
                            DBSet = "''"
                        Else
                            DBSet = ValorNulo
                        End If
                    Else
                        cad = (CStr(vData))
                        NombreSQL cad
                        DBSet = "'" & cad & "'"
                    End If
                    
                Case "N"    'Numero
                    If vData = "" Or vData = 0 Then
                        If EsNulo <> "" Then
                            If EsNulo = "S" Then
                                DBSet = ValorNulo
                            Else
                                DBSet = 0
                            End If
                        Else
                            DBSet = 0
                        End If
                    Else
                        cad = CStr(ImporteFormateado(CStr(vData)))
                        DBSet = TransformaComasPuntos(cad)
                    End If
                    
                Case "F"    'Fecha
'                     '==David
''                    DBLet = "0:00:00"
'                     '==Laura
                    If vData = "" Then
                        If EsNulo = "S" Then
                            DBSet = ValorNulo
                        Else
                            DBSet = "'1900-01-01'"
                        End If
                    Else
                        DBSet = "'" & Format(vData, FormatoFecha) & "'"
                    End If
                    
                Case "FH" 'Fecha/Hora
                    If vData = "" Then
                        If EsNulo = "S" Then DBSet = ValorNulo
                    Else
                        DBSet = "'" & Format(vData, "yyyy-mm-dd hh:mm:ss") & "'"
                    End If
                    
                Case "H" 'Hora
                    If vData = "" Then
                    Else
                        DBSet = "'" & Format(vData, "hh:mm:ss") & "'"
                    End If
                    
                Case "B"  'Boolean
                    If vData Then
                        DBSet = 1
                    Else
                        DBSet = 0
                    End If
            End Select
        End If
End Function

Public Function DBLetMemo(vData As Variant) As Variant
    On Error Resume Next
    
    DBLetMemo = vData
    
    
    
    If Err.Number <> 0 Then
        Err.Clear
        DBLetMemo = ""
    End If
End Function



Public Function DBLet(vData As Variant, Optional Tipo As String) As Variant
'Para cuando recupera Datos de la BD
    If IsNull(vData) Then
        DBLet = ""
        If Tipo <> "" Then
            Select Case Tipo
                Case "T"    'Texto
                    DBLet = ""
                Case "N"    'Numero
                    DBLet = 0
                Case "F"    'Fecha
                     '==David
'                    DBLet = "0:00:00"
                     '==Laura
'                     DBLet = "0000-00-00"
                      DBLet = ""
                Case "D"
                    DBLet = 0
                Case "B"  'Boolean
                    DBLet = False
                Case Else
                    DBLet = ""
            End Select
        End If
    Else
        DBLet = vData
    End If
End Function

'/////////////////////////////////////////////////
'   Esto lo ejecutaremos justo antes de bloquear
'   Prepara la conexion para bloquear
Public Sub PreparaBloquear()
    conn.Execute "commit"
    conn.Execute "set autocommit=0"
End Sub

'/////////////////////////////////////////////////
'   Esto lo ejecutaremos justo despues de un bloque
'   Prepara la conexion para bloquear
Public Sub TerminaBloquear()
    conn.Execute "commit"
    conn.Execute "set autocommit=1"
End Sub

'///////////////////////////////////////////////////////////////
'
'   Cogemos un numero formateado: 1.256.256,98  y deevolvemos 1256256,98
'   Tiene que venir numérico
Public Function ImporteFormateado(Importe As String) As Currency
Dim i As Integer

    If Importe = "" Then
        ImporteFormateado = 0
    Else
        'Primero quitamos los puntos
        Do
            i = InStr(1, Importe, ".")
            If i > 0 Then Importe = Mid(Importe, 1, i - 1) & Mid(Importe, i + 1)
        Loop Until i = 0
        ImporteFormateado = Importe
    End If
End Function

' ### [Monica] 11/09/2006
Public Function ImporteSinFormato(cadena As String) As String
Dim i As Integer
'Quitamos puntos
Do
    i = InStr(1, cadena, ".")
    If i > 0 Then cadena = Mid(cadena, 1, i - 1) & Mid(cadena, i + 1)
Loop Until i = 0
ImporteSinFormato = TransformaPuntosComas(cadena)
End Function



'Cambia los puntos de los numeros decimales
'por comas
Public Function TransformaComasPuntos(cadena As String) As String
Dim i As Integer
    Do
        i = InStr(1, cadena, ",")
        If i > 0 Then
            cadena = Mid(cadena, 1, i - 1) & "." & Mid(cadena, i + 1)
        End If
    Loop Until i = 0
    TransformaComasPuntos = cadena
End Function

'Para los nombre que pueden tener ' . Para las comillas habra que hacer dentro otro INSTR
Public Sub NombreSQL(ByRef cadena As String)
Dim J As Integer
Dim i As Integer
Dim Aux As String
    J = 1
    Do
        i = InStr(J, cadena, "'")
        If i > 0 Then
            Aux = Mid(cadena, 1, i - 1) & "\"
            cadena = Aux & Mid(cadena, i)
            J = i + 2
        End If
    Loop Until i = 0
End Sub

Public Function EsFechaOKString(ByRef T As String) As Boolean
Dim cad As String
    
    cad = T
    If InStr(1, cad, "/") = 0 Then
        If Len(T) = 8 Then
            cad = Mid(cad, 1, 2) & "/" & Mid(cad, 3, 2) & "/" & Mid(cad, 5)
        Else
            If Len(T) = 6 Then cad = Mid(cad, 1, 2) & "/" & Mid(cad, 3, 2) & "/" & Mid(cad, 5)
        End If
    End If
    If IsDate(cad) Then
        EsFechaOKString = True
        T = Format(cad, "dd/mm/yyyy")
    Else
        EsFechaOKString = False
    End If
End Function

Public Function DevNombreSQL(cadena As String) As String
Dim J As Integer
Dim i As Integer
Dim Aux As String
    J = 1
    Do
        i = InStr(J, cadena, "'")
        If i > 0 Then
            Aux = Mid(cadena, 1, i - 1) & "\"
            cadena = Aux & Mid(cadena, i)
            J = i + 2
        End If
    Loop Until i = 0
    DevNombreSQL = cadena
End Function


Public Function DevuelveDesdeBD(kCampo As String, Ktabla As String, Kcodigo As String, ValorCodigo As String, Optional Tipo As String, Optional ByRef otroCampo As String) As String
    Dim rs As Recordset
    Dim cad As String
    Dim Aux As String
    
    On Error GoTo EDevuelveDesdeBD
    DevuelveDesdeBD = ""
    cad = "Select " & kCampo
    If otroCampo <> "" Then cad = cad & ", " & otroCampo
    cad = cad & " FROM " & Ktabla
    cad = cad & " WHERE " & Kcodigo & " = "
    If Tipo = "" Then Tipo = "N"
    Select Case Tipo
    Case "N"
        'No hacemos nada
        cad = cad & ValorCodigo
    Case "T", "F"
        cad = cad & "'" & ValorCodigo & "'"
    Case Else
        MsgBox "Tipo : " & Tipo & " no definido", vbExclamation
        Exit Function
    End Select
    
    
    
    'Creamos el sql
    Set rs = New ADODB.Recordset
    rs.Open cad, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not rs.EOF Then
        DevuelveDesdeBD = DBLet(rs.Fields(0))
        If otroCampo <> "" Then otroCampo = DBLet(rs.Fields(1))
    End If
    rs.Close
    Set rs = Nothing
    Exit Function
EDevuelveDesdeBD:
        MuestraError Err.Number, "Devuelve DesdeBD.", Err.Description
End Function




'
Public Function DevuelveDesdeBDNew(vBD As Byte, Ktabla As String, kCampo As String, Kcodigo1 As String, valorCodigo1 As String, Optional tipo1 As String, Optional ByRef otroCampo As String, Optional KCodigo2 As String, Optional ValorCodigo2 As String, Optional tipo2 As String, Optional KCodigo3 As String, Optional ValorCodigo3 As String, Optional tipo3 As String) As String
'IN: vBD --> Base de Datos a la que se accede
Dim rs As Recordset
Dim cad As String
Dim Aux As String

On Error GoTo EDevuelveDesdeBDnew
    DevuelveDesdeBDNew = ""
'    If valorCodigo1 = "" And ValorCodigo2 = "" Then Exit Function
    cad = "Select " & kCampo
    If otroCampo <> "" Then cad = cad & ", " & otroCampo
    cad = cad & " FROM " & Ktabla
    If Kcodigo1 <> "" Then
        cad = cad & " WHERE " & Kcodigo1 & " = "
        If tipo1 = "" Then tipo1 = "N"
    Select Case tipo1
        Case "N"
            'No hacemos nada
            cad = cad & Val(valorCodigo1)
        Case "T"
            cad = cad & DBSet(valorCodigo1, "T")
        Case "F"
            cad = cad & DBSet(valorCodigo1, "F")
        Case Else
            MsgBox "Tipo : " & tipo1 & " no definido", vbExclamation
            Exit Function
    End Select
    End If

    If KCodigo2 <> "" Then
        cad = cad & " AND " & KCodigo2 & " = "
        If tipo2 = "" Then tipo2 = "N"
        Select Case tipo2
        Case "N"
            'No hacemos nada
            If ValorCodigo2 = "" Then
                cad = cad & "-1"
            Else
                cad = cad & Val(ValorCodigo2)
            End If
        Case "T"
'            cad = cad & "'" & ValorCodigo2 & "'"
            cad = cad & DBSet(ValorCodigo2, "T")
        Case "F"
            cad = cad & "'" & Format(ValorCodigo2, FormatoFecha) & "'"
        Case Else
            MsgBox "Tipo : " & tipo2 & " no definido", vbExclamation
            Exit Function
        End Select
    End If

    If KCodigo3 <> "" Then
        cad = cad & " AND " & KCodigo3 & " = "
        If tipo3 = "" Then tipo3 = "N"
        Select Case tipo3
        Case "N"
            'No hacemos nada
            If ValorCodigo3 = "" Then
                cad = cad & "-1"
            Else
                cad = cad & Val(ValorCodigo3)
            End If
        Case "T"
            cad = cad & "'" & ValorCodigo3 & "'"
        Case "F"
            cad = cad & "'" & Format(ValorCodigo3, FormatoFecha) & "'"
        Case Else
            MsgBox "Tipo : " & tipo3 & " no definido", vbExclamation
            Exit Function
        End Select
    End If


    'Creamos el sql
    Set rs = New ADODB.Recordset
    'Select Case vBD
    '    Case cPTours   'BD 1: Ariges
            rs.Open cad, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    'End Select

    If Not rs.EOF Then
        DevuelveDesdeBDNew = DBLet(rs.Fields(0))
        If otroCampo <> "" Then otroCampo = DBLet(rs.Fields(1))
    End If
    rs.Close
    Set rs = Nothing
    Exit Function

EDevuelveDesdeBDnew:
        MuestraError Err.Number, "Devuelve DesdeBD.", Err.Description
End Function
'




'CESAR
Public Function DevuelveDesdeBDnew2(kBD As Integer, kCampo As String, Ktabla As String, Kcodigo As String, ValorCodigo As String, Optional Tipo As String, Optional num As Byte, Optional ByRef otroCampo As String) As String
Dim rs As Recordset
Dim cad As String
Dim Aux As String
Dim v_aux As Integer
Dim campo As String
Dim Valor As String
Dim tip As String

On Error GoTo EDevuelveDesdeBDnew2
DevuelveDesdeBDnew2 = ""

cad = "Select " & kCampo
If otroCampo <> "" Then cad = cad & ", " & otroCampo
cad = cad & " FROM " & Ktabla

If Kcodigo <> "" Then cad = cad & " where "

For v_aux = 1 To num
    campo = RecuperaValor(Kcodigo, v_aux)
    Valor = RecuperaValor(ValorCodigo, v_aux)
    tip = RecuperaValor(Tipo, v_aux)
        
    cad = cad & campo & "="
    If tip = "" Then Tipo = "N"
    
    Select Case tip
            Case "N"
                'No hacemos nada
                cad = cad & Valor
            Case "T", "F"
                cad = cad & "'" & Valor & "'"
            Case Else
                MsgBox "Tipo : " & tip & " no definido", vbExclamation
            Exit Function
    End Select
    
    If v_aux < num Then cad = cad & " AND "
  Next v_aux

'Creamos el sql
Set rs = New ADODB.Recordset
Select Case kBD
    Case 1
        rs.Open cad, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
End Select

If Not rs.EOF Then
    DevuelveDesdeBDnew2 = DBLet(rs.Fields(0))
    If otroCampo <> "" Then otroCampo = DBLet(rs.Fields(1))
Else
     If otroCampo <> "" Then otroCampo = ""
End If
rs.Close
Set rs = Nothing
Exit Function
EDevuelveDesdeBDnew2:
    MuestraError Err.Number, "Devuelve DesdeBDnew2.", Err.Description
End Function


Public Function EsEntero(Texto As String) As Boolean
Dim i As Integer
Dim C As Integer
Dim L As Integer
Dim res As Boolean

    res = True
    EsEntero = False

    If Not IsNumeric(Texto) Then
        res = False
    Else
        'Vemos si ha puesto mas de un punto
        C = 0
        L = 1
        Do
            i = InStr(L, Texto, ".")
            If i > 0 Then
                L = i + 1
                C = C + 1
            End If
        Loop Until i = 0
        If C > 1 Then res = False
        
        'Si ha puesto mas de una coma y no tiene puntos
        If C = 0 Then
            L = 1
            Do
                i = InStr(L, Texto, ",")
                If i > 0 Then
                    L = i + 1
                    C = C + 1
                End If
            Loop Until i = 0
            If C > 1 Then res = False
        End If
        
    End If
        EsEntero = res
End Function

Public Function TransformaPuntosComas(cadena As String) As String
    Dim i As Integer
    Do
        i = InStr(1, cadena, ".")
        If i > 0 Then
            cadena = Mid(cadena, 1, i - 1) & "," & Mid(cadena, i + 1)
        End If
        Loop Until i = 0
    TransformaPuntosComas = cadena
End Function

Public Sub InicializarFormatos()
    FormatoFecha = "yyyy-mm-dd"
    FormatoHora = "hh:mm:ss"
'    FormatoFechaHora = "yyyy-mm-dd hh:mm:ss"
    FormatoImporte = "#,###,###,##0.00"  'Decimal(12,2)
    FormatoPrecio = "##,##0.000"  'Decimal(8,3) antes decimal(10,4)
'    FormatoCantidad = "##,###,##0.00"   'Decimal(10,2)
    FormatoPorcen = "##0.00" 'Decima(5,2) para porcentajes
    
  '  FormatoDec10d2 = "##,###,##0.00"   'Decimal(10,2)
  '  FormatoDec10d3 = "##,###,##0.000"   'Decimal(10,3)
   ' FormatoDec5d4 = "0.0000"   'Decimal(5,4)
    FormatoExp = "0000000000"
'    FormatoKms = "#,##0.00##" 'Decimal(8,4)
End Sub


Public Sub AccionesCerrar()
'cosas que se deben hacen cuando finaliza la aplicacion
    On Error Resume Next
    
    'cerrar clases q estan abiertas durante la ejecucion
   ' Set vEmpresa = Nothing
    Set vSesion = Nothing
    
'    Set vParam = Nothing
'    Set vParamAplic = Nothing
'    Set vParamConta = Nothing
    
    
    'Cerrar Conexiones a bases de datos
    conn.Close
    Set conn = Nothing
    
    If Err.Number <> 0 Then Err.Clear
End Sub

'Public Sub AccionesCerrarContabilidades()
''cosas que se deben hacen cuando finaliza la aplicacion
'    On Error Resume Next
'
'    'cerrar clases q estan abiertas durante la ejecucion
'    Set vEmpresa = Nothing
'
''    Set vParam = Nothing
''    Set vParamAplic = Nothing
''    Set vParamConta = Nothing
'
'
'    If Err.Number <> 0 Then Err.Clear
'End Sub



Public Function OtrosPCsContraAplicacion() As String
Dim MiRS As Recordset
Dim cad As String
Dim Equipo As String

    Set MiRS = New ADODB.Recordset
    cad = "show processlist"
    MiRS.Open cad, conn, adOpenKeyset, adLockOptimistic, adCmdText
    cad = ""
    While Not MiRS.EOF
        If MiRS.Fields(3) = vSesion.CadenaConexion Then
            Equipo = MiRS.Fields(2)
            'Primero quitamos los dos puntos del puerot
            NumRegElim = InStr(1, Equipo, ":")
            If NumRegElim > 0 Then Equipo = Mid(Equipo, 1, NumRegElim - 1)
            
            'El punto del dominio
            NumRegElim = InStr(1, Equipo, ".")
            If NumRegElim > 0 Then Equipo = Mid(Equipo, 1, NumRegElim - 1)
            
            Equipo = UCase(Equipo)
            
            If Equipo <> vSesion.Codusu Then
                    If Equipo <> "LOCALHOST" Then
                        If InStr(1, cad, Equipo & "|") = 0 Then cad = cad & Equipo & "|"
                    End If
            End If
        End If
        'Siguiente
        MiRS.MoveNext
    Wend
    NumRegElim = 0
    MiRS.Close
    Set MiRS = Nothing
    OtrosPCsContraAplicacion = cad
End Function


Public Function UsuariosConectados2() As Boolean
Dim i As Integer
Dim cad As String
Dim metag As String
Dim SQL As String
cad = OtrosPCsContraAplicacion
UsuariosConectados2 = False
If cad <> "" Then
    UsuariosConectados2 = True
    i = 1
  '  metag = "Los siguientes PC's están conectados a: " & vEmpresa.nomEmpre & " (" & vSesion.CadenaConexion & ")" & vbCrLf & vbCrLf
    Do
        SQL = RecuperaValor(cad, i)
        If SQL <> "" Then
            metag = metag & "    - " & SQL & vbCrLf
            i = i + 1
        End If
    Loop Until SQL = ""
    MsgBox metag, vbExclamation
End If
End Function

'
'Public Function AplicacionesporActualizar(cad As String) As String
'Dim Cad1 As String
'Dim Cad2 As String
'Dim Resul As String
'Dim J As Integer
'Dim i As Integer
'Dim Mens As String
'Dim Mens1 As String
'Dim longitud As Integer
'
'Dim Aplic As String
'Dim Situ As String
'
'
'    AplicacionesporActualizar = ""
'
'    Mens = "Tiene las siguientes aplicaciones por actualizar: " & vbCrLf & vbCrLf
'
'    Mens1 = ""
'
'    longitud = Len(cad)
'
'    i = 0
'    Cad1 = cad
'
'    Aplicaciones = ""
'
'    While Len(Cad1) <> 0
'        i = InStr(1, Cad1, "|")
'
'        If i <> 0 Then
'            Cad2 = Mid(Cad1, 1, i)
'            J = InStr(1, Cad2, ":")
'
'            Aplic = Mid(Cad2, 1, J - 1)
'            Situ = Mid(Cad2, J + 1, 1)
'
'            If CInt(Situ) = 0 Then ' situacion sin actualizar
'                Aplicaciones = Aplicaciones & Aplic & ","
'
'                Mens = Mens & DevuelveDesdeBDNew(cPTours, "aplicaciones", "nombre", "idaplicacion", Aplic, "N")
'                Mens = Mens & vbCrLf
'            Else  ' situacion de no poder actualizar
'                Mens1 = Mens1 & DevuelveDesdeBDNew(cPTours, "aplicaciones", "nombre", "idaplicacion", Aplic, "N")
'                Mens1 = Mens1 & vbCrLf
'            End If
'
'            If Len(Cad1) <> i Then
'                Cad1 = Mid(Cad1, i + 1, Len(Cad1))
'            Else
'                Cad1 = ""
'            End If
'        End If
'
'    Wend
'
'    Resul = ""
'    If Aplicaciones <> "" Then
'        Aplicaciones = Mid(Aplicaciones, 1, Len(Aplicaciones) - 1) ' quitamos la ultima coma
'        Resul = Mens & vbCrLf & vbCrLf
'    End If
'    If Mens1 <> "" Then
'        Resul = Resul & "Las siguientes aplicaciones no tienen versión y no se actualizarán: " & vbCrLf & vbCrLf
'        Resul = Resul & Mens1
'    End If
'
'    AplicacionesporActualizar = Resul
'
'End Function


Private Function VersionSuperior(V1 As String, V2 As String) As Boolean
Dim i As Integer
Dim J As Integer

    If InStr(1, V1, ".") = 0 Then
        If InStr(1, V2, ".") = 0 Then
            VersionSuperior = (CInt(Mid(V1, 1, Len(V1))) > CInt(Mid(V2, 1, Len(V2))))
        Else
            VersionSuperior = (Mid(V1, 1, Len(V1)) > Mid(V2, 1, InStr(1, V2, ".") - 1))
        End If
    Else
        If InStr(1, V2, ".") = 0 Then
            VersionSuperior = (Mid(V1, 1, InStr(1, V1, ".") - 1) > Mid(V2, 1, Len(V2)))
        Else
            If Mid(V1, 1, InStr(1, V1, ".") - 1) = Mid(V2, 1, InStr(1, V2, ".") - 1) Then
                VersionSuperior = VersionSuperior(Mid(V1, InStr(1, V1, ".") + 1, Len(V1)), Mid(V2, InStr(1, V2, ".") + 1, Len(V2)))
            Else
                VersionSuperior = (CInt(Mid(V1, 1, InStr(1, V1, ".") - 1)) > CInt(Mid(V2, 1, InStr(1, V2, ".") - 1)))
            End If
        End If
    End If
    
'    If Mid(V1, 1, InStr(1, V1, ".") - 1) = Mid(V2, 1, InStr(1, V2, ".") - 1) Then
'        VersionSuperior = VersionSuperior(Mid(V1, InStr(1, V1, ".") + 1, Len(V1)), Mid(V2, InStr(1, V2, ".") + 1, Len(V2)))
'    Else
'    End If
End Function





