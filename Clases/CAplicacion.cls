VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAplicacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ### [Monica] 25/07/2008


Option Explicit

'Aplicaciones
'TABLA: aplicaciones


'ATRIBUTOS
'Variables locales que contienen valores de propiedad
Private mIdAplic  As String 'id aplicacion
Private mNomAplic As String 'Nombre aplicacion
Private mUltVers As String  'Ultima Version
Private mServidor As String 'Servidor
Private mPathServ As String 'Path Servidor
Private mEjecutable As String 'Ejecutable de la aplicacion

'------------------------------------------------
'Propiedades del modulo CAplicacion
'------------------------------------------------

'**** Id Aplicacion

Public Property Let IdAplic(ByVal vData As String)
     mIdAplic = vData
End Property

Public Property Get IdAplic() As String
     IdAplic = mIdAplic
End Property

'**** Nombre Aplicacion

Public Property Let NomAplic(ByVal vData As String)
     NomAplic = vData
End Property

Public Property Get NomAplic() As String
     NomAplic = mNomAplic
End Property


'**** Ultima Version

Public Property Let UltVers(ByVal vData As String)
     mUltVers = vData
End Property

Public Property Get UltVers() As String
     UltVers = mUltVers
End Property


'**** Servidor
Public Property Let Servidor(ByVal vData As String)
     mServidor = vData
End Property

Public Property Get Servidor() As String
     Servidor = mServidor
End Property


'**** Path del Servidor

Public Property Let PathServ(ByVal vData As String)
     mPathServ = vData
End Property

Public Property Get PathServ() As String
     PathServ = mPathServ
End Property

'**** Ejecutable

Public Property Let Ejecutable(ByVal vData As String)
     mEjecutable = vData
End Property

Public Property Get Ejecutable() As String
     Ejecutable = mEjecutable
End Property



'------------------------------------------------
'Procedimientos del modulo CAplicacion
'------------------------------------------------

'Public Function Existe(vCodsocio As String) As Boolean
''Comprueba si existe el socio en la BD
'Dim devuelve As String
'
'    On Error GoTo EExiste
'
'    devuelve = DevuelveDesdeBD("codsocio", "ssocio", "codsocio", vCodsocio, "N")
'    If devuelve = "" Then
'        Existe = False
'        devuelve = "No existe el socio " & vCodsocio & ". "
'        MsgBox devuelve, vbExclamation
'    Else
'        Existe = True
'    End If
'
'EExiste:
'    If Err.Number <> 0 Then Existe = False
'End Function



Public Function LeerDatos(vAplicacion As String) As Boolean
'Leer los datos de una aplicacion dada
'Lee de la BD: CopiaVersiones, Tabla: aplicaciones
'OUT: True si lee los datos correctamente
Dim Rs As ADODB.Recordset
Dim Sql As String

    On Error GoTo ELeer
    
    LeerDatos = False
    
    Sql = "SELECT idaplicacion, nombre, ultimaversion, servidor, pathservidor, ejecutable "
    Sql = Sql & " FROM aplicaciones "
    Sql = Sql & " WHERE idaplicacion=" & DBSet(vAplicacion, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    If Rs.EOF Then
        LeerDatos = False
    Else
        mIdAplic = CStr(Rs!idaplicacion)
        
        mNomAplic = Rs!Nombre
        mUltVers = Rs!ultimaversion
        mServidor = Rs!Servidor
        mPathServ = Rs!pathservidor
        mEjecutable = Rs!Ejecutable
        LeerDatos = True
    End If

    Rs.Close
    Set Rs = Nothing
    Exit Function

ELeer:
    Sql = "Se ha producido un error. " & "Datos aplicación" & vbCrLf
    Sql = Sql & "Número: " & Err.Number & vbCrLf
    Sql = Sql & "Descripción: " & Err.Description
    MsgBox Sql, vbExclamation
    Set Rs = Nothing
    LeerDatos = False
End Function



Public Function ActualizarVersionAplic(vAplicacion As String, vVersion As String) As Boolean
'Actualiza la ultima version de una aplicacion
Dim Sql As String
Dim devuelve As String

    On Error GoTo EActualizarversion
    ActualizarVersionAplic = False
    
    devuelve = DevuelveDesdeBDNew("aplicaciones", "ultimaversion", "idaplicacion", vAplicacion, "T")
    
    If devuelve <> "" Then 'Existe la aplicacion y la actualizamos
        Sql = "UPDATE aplicaciones SET ultimaversion=" & DBSet(vVersion, "T")
        Sql = Sql & " WHERE idaplicacion=" & DBSet(vAplicacion, "N")
'    Else 'No existe el socio y lo añadimos
'        Sql = "INSERT INTO ssocio (nifsocio, nomsocio, domsocio, codpobla, pobsocio, prosocio, telsocio) "
'        Sql = Sql & " VALUES (" & DBSet(vNIF, "T") & ", " & DBSet(Nombre, "T") & ", " & DBSet(Domicilio, "T") & ", "
'        Sql = Sql & DBSet(CPostal, "T") & ", " & DBSet(Poblacion, "T") & ", " & DBSet(Provincia, "T") & ", " & DBSet(Tfnosocio, "T") & ")"
    End If
    conn.Execute Sql, , adCmdText
    
    ActualizarVersionAplic = True
    
EActualizarversion:
    If Err.Number <> 0 Then
        ActualizarVersionAplic = False
        MuestraError Err.Number, "Actualizando versión de la Aplicación.", Err.Description
    End If
End Function


