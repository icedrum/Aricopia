VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSesion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Clase donde se guarda la informacion de la sesion establecida.
'se crea cuando se hace login y aqui se guardan los datos de
'quien hizo login, el codigo y nombre del trabajador

'variables locales que contienen valores de propiedad
Private mvarCodigo As Long 'copia local
Private mvarCodusu As Long 'copia local
Private mvarNombre As String 'copia local
Private mvarPasswd As String 'copia local
Private mvarPasswdPROPIO As String 'copia local
Private mvarlogin As String 'copia local
Private mvarDirFich As String 'copia local
Private mvarNivel As Long 'copia local
Private mvarCadenaConexion As String
Private mvarPC As String

Private RS As ADODB.Recordset
Private SQL As String
Private Mens As String

'------------------------------------------------
'Propiedades del modulo CSesion
'------------------------------------------------
Public Property Let Nivel(ByVal vData As Long)
    mvarNivel = vData
End Property


Public Property Get Nivel() As Long
    Nivel = mvarNivel
End Property

Public Property Let Passwd(ByVal vData As String)
    mvarPasswd = vData
End Property


Public Property Get Passwd() As String
    Passwd = mvarPasswd
End Property



Public Property Let CadenaConexion(ByVal vData As String)
    mvarCadenaConexion = vData
End Property


Public Property Get CadenaConexion() As String
    CadenaConexion = mvarCadenaConexion
End Property



Public Property Let PasswdPROPIO(ByVal vData As String)
'se usa cuando se asigna un valor a una propiedad, en el lado izquierdo de la asignaci�n.
'Syntax: X.Passwd = 5
    mvarPasswdPROPIO = vData
End Property


Public Property Get PasswdPROPIO() As String
'se usa cuando se asigna un valor a una propiedad, en el lado derecho de la asignaci�n.
'Syntax: Debug.Print X.Passwd
    PasswdPROPIO = mvarPasswdPROPIO
End Property


Public Property Let Nombre(ByVal vData As String)
'se usa cuando se asigna un valor a una propiedad, en el lado izquierdo de la asignaci�n.
'Syntax: X.Nombre= 5
    mvarNombre = vData
End Property


Public Property Get Nombre() As String
'se usa cuando se asigna un valor a una propiedad, en el lado derecho de la asignaci�n.
'Syntax: Debug.Print X.Nombre
    Nombre = mvarNombre
End Property

'---------------------------
'Dirfich
Public Property Let Dirfich(ByVal vData As String)
'se usa cuando se asigna un valor a una propiedad, en el lado izquierdo de la asignaci�n.
'Syntax: X.Nombre= 5
    mvarDirFich = vData
End Property


Public Property Get Dirfich() As String
'se usa cuando se asigna un valor a una propiedad, en el lado derecho de la asignaci�n.
'Syntax: Debug.Print X.Nombre
    Dirfich = mvarDirFich
End Property

'Login
Public Property Let Login(ByVal vData As String)
'se usa cuando se asigna un valor a una propiedad, en el lado izquierdo de la asignaci�n.
'Syntax: X.Nombre= 5
    mvarlogin = vData
End Property


Public Property Get Login() As String
'se usa cuando se asigna un valor a una propiedad, en el lado derecho de la asignaci�n.
'Syntax: Debug.Print X.Nombre
    Login = mvarlogin
End Property



Public Property Let Codigo(ByVal vData As Long)
'se usa cuando se asigna un valor a una propiedad, en el lado izquierdo de la asignaci�n.
'Syntax: X.Codigo= 5
    mvarCodigo = vData
End Property


Public Property Get Codigo() As Long
'se usa cuando se asigna un valor a una propiedad, en el lado derecho de la asignaci�n.
'Syntax: Debug.Print X.Codigo
    Codigo = mvarCodigo
End Property

Public Property Let Codusu(ByVal vData As Long)
'se usa cuando se asigna un valor a una propiedad, en el lado izquierdo de la asignaci�n.
'Syntax: X.Codigo= 5
    mvarCodusu = vData
End Property


Public Property Get Codusu() As Long
'se usa cuando se asigna un valor a una propiedad, en el lado derecho de la asignaci�n.
'Syntax: Debug.Print X.Codigo
    Codusu = mvarCodusu
End Property


Public Property Get pc() As String
'se usa cuando se asigna un valor a una propiedad, en el lado derecho de la asignaci�n.
'Syntax: Debug.Print X.Codigo
    pc = mvarPC
End Property



Public Function Leer(ByVal vlogin As String) As Byte
Dim i As Long
    On Error GoTo Err1
    SQL = "SELECT * FROM usuarios.usuarios WHERE login = '" & vlogin & "'"
    'SQL = SQL & " AND Nivelusuges >=0"
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If RS.EOF Then
        Leer = 1
    Else
        i = 0 'DevuelveAumentoPC
        mvarCodigo = Val(DBLet(RS!Codusu)) + i
        mvarCodusu = DBLet(RS!Codusu)
        mvarNombre = DBLet(RS!nomusu)
        mvarNivel = 0 ' Val(DBLet(RS!nivelusuges))
        mvarPasswdPROPIO = RS!passwordpropio
        mvarlogin = vlogin
        mvarCadenaConexion = ""  'Se le asigna en tiempo de dise�o
        Leer = 0
    End If
    RS.Close
    Exit Function
Err1:
    Mens = "Error: " & Err.Number & " : " & Err.Description
    MsgBox Mens, vbExclamation
    RS.Close
    Leer = 1
End Function


Private Function DevuelveAumentoPC() As Long

DevuelveAumentoPC = 32000
'mvarPC = ComputerName
mvarPC = ComputerName
SQL = DevuelveDesdeBD("codpc", "pcs", "nompc", mvarPC, "T")
If SQL <> "" Then DevuelveAumentoPC = Val(SQL) * 1000

End Function



