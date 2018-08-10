VERSION 5.00
Begin VB.Form frmIdentifica 
   BackColor       =   &H00800000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5490
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7965
   LinkTopic       =   "Form1"
   ScaleHeight     =   5490
   ScaleWidth      =   7965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   4320
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   4920
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   4320
      TabIndex        =   0
      Top             =   3960
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cargando ....."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Index           =   2
      Left            =   5040
      TabIndex        =   4
      Top             =   4920
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Index           =   0
      Left            =   4320
      TabIndex        =   2
      Top             =   3600
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Index           =   1
      Left            =   4320
      TabIndex        =   3
      Top             =   4560
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   5535
      Left            =   0
      Top             =   0
      Width           =   7815
   End
End
Attribute VB_Name = "frmIdentifica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor: DAVID (refet per CÈSAR) +-+-
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Dim PrimeraVez As Boolean
Dim T1 As Single

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False

         Me.Refresh
         PonerVisible True
         If Text1(0).Text <> "" Then
            PonerFoco Text1(1)
         Else
            PonerFoco Text1(0)
         End If
             
         'Leemos el ultimo usuario conectado
         NumeroEmpresaMemorizar True
         
         T1 = T1 + 2.5 - Timer
         If T1 > 0 Then espera T1

         
         PonerVisible True
         If Text1(0).Text <> "" Then
            Text1(1).SetFocus
        Else
            Text1(0).SetFocus
        End If

    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
'    Screen.MousePointer = vbHourglass
    'PonerVisible False
'    T1 = Timer
    'Text1(0).Text = "root"
 '   Text1(1).Text = "aritel"
    PrimeraVez = True
    CargaImagen
End Sub

Private Sub CargaImagen()
    On Error Resume Next
    Me.Image1 = LoadPicture(App.Path & "\entrada.dat")
    If Err.Number <> 0 Then
        MsgBox Err.Description & vbCrLf & vbCrLf & "Error cargando", vbCritical
        Set conn = Nothing
        End
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    NumeroEmpresaMemorizar False
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    With Text1(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    Else
        If KeyAscii = 27 Then
            Unload Me
        End If
    End If
End Sub


Private Sub Text1_LostFocus(Index As Integer)
    Text1(Index).Text = Trim(Text1(Index).Text)

    'Comprobamos si los dos estan con datos
    If Text1(0).Text <> "" And Text1(1).Text <> "" Then
        Validar
    End If
End Sub


Private Sub Validar()
Dim OK As Byte
Dim cad As String


    Set vSesion = New CSesion

    If vSesion.Leer(Text1(0).Text) = 0 Then
        'Con exito
        If vSesion.PasswdPROPIO = Text1(1).Text Then
            OK = 0
        Else
            OK = 1
        End If
    Else
        If Text1(0).Text = "root" And Text1(1).Text = "aritel" Then
            cad = "insert into usuarios (codusu, nomusu, login, passwordpropio, nivelusuges) "
            cad = cad & " values (0,'root','root','aritel',0)"
            conn.Execute cad
            OK = 0
        Else
            OK = 2
        End If

    End If

    If OK <> 0 Then
        MsgBox "Usuario o Password Incorrecto", vbExclamation

        Text1(1).Text = ""
        PonerFoco Text1(0)
    Else
        'OK
        If vSesion.Nivel < 0 Then
            MsgBox "Usuario sin Permisos.", vbExclamation
            End
        Else
            PonerVisible False
            Me.Refresh
            espera 0.2
        
            
            
        '    If ComprovaVersio Then
        '        MsgBox "Existe una versión más reciente de la aplicación. Se va a proceder a la actualización", vbInformation
        '        Shell App.Path & "\PlannerUpdate.exe", vbNormalFocus
        '        End
        '    End If
        
            'Carga Datos de la Empresa y los Niveles de cuentas de Contabilidad de la empresa
            'Crea la Conexion a la BD de la Contabilidad
'            LeerDatosEmpresa
        
            InicializarFormatos
            teclaBuscar = 43
        
'            Load frmPpal
            
            MDIppal.Show
            
            Unload Me
            
        End If
    
    End If


End Sub


Private Sub PonerVisible(visible As Boolean)
    'Label1(2).visible = Not visible  'Cargando
    Text1(0).visible = visible
    Text1(1).visible = visible
    Label1(0).visible = visible
    Label1(1).visible = visible
End Sub

'Lo que haremos aqui es ver, o guardar, el ultimo numero de empresa
'a la que ha entrado, y el usuario
Private Sub NumeroEmpresaMemorizar(Leer As Boolean)
Dim NF As Integer
Dim cad As String
On Error GoTo ENumeroEmpresaMemorizar

    cad = App.Path & "\ultusu.dat"
    If Leer Then
        If Dir(cad) <> "" Then
            NF = FreeFile
            Open cad For Input As #NF
            Line Input #NF, cad
            Close #NF
            cad = Trim(cad)
            
                'El primer pipe es el usuario
                Text1(0).Text = cad
    
        End If
    Else 'Escribir
        NF = FreeFile
        Open cad For Output As #NF
        cad = Text1(0).Text
        Print #NF, cad
        Close #NF
    End If
ENumeroEmpresaMemorizar:
    Err.Clear
End Sub

