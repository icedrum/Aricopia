VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmActualizarVers 
   BorderStyle     =   0  'None
   Caption         =   "Actualizador de Versiones Ariadna"
   ClientHeight    =   5490
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7950
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   7950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdContinuar 
      Caption         =   "&Continuar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4905
      Width           =   1275
   End
   Begin MSComctlLib.ProgressBar Pb1 
      Height          =   240
      Left            =   3510
      TabIndex        =   3
      Top             =   4680
      Width           =   4155
      _ExtentX        =   7329
      _ExtentY        =   423
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6390
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4905
      Width           =   1275
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1095
      Left            =   3510
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "frmActualizarVers.frx":0000
      Top             =   3645
      Width           =   4155
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   4815
      TabIndex        =   5
      Top             =   1890
      Width           =   1950
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Actualizando Aplicación :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   3510
      TabIndex        =   0
      Top             =   4230
      Visible         =   0   'False
      Width           =   4155
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   5520
      Left            =   0
      Picture         =   "frmActualizarVers.frx":0008
      Top             =   0
      Width           =   7965
   End
End
Attribute VB_Name = "frmActualizarVers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Aplicaciones As String
Dim NomPc As String
Dim HayPdte As Boolean
Dim cerrar As Boolean

Private Sub cmdAceptar_Click()
'Dim NomPc As String
Dim b As Boolean

    If Not HayPdte Then
        CmdContinuar_Click
        Exit Sub
    End If
    
    CmdAceptar.visible = False
    
    b = ActualizaPc(NomPc)

    CmdContinuar.visible = True

    If b Then
        Label1.Caption = "Proceso realizado correctamente"
    Else
        Label1.Caption = "No se ha podido realizar el proceso correctamente. Llame a Ariadna."
    End If
    
    Me.Refresh
    DoEvents
 
    CmdContinuar.SetFocus
    
End Sub

Private Sub CmdContinuar_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If cerrar Then Unload Me
End Sub

Private Sub Form_Load()
Dim cad As String, Cad1 As String
Dim mens As String
Dim b As Boolean

'++10/09/2008
        Me.CmdAceptar.visible = True
        Me.CmdContinuar.visible = False
        Me.Text1.visible = True
        Me.Text1.Text = ""
        Me.Label1.visible = False
        Me.Label1.Caption = ""
        Me.Pb1.visible = False
        cad = ""

'        Me.Label2.Caption = "Versión: " & App.Major & "." & App.Minor & "." & App.Revision
'        Me.Label2.Refresh
'        DoEvents
        
        NomPc = ComputerName

        cerrar = False

        b = ComprobarVersionesPcPrevia(NomPc, cad)

        If b Then
            If cad <> "" Then
                mens = AplicacionesporActualizar(cad)
                Text1.Text = mens
                Label1.Caption = mens
                Me.Refresh
                DoEvents
'                Me.CmdAceptar.SetFocus
            Else
                cerrar = True
            End If
        End If

End Sub



Private Function ActualizarVersionPC(pc As Integer, ByRef vAplic As CAplicacion) As Boolean
Dim Sql As String
Dim RS As ADODB.Recordset
Dim Destino As String
Dim PathDestino As String
Dim PathFuente As String
Dim Fuente As String
Dim Donde As String
Dim NumFic As Long



Dim Nombre As String



    On Error GoTo eActualizarVersionPC

    ActualizarVersionPC = False
    
    Me.Label1.Caption = "Actualizando aplicación: " & vAplic.NomAplic
    Me.Label1.Refresh
    DoEvents
    
    Sql = "select count(*) from ficheroscopia where idaplicacion = " & DBSet(vAplic.IdAplic, "N")
    NumFic = TotalRegistros(Sql) + 1
    Me.Pb1.visible = True
    CargarProgres Me.Pb1, CInt(NumFic)
    
    
    Sql = "select * from ficheroscopia where idaplicacion = " & DBSet(vAplic.IdAplic, "N")
    
    PathDestino = ""
    PathDestino = DevuelveDesdeBDNew(cPTours, "pcscopia", "pathcopia", "idpcs", CStr(pc), "N", , "idaplicacion", CStr(vAplic.IdAplic), "N")
    
'    If InStr(1, LCase(PathDestino), "archivos de programa\ariadna") = 0 Then
'        MsgBox "El path destino es incorrecto. No se va a realizar la actualización de la aplicación " & DBLet(vAplic.NomAplic, "T"), vbExclamation
'        Exit Function
'    End If
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    While Not RS.EOF
        IncrementarProgres Me.Pb1, 1
  
        Fuente = "\\" & vAplic.Servidor & "\" & Mid(vAplic.PathServ, InStr(1, vAplic.PathServ, "\") + 1, Len(vAplic.PathServ)) & "\" & DBLet(RS!Nombre, "T")
        Destino = PathDestino & "\" & DBLet(RS!Nombre, "T")
        
     
        
        If DBLet(RS!Tipo, "N") = 0 Then ' fichero
            ' si el fichero es el ejecutable lo procesaremos el ultimo
            If InStr(1, Fuente, "exe") = 0 Then
                If InStr(1, Destino, "*") Then
                    
                    Donde = "Dir " & Fuente
                    Nombre = Dir(Fuente)
                    PathFuente = "\\" & vAplic.Servidor & "\" & Mid(vAplic.PathServ, InStr(1, vAplic.PathServ, "\") + 1, Len(vAplic.PathServ))
                    Do While Len(Nombre)
                        
                        QuitarPropSoloLectura PathFuente & "\" & Nombre
                        Donde = "Copia: " & PathFuente & "\" & Nombre & " --> " & PathDestino & "\" & Nombre
                        CopiarArchivo PathFuente & "\" & Nombre, PathDestino & "\" & Nombre  ' si no existiera el fichero
                        Nombre = Dir
                    Loop
                Else
                   
                    Donde = "Copia: " & Fuente & " --> " & Destino
                    CopiarArchivo Fuente, Destino ' si no existiera el fichero
                    
                End If
            End If
        Else
             ' carpeta
            If Dir(Destino, vbDirectory) = "" Then
                Donde = "Crea carpeta:" & Destino
                MkDir Destino
            End If
             
             
            Donde = "Dir: " & Fuente & "\"
            Nombre = Dir(Fuente & "\")
            PathFuente = "\\" & vAplic.Servidor & "\" & Mid(vAplic.PathServ, InStr(1, vAplic.PathServ, "\") + 1, Len(vAplic.PathServ))
            
    
            
            Do While Len(Nombre)
                
                QuitarPropSoloLectura Destino & "\" & Nombre
                
                CopiarArchivo Fuente & "\" & Nombre, Destino & "\" & Nombre
                

                Nombre = Dir
            Loop
             

        End If
            

       
    
        RS.MoveNext
    Wend
    
    Set RS = Nothing
    
    'dejamos para ultimo lugar el exe
    Sql = "select * from ficheroscopia where nombre like '%exe' and idaplicacion = " & DBSet(vAplic.IdAplic, "N")
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not RS.EOF Then
        IncrementarProgres Me.Pb1, 1
        
  
        Fuente = "\\" & vAplic.Servidor & "\" & Mid(vAplic.PathServ, InStr(1, vAplic.PathServ, "\") + 1, Len(vAplic.PathServ)) & "\" & DBLet(RS!Nombre, "T")
        Destino = PathDestino & "\" & DBLet(RS!Nombre, "T")
        Donde = "Copia: " & Fuente & " --> " & Destino
        
        
        If DBLet(RS!Tipo, "N") = 0 Then ' fichero
            QuitarPropSoloLectura PathFuente & "\" & Nombre
            CopiarArchivo Fuente, Destino

        End If
        
    End If
    
    Set RS = Nothing
    
    Me.Pb1.visible = False

    ActualizarVersionPC = True
    Exit Function
    
eActualizarVersionPC:
    If Err.Number <> 0 Then
        Screen.MousePointer = vbDefault
        
        MuestraError Err.Number, "Error actualizando version de " & vAplic.NomAplic & vbCrLf & vbCrLf & Donde
    End If
End Function
 



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
    
End Function



Private Function ComprobarVersionesPcPrevia(pc As String, ByRef cad As String) As Boolean
Dim Sql As String
Dim RS As ADODB.Recordset
Dim Aplic As String
Dim vAplic As CAplicacion
Dim Version As String
Dim fichero As String
Dim b As Boolean


    On Error GoTo eComprobarVersionesPCPrevia


    ComprobarVersionesPcPrevia = False

    Sql = "select pcscopia.* from pcscopia, pcs where ucase(pcs.nompc) = " & DBSet(UCase(pc), "T")
    Sql = Sql & " and pcscopia.idpcs = pcs.idpcs "
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    b = True
    
    cad = ""
    
    While Not RS.EOF
        Set vAplic = New CAplicacion
        If vAplic.LeerDatos(RS!idaplicacion) Then
            
            
            
            
            '
            'fichero = DBLet(RS!pathcopia, "T") & "\" & vAplic.Ejecutable
            'Version = A.GetFileVersion(fichero)
            Version = RS!version1 & "." & RS!version2 & ".0." & RS!version3
            
            
            If Version <> "" Then
                If VersionSuperior(vAplic.UltVers, Version) = True Then
                    cad = cad & DBLet(RS!idaplicacion, "N") & ":0|"
                End If
            Else
                cad = cad & DBLet(RS!idaplicacion, "N") & ":1|"
                b = False
            End If
            

        End If
        Set vAplic = Nothing
        RS.MoveNext
    Wend

    Set RS = Nothing

    ComprobarVersionesPcPrevia = True
    Exit Function

eComprobarVersionesPCPrevia:
    If Err.Number <> 0 Then
        Screen.MousePointer = vbDefault
    
        MuestraError Err.Number, "Comprobando Versiones Previa", Err.Description
    End If
End Function

Public Function AplicacionesporActualizar(cad As String) As String
Dim Cad1 As String
Dim Cad2 As String
Dim Resul As String
Dim J As Integer
Dim i As Integer
Dim mens As String
Dim Mens1 As String
Dim longitud As Integer

Dim Aplic As String
Dim Situ As String

    AplicacionesporActualizar = ""

    mens = "Se va a proceder a actualizar las siguientes aplicaciones: " & vbCrLf '& vbCrLf
    
    Mens1 = ""
    
    longitud = Len(cad)
    
    i = 0
    Cad1 = cad
    
    Aplicaciones = ""
    
    While Len(Cad1) <> 0
        i = InStr(1, Cad1, "|")
        
        If i <> 0 Then
            Cad2 = Mid(Cad1, 1, i)
            J = InStr(1, Cad2, ":")
            
            Aplic = Mid(Cad2, 1, J - 1)
            Situ = Mid(Cad2, J + 1, 1)
            
            If CInt(Situ) = 0 Then ' situacion sin actualizar
                Aplicaciones = Aplicaciones & Aplic & ","
                
                mens = mens & DevuelveDesdeBDNew(cPTours, "aplicaciones", "nombre", "idaplicacion", Aplic, "N")
                mens = mens & "    " '& vbCrLf
            Else  ' situacion de no poder actualizar
                Mens1 = Mens1 & DevuelveDesdeBDNew(cPTours, "aplicaciones", "nombre", "idaplicacion", Aplic, "N")
                Mens1 = Mens1 & "    " '& vbCrLf
            End If
            
            If Len(Cad1) <> i Then
                Cad1 = Mid(Cad1, i + 1, Len(Cad1))
            Else
                Cad1 = ""
            End If
        End If
    
    Wend

    Resul = ""
    If Aplicaciones <> "" Then
        Aplicaciones = Mid(Aplicaciones, 1, Len(Aplicaciones) - 1) ' quitamos la ultima coma
        Resul = mens & vbCrLf '& vbCrLf
    End If
    If Mens1 <> "" Then
'        Resul = Resul & "Las siguientes aplicaciones no tienen versión y no se actualizarán: " & vbCrLf '& vbCrLf
        Resul = Resul & "No se encontró archivo ejecutable y no se actualizarán: " & vbCrLf '& vbCrLf
        Resul = Resul & Mens1
    End If
    
    HayPdte = True
    If Aplicaciones = "" Then HayPdte = False
        
    
    AplicacionesporActualizar = Resul
    
End Function

Private Function ActualizaPc(pc As String) As Boolean
Dim Sql As String
Dim RS As ADODB.Recordset
Dim Aplic As String
Dim vAplic As CAplicacion
Dim Version As String
Dim fichero As String
Dim b As Boolean

    On Error GoTo eActualizaPc

    Me.Label1.Caption = ""
    Me.Label1.visible = True
    Me.Text1.visible = False
    
    Me.Refresh
    DoEvents
    
    ActualizaPc = False

    If Aplicaciones = "" Then Exit Function
    
    Sql = "select pcscopia.* from pcscopia, pcs where ucase(pcs.nompc) = " & DBSet(UCase(pc), "T")
    Sql = Sql & " and pcscopia.idaplicacion in (" & Aplicaciones & ") "
    Sql = Sql & " and pcscopia.idpcs = pcs.idpcs "
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    b = True
    
    While Not RS.EOF And b
        Set vAplic = New CAplicacion
        If vAplic.LeerDatos(RS!idaplicacion) Then
               b = ActualizarVersionPC(DBLet(RS!idpcs, "N"), vAplic)
               If b Then ActualizaTablaConLaVersion DBLet(RS!idpcs, "N"), vAplic
        End If
        Set vAplic = Nothing
        RS.MoveNext
    Wend

    Set RS = Nothing

    ActualizaPc = b
    Exit Function

eActualizaPc:
    If Err.Number <> 0 Then
        Screen.MousePointer = vbDefault
    
        MuestraError Err.Number, "Actualizando Versiones", Err.Description
    End If
End Function

Private Function ActualizaTablaConLaVersion(pc As Integer, ByRef vAplic As CAplicacion) As Boolean
Dim Sql As String
Dim cad As String

    Sql = vAplic.UltVers & "."
    Sql = Replace(Sql, ".", "|")
    cad = "UPDATE pcscopia set version1 =" & RecuperaValor(Sql, 1)
    cad = cad & " ,version2 =" & RecuperaValor(Sql, 2) & ", version3 =" & RecuperaValor(Sql, 4)
    cad = cad & " WHERE idpcs =" & pc & " AND idaplicacion =" & vAplic.IdAplic
    'No pogno error
    conn.Execute cad
End Function





Private Sub QuitarPropSoloLectura(Nombre)
    On Error Resume Next
    If Nombre = "" Then Exit Sub
    'If Dir(nombre, vbArchive) = "" Then Exit Sub
    Shell "attrib  -r """ & Nombre & """", vbHide
    If Err.Number <> 0 Then
        
        Err.Clear
    End If
End Sub


Private Sub CopiarArchivo(vOrigen As String, vDestino As String)
    'No pongo rutina de errores
    FileCopy vOrigen, vDestino  ' si no existiera el fichero
    
End Sub
