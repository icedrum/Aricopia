VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNuevoActualizar 
   BorderStyle     =   0  'None
   Caption         =   "Aricopia. Ariadna SW"
   ClientHeight    =   6660
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10740
   Icon            =   "frmNuevoActualizar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   10740
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar pb1 
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   6240
      Visible         =   0   'False
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9360
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6240
      Width           =   1275
   End
   Begin VB.CommandButton CmdContinuar 
      Caption         =   "&Continuar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9360
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6240
      Width           =   1275
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2535
      Left            =   240
      TabIndex        =   0
      Top             =   3480
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   4471
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Aplicacion"
         Object.Width           =   7126
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Version PC"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Version"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Cambios"
         Object.Width           =   2926
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "pathExplorer"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNuevoActualizar.frx":1BCC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNuevoActualizar.frx":1C99C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNuevoActualizar.frx":231FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNuevoActualizar.frx":29A60
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNuevoActualizar.frx":29D7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNuevoActualizar.frx":305DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNuevoActualizar.frx":308F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNuevoActualizar.frx":37158
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNuevoActualizar.frx":3D9BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNuevoActualizar.frx":4421C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNuevoActualizar.frx":4AA7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNuevoActualizar.frx":67D4B
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNuevoActualizar.frx":6A4FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNuevoActualizar.frx":6FCEF
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNuevoActualizar.frx":77A31
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image Image2 
      Height          =   840
      Left            =   8400
      Picture         =   "frmNuevoActualizar.frx":7E293
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2400
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   2775
      Left            =   120
      Top             =   3360
      Width           =   10575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "AriCopia"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1170
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   4560
   End
   Begin VB.Label lblI 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   4680
      TabIndex        =   3
      Top             =   6300
      Width           =   5895
   End
   Begin VB.Image Image1 
      Height          =   3075
      Left            =   0
      Picture         =   "frmNuevoActualizar.frx":7F49E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   19200
   End
End
Attribute VB_Name = "frmNuevoActualizar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Aplicaciones As String
Dim NomPc As String
Dim HayPdte2 As Boolean



Dim TamanyoTotal As Long
Dim llevoCopiado As Long

Dim TiempoTranscurrido As Single

Private Sub cmdAceptar_Click()
'Dim NomPc As String
Dim b As Boolean

    If Not HayPdte2 Then
        CmdContinuar_Click
        Exit Sub
    End If
    
    CmdAceptar.visible = False
    
    b = ActualizaPc(NomPc)

    CmdContinuar.visible = True

    If b Then
        lblI.Caption = "Proceso realizado correctamente"
    Else
        lblI.Caption = "No  se  ha  podido  realizar  el  proceso  correctamente."
        lblI.Left = 240
        lblI.Width = 8000
        lblI.ForeColor = vbRed
    End If
    
    Me.Refresh
    DoEvents
 
    CmdContinuar.SetFocus
    
End Sub

Private Sub CmdContinuar_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Activate()
    If cerrar_ Then
        Unload Me
    Else
        PonerFocoBtn CmdAceptar
    End If
End Sub

Private Sub Form_Load()
Dim cad As String, Cad1 As String
Dim Mens As String
Dim Actualizar As Integer
        
        Set Me.ListView1.SmallIcons = ImageList1
        
        
        Me.CmdAceptar.visible = True
        Me.CmdContinuar.visible = False
        Me.ListView1.ListItems.Clear
        lblI.Caption = ""   '    Ariadna Software"
       
        cad = ""
    
        NomPc = ComputerName

        cerrar_ = False

        Actualizar = ComprobarVersionesPcPrevia(NomPc, cad)

        If Actualizar >= 0 Then
            If Actualizar > 0 Then
                'mens = AplicacionesporActualizar(cad)
                Me.Refresh
                HayPdte2 = True
                DoEvents
'                Me.CmdAceptar.SetFocus
            Else
                cerrar_ = True
            End If
        Else
            MsgBox "Avise soporte tecnico", vbExclamation
        End If

End Sub

Private Sub Lbl2()
    On Error Resume Next
    lblI.Refresh
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Function ActualizarVersionPC(pc As Integer, ByRef vAplic As CAplicacion) As Boolean
Dim SQL As String
Dim RS As ADODB.Recordset
Dim Destino As String
Dim PathDestino As String
Dim PathFuente As String
Dim Fuente As String
Dim Donde As String
Dim NumFic As Long
Dim EsElExeDeLaAplicacion As Boolean
Dim k As Integer   '2 veces el proceso
Dim tma As Long
Dim Nombre As String

Dim C As Collection
Dim M As Long

    On Error GoTo eActualizarVersionPC

    ActualizarVersionPC = False
    
    Me.lblI.Caption = "Actualizando aplicación: " & vAplic.NomAplic
    Lbl2
    DoEvents
    
    SQL = "select count(*) from ficheroscopia where idaplicacion = " & DBSet(vAplic.IdAplic, "N")
    NumFic = TotalRegistros(SQL) + 1
    
    'CargarProgres Me.pb1, CInt(NumFic)
    Set RS = New ADODB.Recordset
    
    TiempoTranscurrido = Timer
    TamanyoTotal = 0
    pb1.Value = 0
    For k = 1 To 2
        'primera vez "saber tamaño"
        If k = 1 Then
            pb1.visible = False
            
            Me.lblI.Caption = vAplic.NomAplic
            Lbl2
        Else
            
            pb1.visible = True
            Me.lblI.Caption = vAplic.NomAplic & "... copiando"
            Lbl2
            espera 2
            
        End If
        DoEvents

            
            
        SQL = "select * from ficheroscopia where idaplicacion = " & DBSet(vAplic.IdAplic, "N")
        
        PathDestino = ""
        PathDestino = DevuelveDesdeBDNew(1, "pcscopia", "pathcopia", "idpcs", CStr(pc), "N", , "idaplicacion", CStr(vAplic.IdAplic), "N")
    
        
        
        RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        
        While Not RS.EOF
        
            Fuente = "\\" & vAplic.Servidor & "\" & Mid(vAplic.PathServ, InStr(1, vAplic.PathServ, "\") + 1, Len(vAplic.PathServ)) & "\" & DBLet(RS!Nombre, "T")
            Destino = PathDestino & "\" & DBLet(RS!Nombre, "T")
            
            Me.lblI.Caption = RS!Nombre
            Lbl2
            
            If DBLet(RS!Tipo, "N") = 0 Then ' fichero
            
                ' si el fichero es el ejecutable lo procesaremos el ultimo
                EsElExeDeLaAplicacion = False
                
                If LCase(Right(Fuente, 3)) = "exe" Then
                   
                   '
                   If InStr(1, LCase(Fuente), LCase(vAplic.Ejecutable)) > 0 Then
                        'Stop
                        EsElExeDeLaAplicacion = True
                    End If
                End If
                If Not EsElExeDeLaAplicacion Then
                
                
                    
                
                    If InStr(1, Destino, "*") Then
                        Donde = "Poniendo nombre"
                        PathFuente = "\\" & vAplic.Servidor & "\" & Mid(vAplic.PathServ, InStr(1, vAplic.PathServ, "\") + 1, Len(vAplic.PathServ))
                        
                        
                        
                        Donde = "Dir (*): " & Fuente
                        Nombre = Dir(Fuente)
                        Set C = New Collection
                        Do While Len(Nombre) > 0
                            'Me.lblI.Caption = Nombre
                            Lbl2
                            If Nombre <> "" Then C.Add Nombre
                            Nombre = Dir
                        Loop
                                      
                        
                        
                        lblI.Caption = vAplic.NomAplic & "    Procesando ficheros (" & k & ")"
                        Lbl2
                        For M = 1 To C.Count
                            Nombre = C.Item(M)
                            Donde = "Filelen: " & PathFuente & "\" & Nombre
                            
                            tma = FileLen(PathFuente & "\" & Nombre)
                            tma = tma / 1024
                            If k = 1 Then
                                Donde = "Leyendo nº " & M & ": " & PathFuente & "\" & Nombre & "   " & tma & "  " & TamanyoTotal
                                
                                Lbl2
                                Donde = "En nº " & M & ": " & PathFuente & "\" & Nombre & "   " & tma & " -- " & TamanyoTotal
                                TamanyoTotal = TamanyoTotal + CLng(tma)
                                
                            Else
                                
                                llevoCopiado = llevoCopiado + tma
                                PonPb1
                    
                                QuitarPropSoloLectura Destino & "\" & Nombre
                                
                                Donde = "Copia: " & PathFuente & "\" & Nombre & " --> " & PathDestino & "\" & Nombre
                                CopiarArchivo PathFuente & "\" & Nombre, PathDestino & "\" & Nombre  ' si no existiera el fichero
                                
                            End If
                            
                        Next M
                        Set C = Nothing
                    Else
                        tma = FileLen(Fuente)
                        tma = tma / 1024
                        If k = 1 Then
                            TamanyoTotal = TamanyoTotal + tma
                        Else
                            llevoCopiado = llevoCopiado + tma
                            PonPb1
                
                            Donde = "Copia: " & Fuente & " --> " & Destino
                            CopiarArchivo Fuente, Destino ' si no existiera el fichero
                            
                            If Timer - TiempoTranscurrido > 3 Then
                                  DoEvents
                                  TiempoTranscurrido = Timer
                            End If
 
                            
                            
                        End If
                    End If
                Else
                    If k = 1 Then
                        
                        tma = FileLen(Fuente)
                        tma = tma / 1024
                        TamanyoTotal = TamanyoTotal + tma
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
                
                lblI.Caption = vAplic.NomAplic & "     " & RS!Nombre
                
                Do While Len(Nombre)
                    
                    
                    Donde = lblI.Caption
                    Lbl2
                    tma = FileLen(Fuente & "\" & Nombre)
                    tma = tma / 1024
                    If k = 1 Then
                        TamanyoTotal = TamanyoTotal + tma
                    Else
                        llevoCopiado = llevoCopiado + tma
                        PonPb1
                        Donde = "fileCOPY " & Fuente & "\" & Nombre & vbCrLf & Destino & "\" & Nombre
                        QuitarPropSoloLectura Destino & "\" & Nombre
                        CopiarArchivo Fuente & "\" & Nombre, Destino & "\" & Nombre
                    End If
                    
                    
                    Nombre = Dir
                    
                    If Timer - TiempoTranscurrido > 5 Then
                          DoEvents
                          TiempoTranscurrido = Timer
                    End If
                
                Loop
                 
    
            End If
                
    
           
        
            RS.MoveNext
        Wend
        RS.Close
        
        lblI.Caption = "... leyendo SRV"
        Lbl2
    Next k
    Set RS = Nothing
    
    'dejamos para ultimo lugar el exe
    lblI.Caption = "... copiando EXE"
    Lbl2
    SQL = "select * from ficheroscopia where nombre like '%exe' and idaplicacion = " & DBSet(vAplic.IdAplic, "N")
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not RS.EOF
        'IncrementarProgres Me.pb1, 1
        
    
                        'Stop
        Fuente = "\\" & vAplic.Servidor & "\" & Mid(vAplic.PathServ, InStr(1, vAplic.PathServ, "\") + 1, Len(vAplic.PathServ)) & "\" & DBLet(RS!Nombre, "T")
        
        If InStr(1, LCase(Fuente), LCase(vAplic.Ejecutable)) > 0 Then
        
            Destino = PathDestino & "\" & DBLet(RS!Nombre, "T")
            Donde = "Copia: " & Fuente & " --> " & Destino
            
            
            
            
            
            
            If DBLet(RS!Tipo, "N") = 0 Then ' fichero
                QuitarPropSoloLectura Destino
                CopiarArchivo Fuente, Destino
    
            End If
        End If
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    
    
    'PONGO ACTUALIZADO
    For k = 1 To Me.ListView1.ListItems.Count
        If Me.ListView1.ListItems(k).Tag = vAplic.IdAplic Then
            ListView1.ListItems(k).SubItems(1) = "Actualizado"
            ListView1.ListItems(k).ListSubItems(1).ForeColor = vbBlue
            ListView1.ListItems(k).ListSubItems(1).Bold = True
            ListView1.Refresh
            Exit For
        End If
        
    Next k
    
   pb1.visible = False
    pb1.Value = 0
    llevoCopiado = 0

    ActualizarVersionPC = True
    Exit Function
    
eActualizarVersionPC:
    If Err.Number <> 0 Then
        Screen.MousePointer = vbDefault
        
        MuestraError Err.Number, "Error actualizando version de " & vAplic.NomAplic & vbCrLf & vbCrLf & Donde, Err.Description
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



Private Function ComprobarVersionesPcPrevia(pc As String, ByRef cad As String) As Integer
Dim SQL As String
Dim RS As ADODB.Recordset
Dim Aplic As String
Dim vAplic As CAplicacion
Dim Version As String
Dim fichero As String
Dim b As Boolean
Dim iT As ListItem

    On Error GoTo eComprobarVersionesPCPrevia


    ComprobarVersionesPcPrevia = -1

    SQL = "select pcscopia.*,nombre,pathAyuda from pcscopia, pcs,aplicaciones where ucase(pcs.nompc) = " & DBSet(UCase(pc), "T")
    SQL = SQL & " and pcscopia.idpcs = pcs.idpcs and pcscopia.idaplicacion=aplicaciones.idaplicacion"
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    b = True
    
    cad = ""
    ListView1.ListItems.Clear
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
                
                    Set iT = Me.ListView1.ListItems.Add
                    iT.Text = RS!Nombre
                    iT.Tag = RS!idaplicacion
                    iT.SubItems(1) = Version
                    iT.SubItems(2) = vAplic.UltVers
                    iT.SmallIcon = DevuelveIcono2(vAplic.Ejecutable)
                    
                    cad = DBLet(RS!pathAyuda, "T")
                    If cad = "" Then
                        iT.SubItems(3) = ""
                    Else
                        iT.SubItems(3) = "Ver cambios"
                    End If
                    iT.SubItems(4) = cad
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

    ComprobarVersionesPcPrevia = ListView1.ListItems.Count
    Exit Function

eComprobarVersionesPCPrevia:
    If Err.Number <> 0 Then
        Screen.MousePointer = vbDefault
    
        MuestraError Err.Number, "Comprobando Versiones Previa", Err.Description
    End If
End Function


Private Function DevuelveIcono2(ByVal Nombre As String)
Dim Ico As Integer
    
    Nombre = LCase(Nombre)
    Nombre = Replace(Nombre, ".exe", "")
     
    Ico = 1
    Select Case Nombre
    Case "aritaxi"
        Ico = 2
    Case "ariges4"
        Ico = 3
    Case "contab"
        Ico = 4
    Case "arimoney"
        Ico = 5
    Case "arigasol"
        Ico = 6
    Case "arioli"
        Ico = 7
    Case "ariagroutil"
        Ico = 8
    Case "ariagrorec"
        Ico = 9
    Case "ariagro"
        Ico = 10
    Case "ariconta6"
        Ico = 11
    Case "aripres4"
        Ico = 15
    Case "arigestion"
        Ico = 1
    End Select
    DevuelveIcono2 = Ico
End Function
'
'
'Public Function AplicacionesporActualizar(Cad As String) As String
'Dim Cad1 As String
'Dim Cad2 As String
'Dim Resul As String
'Dim J As Integer
'Dim i As Integer
'Dim mens As String
'Dim Mens1 As String
'Dim longitud As Integer
'
'Dim Aplic As String
'Dim Situ As String
'
'    AplicacionesporActualizar = ""
'
'    mens = "Se va a proceder a actualizar las siguientes aplicaciones: " & vbCrLf '& vbCrLf
'
'    Mens1 = ""
'
'    longitud = Len(Cad)
'
'    i = 0
'    Cad1 = Cad
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
'                mens = mens & DevuelveDesdeBDNew(cPTours, "aplicaciones", "nombre", "idaplicacion", Aplic, "N")
'                mens = mens & "    " '& vbCrLf
'            Else  ' situacion de no poder actualizar
'                Mens1 = Mens1 & DevuelveDesdeBDNew(cPTours, "aplicaciones", "nombre", "idaplicacion", Aplic, "N")
'                Mens1 = Mens1 & "    " '& vbCrLf
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
'        Resul = mens & vbCrLf '& vbCrLf
'    End If
'    If Mens1 <> "" Then
''        Resul = Resul & "Las siguientes aplicaciones no tienen versión y no se actualizarán: " & vbCrLf '& vbCrLf
'        Resul = Resul & "No se encontró archivo ejecutable y no se actualizarán: " & vbCrLf '& vbCrLf
'        Resul = Resul & Mens1
'    End If
'
'    HayPdte2 = True
'    If Aplicaciones = "" Then HayPdte2 = False
'
'
'    AplicacionesporActualizar = Resul
'
'End Function

Private Function ActualizaPc(pc As String) As Boolean
Dim SQL As String
Dim RS As ADODB.Recordset
Dim Aplic As String
Dim vAplic As CAplicacion
Dim Version As String
Dim fichero As String
Dim b As Boolean
Dim J As Integer

    On Error GoTo eActualizaPc

    Me.lblI.Caption = ""
   
    
    Me.Refresh
    DoEvents
    
    ActualizaPc = False

    If ListView1.ListItems.Count = 0 Then Exit Function
    
    Aplicaciones = ""
    For J = 1 To ListView1.ListItems.Count
        Aplicaciones = Aplicaciones & ", " & ListView1.ListItems(J).Tag
    Next
    Aplicaciones = Mid(Aplicaciones, 2) 'quitamos la coma
    SQL = "select pcscopia.* from pcscopia, pcs where ucase(pcs.nompc) = " & DBSet(UCase(pc), "T")
    SQL = SQL & " and pcscopia.idaplicacion in (" & Aplicaciones & ") "
    SQL = SQL & " and pcscopia.idpcs = pcs.idpcs "
        
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    b = True
    
    While Not RS.EOF And b
        Set vAplic = New CAplicacion
        If vAplic.LeerDatos(RS!idaplicacion) Then
               b = ActualizarVersionPC(DBLet(RS!idpcs, "N"), vAplic)
               If b Then ActualizaTablaConLaVersion DBLet(RS!idpcs, "N"), vAplic
               'b = True 'para que siga
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
Dim SQL As String
Dim cad As String

    SQL = vAplic.UltVers & "."
    SQL = Replace(SQL, ".", "|")
    cad = "UPDATE pcscopia set version1 =" & RecuperaValor(SQL, 1)
    cad = cad & " ,version2 =" & RecuperaValor(SQL, 2) & ", version3 =" & RecuperaValor(SQL, 4)
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

Private Sub Form_Resize()
    On Error Resume Next
    Me.Image1.Width = Me.Width
    Me.Image1.Height = Me.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set conn = Nothing
    End
End Sub

Private Sub ListView1_DblClick()
    If ListView1.ListItems.Count = 0 Then Exit Sub
    If ListView1.SelectedItem Is Nothing Then Exit Sub
    If Trim(ListView1.SelectedItem.SubItems(3)) = "" Then Exit Sub
    AbrirNavegador ListView1.SelectedItem.SubItems(4)
End Sub



Private Sub AbrirNavegador(path As String)
    
    Screen.MousePointer = vbHourglass

    LanzaVisorMimeDocumento Me.hWnd, path
    espera 2
    Screen.MousePointer = vbDefault

End Sub

Private Sub PonPb1()
    If llevoCopiado > TamanyoTotal Then llevoCopiado = TamanyoTotal
    Me.pb1.Value = Int((llevoCopiado / TamanyoTotal) * 100)
    
End Sub

