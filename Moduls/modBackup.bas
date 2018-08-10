Attribute VB_Name = "modBackup"
Option Explicit


Public Sub BACKUP_TablaIzquierda(ByRef RS As ADODB.Recordset, ByRef cadena As String)
Dim i As Integer
Dim nexo As String

    cadena = ""
    nexo = ""
    For i = 0 To RS.Fields.Count - 1
        cadena = cadena & nexo & RS.Fields(i).Name
        nexo = ","
    Next i
    cadena = "(" & cadena & ")"
End Sub





'---------------------------------------------------
'El fichero siempre sera NF
Public Sub BACKUP_Tabla2(ByRef RS As ADODB.Recordset, ByRef Derecha As String, Optional canvi_nom As String, Optional canvi_valor As String)
Dim i As Integer
Dim nexo As String
Dim Valor As String
Dim Tipo As Integer
    Derecha = ""
    nexo = ""
    For i = 0 To RS.Fields.Count - 1
    
        If (canvi_nom <> "" And RS.Fields(i).Name = canvi_nom) Then
            Valor = canvi_valor
        Else
            Tipo = RS.Fields(i).Type
            
            If IsNull(RS.Fields(i)) Then
                Valor = "NULL"
            Else
                
                'pruebas
                Select Case Tipo
                'TEXTO
                Case 129, 200, 201
                    Valor = RS.Fields(i)
                    NombreSQL Valor    '.-----------> 23 Octubre 2003.
                    Valor = "'" & Valor & "'"
                'Fecha
                Case 133
                    Valor = CStr(RS.Fields(i))
                    Valor = "'" & Format(Valor, FormatoFecha) & "'"
                    
                'Fecha Hora
                Case 135
                    Valor = CStr(RS.Fields(i))
                    Valor = DBSet(Valor, "FH")
                
                    
                'Horas
                Case 134
                    Valor = CStr(RS.Fields(i))
                    Valor = "'" & Format(Valor, FormatoHora) & "'"
                
                'Numero normal, sin decimales
                Case 2, 3, 16 To 19, 21
                    Valor = RS.Fields(i)
                
                'Numero con decimales
                Case 6, 131
                    Valor = CStr(RS.Fields(i))
                    Valor = TransformaComasPuntos(Valor)
                Case Else
                    Valor = "Error grave. Tipo de datos no tratado." & vbCrLf
                    Valor = Valor & vbCrLf & "SQL: " & RS.Source
                    Valor = Valor & vbCrLf & "Pos: " & i
                    Valor = Valor & vbCrLf & "Campo: " & RS.Fields(i).Name
                    Valor = Valor & vbCrLf & "Valor: " & RS.Fields(i)
                    MsgBox Valor, vbExclamation
                    MsgBox "El programa finalizara. Avise al soporte técnico.", vbCritical
                    End
                End Select
            End If
        End If
        Derecha = Derecha & nexo & Valor
        nexo = ","
    Next i
    Derecha = "(" & Derecha & ")"
End Sub


'Para los nombre que pueden tener ' . Para las comillas habra que hacer dentro otro INSTR
Public Sub QUITASALTOSLINEA(ByRef cadena As String)
Dim J As Long
Dim i As Long
Dim Aux As String
    J = 1
    Do
        i = InStr(J, cadena, vbCrLf)
        If i > 0 Then
            Aux = Mid(cadena, 1, i - 1) & "\n\r"
            cadena = Aux & Mid(cadena, i + 2)
            J = i + 2
        End If
    Loop Until i = 0
End Sub


