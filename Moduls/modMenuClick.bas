Attribute VB_Name = "modMenuClick"
Option Explicit

Private Sub Construc(nom As String)
    MsgBox nom & ": en construcció..."
End Sub

' ******* DATOS BASICOS *********

Public Sub SubmnP_Generales_Click(Index As Integer)

    Select Case Index
        Case 1: frmManAplic.Show vbModal
        Case 2: frmManPCs.Show vbModal
        Case 3: frmManFicCopia.Show vbModal
        Case 4: frmManPcsCopia.Show vbModal
        Case 6: End
    End Select
End Sub


' *******  UTILIDADES *********
Public Sub SubmnE_Util_Click(Index As Integer)
'    Select Case Index
'        Case 1: frmCaracteresMB.Show vbModal ' comprobacion de caracteres de multibase
'        Case 3: frmBackUP.Show vbModal
'    End Select
End Sub

