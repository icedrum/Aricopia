Attribute VB_Name = "linNavegador"

'------------------------------------------------------------------------
'------------------------------------------------------------------------
' Lanza visores predeterminados por MIME
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long






Public Function LanzaVisorMimeDocumento(Formhwnd As Long, Archivo As String)
    Call ShellExecute(Formhwnd, "open", Archivo, "", "", 1)
End Function

