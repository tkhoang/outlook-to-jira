Attribute VB_Name = "SystemHelper"
Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hwnd As LongPtr, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long _
) As LongPtr

Public Function OuvrirLienAvecNavigateur(Key As String)
    Dim url As String
    url = API_URL + "/browse/"
    
    Dim ret As LongPtr
    ret = ShellExecute(0, "open", url + Key, vbNullString, vbNullString, vbNormalFocus)
    
    If ret <= 32 Then
        MsgBox "Erreur lors de l'ouverture du lien"
    End If
End Function

