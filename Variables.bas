Attribute VB_Name = "Variables"
Public API_URL As String
Public API_BEARER As String
Public IsTestMode As Boolean

Sub SetupEnvironment()
    If IsTestMode Then
        Debug.Print ("Test Mode")
        API_URL = "" 'put your test base url here if you have any, whithout the resource
        API_BEARER = "" 'put your bearer key here
        DEFAULT_EPIC = "" 'put a default epic here if you wish
    Else
        Debug.Print ("Prod Mode")
        API_URL = "" 'put your test base url here if you have any, whithout the resource
        API_BEARER = "" 'put your bearer key here
        DEFAULT_EPIC = "" 'put a default epic here if you wish
    End If
End Sub




