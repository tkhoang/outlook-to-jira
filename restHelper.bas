Attribute VB_Name = "restHelper"
Public Function RequeteHTTPAvecAuthentification(body As String) As String
    Dim url As String
    Dim xmlhttp As Object
    Dim response As String
    Dim username As String
    Dim password As String
    Dim basicAuth As String
    
    ' URL de la ressource � r�cup�rer
    url = API_URL + "/rest/api/latest/issue"
    
    ' Cr�ation de l'objet XMLHTTP
    Dim http: Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    
    ' Pr�paration de la requ�te GET
    http.Open "POST", url
    
    
    ' Ajout de l'en-t�te d'autorisation et de content
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Authorization", "Bearer " + API_BEARER
    
    ' Envoi de la requ�te
    http.Send body
    
    ' R�cup�ration de la r�ponse
    RequeteHTTPAvecAuthentification = http.responseText
    
    ' Lib�ration de l'objet XMLHTTP
    Set xmlhttp = Nothing
End Function

Public Function CreerJSON(Summary As String, Description As String, pIssueType As String) As String
    ' creation de l'obj "issuetype"
    Dim IssueType As Object
    Set IssueType = CreateObject("Scripting.Dictionary")
    IssueType("id") = pIssueType
    
    ' creation de l'obj "project"
    Dim project As Object
    Set project = CreateObject("Scripting.Dictionary")
    project("key") = "PSEP"
    
    ' creation d'un tableau pour les composants
    Dim components As Collection
    Set components = New Collection
    
    'composant 1
    Dim component As Object
    Set component = CreateObject("Scripting.Dictionary")
    component("id") = "11228"
    
    components.Add component
    
    ' creation de l'obj "fields"
    Dim fields As Object
    Set fields = CreateObject("Scripting.Dictionary")
    fields("summary") = Summary
    Set fields("issuetype") = IssueType
    Set fields("project") = project
    Set fields("components") = components
    fields("description") = Description
    fields("customfield_10006") = "PSEP-117044"
    
    
    ' creation de l'objet racine
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    
    ' Ajouter des cl�s et des valeurs � l'objet JSON
    root("title") = Title
    Set root("fields") = fields
    
    ' Convertir l'objet JSON en cha�ne JSON
    Dim jsonText As String
    jsonText = JsonConverter.ConvertToJson(root, Whitespace:=3)
    
    ' Afficher la cha�ne JSON
    ' Debug.Print jsonText
    CreerJSON = jsonText
End Function

