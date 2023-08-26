Attribute VB_Name = "restHelper"
Public Function RequeteHTTPAvecAuthentification(body As String) As String
    Dim url As String
    Dim xmlhttp As Object
    Dim response As String
    Dim username As String
    Dim password As String
    Dim basicAuth As String
    
    ' URL de la ressource à récupérer
    url = API_URL + "/rest/api/latest/issue"
    
    ' Création de l'objet XMLHTTP
    Dim http: Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    
    ' Préparation de la requête GET
    http.Open "POST", url
    
    
    ' Ajout de l'en-tête d'autorisation et de content
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Authorization", "Bearer " + API_BEARER
    
    ' Envoi de la requête
    http.Send body
    
    ' Récupération de la réponse
    RequeteHTTPAvecAuthentification = http.responseText
    
    ' Libération de l'objet XMLHTTP
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
    'component("id") = ""
    
    components.Add component
    
    ' creation de l'obj "fields"
    Dim fields As Object
    Set fields = CreateObject("Scripting.Dictionary")
    fields("summary") = Summary
    Set fields("issuetype") = IssueType
    Set fields("project") = project
    'Set fields("components") = components
    fields("description") = Description
    
    
    ' creation de l'objet racine
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    
    ' Ajouter des clés et des valeurs à l'objet JSON
    root("title") = Title
    Set root("fields") = fields
    
    ' Convertir l'objet JSON en chaîne JSON
    Dim jsonText As String
    jsonText = JsonConverter.ConvertToJson(root, Whitespace:=3)
    
    ' Afficher la chaîne JSON
    ' Debug.Print jsonText
    CreerJSON = jsonText
End Function

