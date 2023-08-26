Attribute VB_Name = "MailHelper"
Function ObtenirTexteSelectionne() As String
    Dim objApp As Outlook.Application
    Dim objExplorer As Outlook.Explorer
    Dim objSelection As Outlook.Selection
    Dim objMail As Outlook.MailItem
    Dim texteSelectionne As String
    
    ' Obtenir l'application Outlook en cours d'exécution
    Set objApp = GetObject(, "Outlook.Application")
    
    ' Obtenir l'explorateur actif dans Outlook
    Set objExplorer = objApp.ActiveExplorer
    
    ' Obtenir la sélection actuelle dans l'explorateur
    Set objSelection = objExplorer.Selection
    
    ' Vérifier si une seule sélection a été effectuée
    If objSelection.Count = 1 Then
        ' Vérifier si l'élément sélectionné est un e-mail
        If TypeOf objSelection.Item(1) Is Outlook.MailItem Then
            ' Convertir l'élément sélectionné en objet MailItem
            Set objMail = objSelection.Item(1)
            
                ' Récupérer le texte sélectionné
                texteSelectionne = objMail.GetInspector.WordEditor.Application.Selection.text
                
                ' Retourner le texte sélectionné
                ObtenirTexteSelectionne = texteSelectionne
            
        Else
            ' L'élément sélectionné n'est pas un e-mail
            ObtenirTexteSelectionne = "Veuillez sélectionner un e-mail."
        End If
    Else
        ' Aucune ou plusieurs sélections effectuées
        ObtenirTexteSelectionne = "Veuillez sélectionner un seul e-mail."
    End If
    
    ' Libérer les objets mémoire
    Set objMail = Nothing
    Set objSelection = Nothing
    Set objExplorer = Nothing
    Set objApp = Nothing
End Function

Function ObtenirObjetMail() As String
    Dim objApp As Outlook.Application
    Dim objInspector As Outlook.Inspector
    Dim objMail As Outlook.MailItem
    Dim sujetMail As String
    
    ' Obtenir l'application Outlook en cours d'exécution
    Set objApp = GetObject(, "Outlook.Application")
    
    ' Vérifier si un inspecteur est actif
    If Not objApp.ActiveInspector Is Nothing Then
        ' Obtenir l'inspecteur actif
        Set objInspector = objApp.ActiveInspector
        
        ' Vérifier si l'inspecteur est en train de montrer un mail
        If TypeOf objInspector.CurrentItem Is Outlook.MailItem Then
            ' Convertir l'élément actuel en objet MailItem
            Set objMail = objInspector.CurrentItem
            
            ' Obtenir le sujet du mail ouvert
            sujetMail = objMail.Subject
            
            ObtenirObjetMail = sujetMail
        Else
            ' L'inspecteur n'affiche pas un mail
            MsgBox "Aucun mail ouvert."
        End If
    Else
        ' Aucun inspecteur actif
        MsgBox "Aucun mail ouvert."
    End If
    
    ' Libérer les objets mémoire
    Set objMail = Nothing
    Set objInspector = Nothing
    Set objApp = Nothing
End Function

