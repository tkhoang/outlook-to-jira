Attribute VB_Name = "MailHelper"
Function ObtenirTexteSelectionne() As String
    Dim objApp As Outlook.Application
    Dim objExplorer As Outlook.Explorer
    Dim objSelection As Outlook.Selection
    Dim objMail As Outlook.MailItem
    Dim texteSelectionne As String
    
    ' Obtenir l'application Outlook en cours d'ex�cution
    Set objApp = GetObject(, "Outlook.Application")
    
    ' Obtenir l'explorateur actif dans Outlook
    Set objExplorer = objApp.ActiveExplorer
    
    ' Obtenir la s�lection actuelle dans l'explorateur
    Set objSelection = objExplorer.Selection
    
    ' V�rifier si une seule s�lection a �t� effectu�e
    If objSelection.Count = 1 Then
        ' V�rifier si l'�l�ment s�lectionn� est un e-mail
        If TypeOf objSelection.Item(1) Is Outlook.MailItem Then
            ' Convertir l'�l�ment s�lectionn� en objet MailItem
            Set objMail = objSelection.Item(1)
            
                ' R�cup�rer le texte s�lectionn�
                texteSelectionne = objMail.GetInspector.WordEditor.Application.Selection.text
                
                ' Retourner le texte s�lectionn�
                ObtenirTexteSelectionne = texteSelectionne
            
        Else
            ' L'�l�ment s�lectionn� n'est pas un e-mail
            ObtenirTexteSelectionne = "Veuillez s�lectionner un e-mail."
        End If
    Else
        ' Aucune ou plusieurs s�lections effectu�es
        ObtenirTexteSelectionne = "Veuillez s�lectionner un seul e-mail."
    End If
    
    ' Lib�rer les objets m�moire
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
    
    ' Obtenir l'application Outlook en cours d'ex�cution
    Set objApp = GetObject(, "Outlook.Application")
    
    ' V�rifier si un inspecteur est actif
    If Not objApp.ActiveInspector Is Nothing Then
        ' Obtenir l'inspecteur actif
        Set objInspector = objApp.ActiveInspector
        
        ' V�rifier si l'inspecteur est en train de montrer un mail
        If TypeOf objInspector.CurrentItem Is Outlook.MailItem Then
            ' Convertir l'�l�ment actuel en objet MailItem
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
    
    ' Lib�rer les objets m�moire
    Set objMail = Nothing
    Set objInspector = Nothing
    Set objApp = Nothing
End Function

