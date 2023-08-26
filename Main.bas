Attribute VB_Name = "Main"


Sub OuvrirUserForm()

    IsTestMode = True
    
    SetupEnvironment

    ' selection des infos du mail
    selectionText = ObtenirTexteSelectionne()
    objetMail = ObtenirObjetMail()
        
    Dim uf As CreateJira
    Set uf = New CreateJira
    
    uf.setSummary = objetMail
    uf.setDescription = selectionText
    uf.setSentToTest = IsTestMode
    
    ' Affiche la UserForm avec les infos du mails
    uf.Show
    
End Sub

