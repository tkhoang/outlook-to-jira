VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CreateJira 
   Caption         =   "Jira Form"
   ClientHeight    =   7410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6945
   OleObjectBlob   =   "CreateJira.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CreateJira"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' parametre SendToTest récupérer depuis un mail
Private paramSendToTest As String
' variable pour si envoie en test ou en prod accessible dans tous le formulaire avec le "let" et "get" récupérer depuis un mail
Public m_sentToTest As String
' parametre summary récupérer depuis un mail
Private paramSummary As String
' parametre description récupérer depuis un mail
Private paramDescription As String
' variable pour selection le type d'issue accessible dans tous le formulaire avec le "let" et "get" récupérer depuis un mail
Public m_selectedIssueType As String

Public Property Let setSentToTest(ByVal value As String)
    paramSendToTest = value
End Property

Public Property Let setSummary(ByVal value As String)
    paramSummary = value
End Property

Public Property Let setDescription(ByVal value As String)
    paramDescription = value
End Property

Public Property Get SelectedIssueType() As String
    
    SelectedIssueType = m_selectedIssueType
End Property

Public Property Let SelectedIssueType(ByVal value As String)
    m_selectedIssueType = value
End Property

Private Sub CommandButton1_Click()

End Sub

Private Sub Description_Change()

End Sub



Private Sub envOverride_Click()

End Sub

Private Sub sentToTest_Click()
  
  
  IsTestMode = Me.sentToTest.value
  SetupEnvironment
  
End Sub

Private Sub Submit_Click()
    
    ' appel l'api Jira pour creer un ticket
    response = RequeteHTTPAvecAuthentification(CreerJSON(Me.Summary.value, Me.Description.value, CreateJira.SelectedIssueType))
    
    If InStr(response, "key") > 0 Then
        ' si tout s'est bien passé, ouverture du ticket dans un navigateur
        Set Parsed = JsonConverter.ParseJson(response)
        Key = Parsed("key")
        OuvrirLienAvecNavigateur (Key)
        Unload Me
    Else
    MsgBox (response)
    End If
    

    
End Sub

Private Sub Summary_Change()

End Sub

Private Sub Title_Change()

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()

    
    Me.sentToTest.value = False
    Me.Summary.value = ""
    Me.Description.value = ""
    
    Set IssueTypeCollection = New Collection
    
    ' Ajout des types de tickets à la collection itemIssueType
    IssueTypeCollection.Add Array("3", "Task")
    IssueTypeCollection.Add Array("1", "Bug")
    
    ' Remplir la liste déroulante itemIssueType
    For Each Item In IssueTypeCollection
        IssueType.AddItem Item(1)
    Next Item


End Sub

Private Sub UserForm_Activate()

    Me.Summary.value = paramSummary
    Me.Description.value = paramDescription
    Me.sentToTest.value = paramSendToTest

End Sub

Private Sub IssueType_Change()

    Set IssueTypeCollection = New Collection
    
     ' Ajout des types de tickets à la collection itemIssueType
     ' TODO : initialiser la collection une seule fois pour l'init et le listener
    IssueTypeCollection.Add Array("3", "Task")
    IssueTypeCollection.Add Array("1", "Bug")

    Dim selectedIndex As Long
    selectedIndex = IssueType.ListIndex
    
    ' selection du type de ticket
    If selectedIndex >= 0 Then
        CreateJira.SelectedIssueType = IssueTypeCollection(selectedIndex + 1)(0)
        Debug.Print selectedId
    End If

End Sub
