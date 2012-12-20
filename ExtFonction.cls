VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExtFonction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private pSwFonction             As Feature
Private pNomFonction            As String
Private pPiece                  As ExtPiece
Private pEnregistrerEtat        As EtatFonction_e
Private Const pNomClasse        As String = "ExtFonction"
Private Erreur As Long, Warning As Long

Private Sub Class_Terminate()
    Set pSwFonction = Nothing
    Set pPiece = Nothing
    
End Sub

Public Property Get Nom() As String
    Nom = swFonction.Name
End Property

Public Property Let Nom(s As String)
    If pNomFonction = s Then Exit Property
    
    Dim NomTmp As String
    Dim i As Integer
    i = 1
    NomTmp = s
    While pPiece.Modele.swModele.FeatureManager.IsNameUsed(swFeatureName, NomTmp)
        NomTmp = s & " " & i
        i = i + 1
    Wend
    swFonction.Name = NomTmp
    pNomFonction = swFonction.Name
End Property

Public Sub EnregistrerEtat()
    pEnregistrerEtat = Etat
End Sub

Public Sub RestaurerEtat()
    If pEnregistrerEtat = cActivee Then
        Activer
    Else
        Desactiver
    End If
End Sub

Public Function SetFonction(ByRef Fonction As Feature, Piece As ExtPiece) As Boolean
    
    If Not (Fonction Is Nothing) And Not (Piece Is Nothing) Then
        SetFonction = True
        Set pSwFonction = Fonction
        pNomFonction = Fonction.Name
        Set pPiece = Piece
        Exit Function
    End If
    
    Debug.Print pNomClasse & ".SetFonction : le Modele de base n'est pas valide"
    
End Function

Public Function TypeDeLaFonction() As String
    TypeDeLaFonction = swFonction.GetTypeName2
End Function

Public Function Etat() As EtatFonction_e
    Dim NomConfig   As String
    Dim vListeEtat  As Variant
    NomConfig = pPiece.Modele.GestDeConfigurations.ConfigurationActive.Nom
    vListeEtat = swFonction.IsSuppressed2(swInConfigurationOpts_e.swThisConfiguration, Array(NomConfig))
    If vListeEtat(0) = 0 Then
        Etat = cActivee
    Else
        Etat = cDesactivee
    End If
    
End Function
Public Sub Activer(Optional ActiverLesSousFonctions As Boolean = True)
    pActiver
    
    If ActiverLesSousFonctions Then
        Dim SousFonction As ExtFonction
        For Each SousFonction In ListeDesSousFonctions
            SousFonction.Activer False
        Next SousFonction
        Set SousFonction = Nothing
    End If
    
End Sub

Public Sub Desactiver(Optional DesactiverLesSousFonctions As Boolean = False)
    pDesactiver
    
    If DesactiverLesSousFonctions Then
        Dim SousFonction As ExtFonction
        For Each SousFonction In ListeDesSousFonctions
            SousFonction.Desactiver False
        Next SousFonction
        Set SousFonction = Nothing
    End If
    
End Sub


Private Sub pActiver()
    
    Dim ModeleDoc           As ModelDoc2
    Dim NomSelectFonction   As String
    Dim TypeFonction        As String
    Set ModeleDoc = pPiece.Modele.swModele
    NomSelectFonction = swFonction.GetNameForSelection(TypeFonction)
    ModeleDoc.Extension.SelectByID2 NomSelectFonction, TypeFonction, 0, 0, 0, False, -1, Nothing, 0
    ModeleDoc.EditUnsuppress2
    ModeleDoc.EditUnsuppressDependent2
    Set ModeleDoc = Nothing

End Sub

Private Sub pDesactiver()
    Dim ModeleDoc           As ModelDoc2
    Dim NomSelectFonction   As String
    Dim TypeFonction        As String
    Set ModeleDoc = pPiece.Modele.swModele
    NomSelectFonction = swFonction.GetNameForSelection(TypeFonction)
    ModeleDoc.Extension.SelectByID2 NomSelectFonction, TypeFonction, 0, 0, 0, False, -1, Nothing, 0
    ModeleDoc.EditSuppress2
    Set ModeleDoc = Nothing
    
End Sub

Public Function ListeDesSousFonctions(Optional ByVal NomARechercher As String) As Collection
    
    Dim SousFonction        As Feature
    Dim FonctionExt         As ExtFonction
    
    If NomARechercher = vbNullString Then NomARechercher = "*"
    
    Set ListeDesSousFonctions = New Collection
    
    Set SousFonction = swFonction.GetFirstSubFeature
    
    Do Until SousFonction Is Nothing
        If SousFonction.Name Like NomARechercher Then
            Set FonctionExt = New ExtFonction
            FonctionExt.SetFonction SousFonction, pPiece
            ListeDesSousFonctions.Add FonctionExt
        End If
        Set SousFonction = SousFonction.GetNextSubFeature
    Loop
    
    Set FonctionExt = Nothing
    Set SousFonction = Nothing
    
End Function

Public Function swFonction() As Feature
    Dim Fonction        As Feature
    Dim SousFonction    As Feature
    Dim Sortir          As Boolean
    
    Set swFonction = pSwFonction
    pNomFonction = pSwFonction.Name
    
End Function
