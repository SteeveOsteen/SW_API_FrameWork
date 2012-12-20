Attribute VB_Name = "Macro_Outils"
Public Sub RecalculerLesEquation()
    
    Dim Sw                  As New ExtSldWorks
    Dim ModeleDeBase        As ExtModele
    Dim Modele              As ExtModele
    Dim Composant           As ExtComposant
    Dim Rechercher          As ExtRechercher
    
    Set ModeleDeBase = Sw.Modele
    Set Rechercher = ModeleDeBase.NouvelleRecherche
    Rechercher.PrendreEnCompteConfig = False
    Rechercher.PrendreEnCompteExclus = True
    
    For Each Composant In Rechercher.ListeDesComposants(cPiece)
        Set Modele = Composant.Modele
        Debug.Print Modele.Fichier.NomDuFichier
        
        Modele.Activer
        
        Modele.Piece.GestDeEquations.ToutRecalculer
        
        'Si le modele est identique au modele de base, il ne faut pas le fermer
        If Not (Modele.Fichier.Chemin = ModeleDeBase.Fichier.Chemin) Then
            Modele.Fermer
        End If
        
    Next Composant
    
    Set Rechercher = Nothing
    Set Composant = Nothing
    Set ModeleDeBase = Nothing
    Set Sw = Nothing
    
End Sub

Public Sub InsererLesEpaisseursDeTolerie()
    
    Dim Sw                  As New ExtSldWorks
    Dim ModeleDeBase        As ExtModele
    Dim Modele              As ExtModele
    Dim ComposantDeBase     As ExtComposant
    Dim Composant           As ExtComposant
    Dim Configuration       As ExtConfiguration
    Dim Rechercher          As ExtRechercher
    Dim RechercherConfig    As ExtRechercher
    Dim Dossier             As ExtDossier
    Dim Corps               As ExtCorps
    Dim Fonction            As ExtFonction
    Dim Tole                As SheetMetalFeatureData
    
    Debug.Print " "
    Debug.Print "------------------------------------"
    
    Set ModeleDeBase = Sw.Modele
    Set Rechercher = ModeleDeBase.NouvelleRecherche
    Rechercher.PrendreEnCompteConfig = False
    Rechercher.PrendreEnCompteExclus = True
    
    Set RechercherConfig = ModeleDeBase.NouvelleRecherche
    RechercherConfig.PrendreEnCompteConfig = True
    RechercherConfig.PrendreEnCompteExclus = True
    
    For Each ComposantDeBase In Rechercher.ListeDesComposants(cPiece)
        Set Modele = ComposantDeBase.Modele
        Modele.Activer
        Debug.Print Modele.Fichier.NomDuFichier
        
        For Each Composant In RechercherConfig.ListeDesComposants(cPiece, Modele.Fichier.NomDuFichier)
            Composant.Configuration.Activer
            Debug.Print , "Configuration : "; Composant.Configuration.Nom
            For Each Dossier In Composant.Modele.Piece.ListeDesDossiers(cTole)
                Set Corps = Dossier.PremierCorps
                For Each Fonction In Corps.ListeDesFonctions
                    If Fonction.TypeDeLaFonction = "SheetMetal" Then
                        Set Tole = Fonction.swFonction.GetDefinition
                        Debug.Print , , "Corps : "; Corps.Nom; " -> Epaisseur de la tôle : ", Tole.Thickness * 1000
                        Dossier.GestDeProprietes.AjouterPropriete "Epaisseur de la tôle", swCustomInfoText, CStr(Tole.Thickness * 1000)
                        Exit For
                    End If
                Next Fonction
            Next Dossier
        Next Composant
        
        Modele.Sauver
        
        'Si le modele est identique au modele de base, il ne faut pas le fermer
        If Not (Modele.Fichier.Chemin = ModeleDeBase.Fichier.Chemin) Then
            Modele.Fermer
        End If
        
    Next ComposantDeBase
    
    ModeleDeBase.Sauver
    
    Set Sw = Nothing
    Set ModeleDeBase = Nothing
    Set Modele = Nothing
    Set ComposantDeBase = Nothing
    Set Composant = Nothing
    Set Configuration = Nothing
    Set Rechercher = Nothing
    Set RechercherConfig = Nothing
    Set Dossier = Nothing
    Set Corps = Nothing
    Set Fonction = Nothing
    Set Tole = Nothing
    
End Sub

