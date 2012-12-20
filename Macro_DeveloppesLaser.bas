Attribute VB_Name = "Macro_DeveloppesLaser"
Public Sub CreerConfigsPourDvpLaser()
    Dim Sw                  As New ExtSldWorks
    Dim ModeleDeBase        As ExtModele
    Dim Composant           As ExtComposant
    Dim Modele              As ExtModele
    Dim ConfigDeDepart      As ExtConfiguration
    Dim ConfigPliee         As ExtConfiguration
    Dim Dossier             As ExtDossier
    Dim Corps               As ExtCorps
    Dim Fonction            As ExtFonction
    Dim Rechercher          As ExtRechercher
    
    Set ModeleDeBase = Sw.Modele
    
    Set Rechercher = ModeleDeBase.NouvelleRecherche
    Rechercher.PrendreEnCompteConfig = False
    Rechercher.PrendreEnCompteExclus = False
    
    For Each Composant In Rechercher.ListeDesComposants(cPiece)
            Set Modele = Composant.Modele
            Modele.Activer
            
            Debug.Print Modele.Fichier.NomDuFichier
            Debug.Print , "Contient des toles : "; Modele.Piece.Contient(cTole)
            
            If Modele.Piece.Contient(cTole) Then
                
                Set ConfigDeDepart = Modele.GestDeConfigurations.ConfigurationActive
                If ConfigDeDepart.Est(cDepliee) Then
                    Set ConfigDeDepart = ConfigDeDepart.ConfigurationParent
                    ConfigDeDepart.Activer
                End If
                
                Modele.GestDeConfigurations.SupprimerLesConfigurationsDeplies
                Modele.Sauver
                
                For Each ConfigPliee In Modele.GestDeConfigurations.ListerLesConfigs(cPliee)
                    
                    ConfigPliee.Activer
                    Modele.Sauver
                    Modele.Piece.GestDeMiseAJour.MettreAJourLaListeDesPiecesSoudees
                    
                    For Each Dossier In Modele.Piece.ListeDesDossiers(cTole, False)
                        Set Corps = Dossier.PremierCorps
                        Debug.Print , "Dossier : "; Dossier.Nom
                        Corps.Tolerie.CreerConfigurationDepliee
                    Next Dossier
                
                Next ConfigPliee
                
                ConfigDeDepart.Activer
                
                Modele.Piece.GestDeMiseAJour.MettreAJourLaListeDesPiecesSoudees
                Modele.Piece.GestDeMiseAJour.MettreAJourLesNomsDeConfigs
                
                Modele.Sauver
            End If
            
            'Si le modele est identique au modele de base, il ne faut pas le fermer
            If Not (Modele.Fichier.Chemin = ModeleDeBase.Fichier.Chemin) Then
                Modele.Fermer
            End If
        Next Composant
        
    ModeleDeBase.Sauver
    
    Set Rechercher = Nothing
    Set Dossier = Nothing
    Set ConfigPliee = Nothing
    Set ConfigDeDepart = Nothing
    Set Modele = Nothing
    Set Composant = Nothing
    Set ModeleDeBase = Nothing
    Set Sw = Nothing
End Sub

Public Sub CreerVuePourDvpLaser()
    
    Dim Sw                  As New ExtSldWorks
    Dim GestDeFichier       As New SysGestDeFichiers
    Dim DossierLaser        As String
    Dim ModeleDeBase        As ExtModele
    Dim NomAssRacine        As String
    Dim NomPiece            As String
    Dim ComposantDeBase     As ExtComposant
    Dim Composant           As ExtComposant
    Dim Modele              As ExtModele
    Dim ModeleDessin        As ExtModele
    Dim Dessin              As ExtDessin
    Dim Feuille             As ExtFeuille
    Dim Rechercher          As ExtRechercher
    Dim RechercherComp      As ExtRechercher
    
    Set ModeleDeBase = Sw.Modele
    
    'Si c'est une piece, on filtre la liste des pièces avec le nom de la piece
    'et on refait seulement les dvps de la pièce active
    If ModeleDeBase.TypeDuModele = cPiece Then
        NomPiece = ModeleDeBase.Fichier.NomDuFichier
    End If
    
    'Si le modèle contient la propriété "AssemblageRacine", on ouvre le fichier et on le defini comme assemblage
    'de base pour le decompte des pièces
    NomAssRacine = ModeleDeBase.GestDeProprietes.RecupererPropriete("AssemblageRacine")
    If Not (NomAssRacine = vbNullString) Then
        Set ModeleDeBase = Sw.Modele(ModeleDeBase.Fichier.NomDuDossier & "\" & NomAssRacine & ".SLDASM")
        ModeleDeBase.Activer
    End If
    
    'On creer un dossier à la racine de l'assemblage pour y placer les dessins des dvps
    GestDeFichier.Chemin = ModeleDeBase.Fichier.Chemin
    DossierLaser = GestDeFichier.CreerDossier("Plans Laser")
    
    Set Rechercher = ModeleDeBase.NouvelleRecherche
    Rechercher.PrendreEnCompteConfig = False
    Rechercher.PrendreEnCompteExclus = False
    
    For Each ComposantDeBase In Rechercher.ListeDesComposants(cPiece, NomPiece)
        Debug.Print ComposantDeBase.Modele.Fichier.NomDuFichier
        If ComposantDeBase.Modele.Piece.Contient(cTole) Then
            ComposantDeBase.Modele.Activer
            
            Set ModeleDessin = Sw.CreerDessin(DossierLaser, ComposantDeBase.Modele.Fichier.NomDuFichier(True) & " (" & ComposantDeBase.Modele.GestDeProprietes.RecupererPropriete("Designation") & ")", MODELE_DE_DESSIN_LASER)
            Set Dessin = ModeleDessin.Dessin
            
            Debug.Print , "Dessin : "; ModeleDessin.Fichier.NomDuFichier
            
            Set RechercherComp = ModeleDeBase.NouvelleRecherche
            RechercherComp.PrendreEnCompteConfig = True
            RechercherComp.PrendreEnCompteExclus = False
            
            For Each Composant In RechercherComp.ListeDesComposants(cPiece, ComposantDeBase.Modele.Fichier.NomDuFichier)
                Set Modele = Composant.Modele
                Composant.Configuration.Activer
                
                If Modele.Piece.Contient(cTole) Then
                    Debug.Print , "Contient des toles", Modele.Piece.Contient(cTole)
                    ModeleDessin.Activer
                    Modele.Piece.CreerVuePourDvpLaser Dessin
                    
                End If
            Next Composant
            
            Set Feuille = Dessin.Feuille("Feuille1")
            Feuille.Supprimer
            
            ModeleDessin.Sauver
            ModeleDessin.Fermer
            'Si le modele est identique au modele de base, il ne faut pas le fermer
            If Not (ComposantDeBase.Modele.Fichier.Chemin = ModeleDeBase.Fichier.Chemin) Then
                ComposantDeBase.Modele.Fermer
            End If
        End If
    Next ComposantDeBase
    
    Set Rechercher = Nothing
    Set RechercherComp = Nothing
    Set Feuille = Nothing
    Set Dessin = Nothing
    Set ModeleDessin = Nothing
    Set Modele = Nothing
    Set Composant = Nothing
    Set ComposantDeBase = Nothing
    Set ModeleDeBase = Nothing
    Set GestDeFichier = Nothing
    Set Sw = Nothing
    
End Sub
