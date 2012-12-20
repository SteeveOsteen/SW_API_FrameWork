VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExtPiece"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private pModele                 As ExtModele
Private pGestDeMiseAJour        As New GestDeMiseAJourPiece
Private pGestDeEquations        As New ExtGestDeEquations
Private Const pNomClasse        As String = "ExtPiece"
Private Erreur As Long, Warning As Long

Private Sub Class_Terminate()
    
    Set pModele = Nothing
    Set pGestDeMiseAJour = Nothing
    Set pGestDeEquations = Nothing
End Sub

Public Function Modele() As ExtModele
    Set Modele = pModele
End Function

Public Function SetModele(Modele As ExtModele) As Boolean
    
    If Not (Modele Is Nothing) Then
        If Modele.Est(cPiece) Then
            SetModele = True
            Set pModele = Modele
            pGestDeMiseAJour.SetModele Me
            pGestDeEquations.SetGestionnaire pModele.swModele.GetEquationMgr, pModele
            Exit Function
        End If
    End If
    
    Debug.Print pNomClasse & ".SetModele : le Modele de base n'est pas valide"
    
End Function

Public Function GestDeMiseAJour() As GestDeMiseAJourPiece
    Set GestDeMiseAJour = pGestDeMiseAJour
End Function

Public Function GestDeEquations() As ExtGestDeEquations
    Set GestDeEquations = pGestDeEquations
End Function

'===============================================================================
'
'                                   Fonction
'
'===============================================================================

Public Function ListeDesDossiers(Optional TypeDeCorps As TypeCorps_e = cTousLesTypesDeCorps, Optional PrendreEnCompteExclus As Boolean = False) As Collection
    Dim Dossier     As ExtDossier
    Dim Fonction    As Feature
    Dim DossierSw   As BodyFolder
    
    Set ListeDesDossiers = New Collection
    Set Fonction = ListeDesPiecesSoudees.GetFirstSubFeature
    
    Do Until Fonction Is Nothing
    
        'Si c'est un dossier de pièces soudées
        If Fonction.GetTypeName2 = "CutListFolder" Then
            
            Set DossierSw = Fonction.GetSpecificFeature2
            Set Dossier = New ExtDossier
            
            If Dossier.SetDossier(DossierSw, Me) Then
                If Dossier.Est(TypeDeCorps) And (Dossier.Exclu Imp PrendreEnCompteExclus) Then
                    ListeDesDossiers.Add Dossier
                End If
            End If
            Set Dossier = Nothing
        End If
        Set Fonction = Fonction.GetNextSubFeature
    Loop
    
    Set Dossier = Nothing
    Set Fonction = Nothing
    Set DossierSw = Nothing
    
End Function

Private Function ListeDesPiecesSoudees() As Feature
    Dim PiecesSoudees   As Feature
    Dim DossierSw       As BodyFolder
    
    'Recherche la liste des corps
    Set PiecesSoudees = pModele.swModele.FirstFeature
    
    Do Until PiecesSoudees Is Nothing
    
        If PiecesSoudees.GetTypeName2 = "SolidBodyFolder" Then
        
            Set ListeDesPiecesSoudees = PiecesSoudees
            Set DossierSw = PiecesSoudees.GetSpecificFeature2
            DossierSw.UpdateCutList
            Exit Do
        End If
        Set PiecesSoudees = PiecesSoudees.GetNextFeature
    Loop
    
    Set PiecesSoudees = Nothing
    Set DossierSw = Nothing

End Function

Public Function ListeDesFonctions(Optional ByVal NomARechercher As String) As Collection
    Dim Fonction            As Feature
    Dim FonctionExt         As ExtFonction
    
    If NomARechercher = vbNullString Then NomARechercher = "*"
    
    Set ListeDesFonctions = New Collection
    
    Set Fonction = pModele.swModele.FirstFeature
    
    Do Until Fonction Is Nothing
    
        If Fonction.Name Like NomARechercher Then
            Set FonctionExt = New ExtFonction
            FonctionExt.SetFonction Fonction, pModele.Piece
            ListeDesFonctions.Add FonctionExt
        End If
        
        Set Fonction = Fonction.GetNextFeature
    Loop
    
    Set FonctionExt = Nothing
    Set Fonction = Nothing

End Function

Public Function Contient(ByVal TypeDeCorps As TypeCorps_e) As Boolean
    Dim Fonction As ExtFonction
    
    If (TypeDeCorps And cTole) Then
        For Each Fonction In ListeDesFonctions
            If (Fonction.TypeDeLaFonction = "SMBaseFlange") Or (Fonction.TypeDeLaFonction = "SolidToSheetMetal") Then
                Contient = True
                Exit For
            End If
        Next Fonction
       TypeDeCorps = TypeDeCorps - cTole
    End If
    
    If (TypeDeCorps And cProfil) Then
        For Each Fonction In ListeDesFonctions
        Debug.Print Fonction.Nom, Fonction.TypeDeLaFonction
            If Fonction.TypeDeLaFonction = "WeldMemberFeat" Then
                Contient = True
                Exit For
            End If
        Next Fonction
        TypeDeCorps = TypeDeCorps - cProfil
    End If
    
    If (TypeDeCorps And cAutre) Then
        If Me.ListeDesDossiers(cAutre, False).Count Then
                Contient = Contient Or True
        End If
    End If
    
    Set Fonction = Nothing
    
End Function

Public Sub CreerVuePourDvpLaser(ByRef Dessin As ExtDessin)
    
    Dim NomFeuille              As String
    Dim Feuille                 As ExtFeuille
    Dim VueDepliee              As ExtVue
    Dim ConfigDeDepart          As ExtConfiguration
    Dim ConfigDepliee           As ExtConfiguration
    Dim ListeDesConfigDepliee   As Collection
    
    Dim Centre              As Point
    Dim Enveloppe           As Rectangle
    
    Set ConfigDeDepart = pModele.GestDeConfigurations.ConfigurationActive
    
    Set ListeDesConfigDepliee = pModele.GestDeConfigurations.ListerLesConfigs(cDepliee, ConfigDeDepart.Nom)
    
    If ConfigDeDepart.Est(cPliee) And ListeDesConfigDepliee.Count Then
        
        NomFeuille = pModele.Fichier.NomDuFichier(True) & "-" & _
                            ConfigDeDepart.Nom & _
                            " (" & pModele.GestDeProprietes.RecupererPropriete("Designation") & _
                            pModele.GestDeProprietes.RecupererPropriete("Description") _
                            & ") × " & _
                            pModele.Composant.Nb
             
        Set Feuille = Dessin.AjouterFeuille(NomFeuille)
        Feuille.Activer
        
        Debug.Print , , "Feuille : "; NomFeuille
        
        For Each ConfigDepliee In ListeDesConfigDepliee
            
            ConfigDepliee.Activer
            pModele.Reconstruire
            
            Enveloppe = Feuille.Enveloppe
            
            Set VueDepliee = Feuille.CreerVueDVP(Me, ConfigDepliee)
            
            If Not (VueDepliee Is Nothing) Then
                VueDepliee.OrienterVueDepliee cPortrait
                Centre.X = Enveloppe.MaxX + (VueDepliee.Dimensions.Lg * 0.5)
                Centre.Y = VueDepliee.Dimensions.Ht * 0.5
                VueDepliee.Centre = Centre
            End If
            
'            VueDepliee.Nom = "Config : " & ConfigDepliee.Nom & _
'                            " NomDuDossier : " & ConfigDepliee.CorpsDepliee.Dossier.Nom & _
'                            " NoDeDossier : " & ConfigDepliee.CorpsDepliee.Dossier.GestDeProprietes.RecupererPropriete(NO_DOSSIER)
            
        Next ConfigDepliee
        
        Feuille.RedimensionnerFeuille
        Feuille.ZoomEtendu
        
        Feuille.AjouterListeDesPiecesSoudees
        
        For Each VueDepliee In Feuille.ListeDesVues
            VueDepliee.InsererAnnotationLaser
        Next VueDepliee
        
        ConfigDeDepart.Activer
    Else
        Debug.Print , , "Pas de configuration dépliée"
    End If
    
    Set ListeDesConfigDepliee = Nothing
    Set VueDepliee = Nothing
    Set ConfigDeDepart = Nothing
    Set ConfigDepliee = Nothing
    
End Sub


