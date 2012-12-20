Attribute VB_Name = "Macro_AssemblageEtPiece"
Option Explicit

Private Const pNomClasse    As String = "AssemblageEtPiece"
Dim Erreur As Long

Public Sub ReconstruireTout()
    Dim Sw  As New ExtSldWorks
        Sw.Modele.ForcerAToutReconstruire
    Set Sw = Nothing
End Sub

Public Sub ImporterLesInfosClient()
    Dim Sw                  As New ExtSldWorks
    Dim ModeleDeBase        As ExtModele
    Dim Composant           As ExtComposant
    Dim Modele              As ExtModele
    Dim Rechercher          As ExtRechercher
    Dim GestFichiers        As New SysGestDeFichiers
    Dim Propriete           As ExtPropriete
    
    Set ModeleDeBase = Sw.Modele
    
    ModeleDeBase.GestDeMiseAJour.ImporterLesInfosClient "Infos.txt"
    
    Set Rechercher = ModeleDeBase.NouvelleRecherche
    Rechercher.PrendreEnCompteConfig = False
    Rechercher.PrendreEnCompteExclus = True
    
    For Each Composant In Rechercher.ListeDesComposants(cAssemblage + cPiece)
        Set Modele = Composant.Modele
        Modele.GestDeMiseAJour.ImporterLesInfosClient "Infos.txt"
    Next Composant
    
    Set Rechercher = Nothing
    Set Modele = Nothing
    Set Composant = Nothing
    Set ModeleDeBase = Nothing
    Set Sw = Nothing
End Sub

Public Sub MettreAJourLaListeDesPiecesSoudees()
    Dim Sw                  As New ExtSldWorks
    Dim ModeleDeBase        As ExtModele
    Dim Composant           As ExtComposant
    Dim Modele              As ExtModele
    Dim Rechercher          As ExtRechercher
    
    Set ModeleDeBase = Sw.Modele
    
    Set Rechercher = ModeleDeBase.NouvelleRecherche
    Rechercher.PrendreEnCompteConfig = False
    Rechercher.PrendreEnCompteExclus = False
    
    For Each Composant In Rechercher.ListeDesComposants(cPiece)
        Set Modele = Composant.Modele
        Debug.Print "-------------------------------------------------"
        Debug.Print Modele.Fichier.NomDuFichier
        Modele.Piece.GestDeMiseAJour.MettreAJourLaListeDesPiecesSoudees
        Modele.Piece.GestDeMiseAJour.MettreAJourLesNomsDeConfigs
    Next Composant
    
    Set Rechercher = Nothing
    Set Composant = Nothing
    Set Modele = Nothing
    Set ModeleDeBase = Nothing
    Set Sw = Nothing
    
End Sub


Sub ReconstruireLesFonctionBloquees()

    Dim Sw                  As SldWorks.SldWorks
    Dim Modele              As ModelDoc2
    Dim SelMgr              As SelectionMgr
    Dim Composant           As Component2
    Dim Piece               As ModelDoc2
    Dim i                   As Integer
    
    Set Sw = Application.SldWorks
    Set Modele = Sw.ActiveDoc
    Set SelMgr = Modele.SelectionManager
    
    For i = 1 To SelMgr.GetSelectedObjectCount2(-1)
        If SelMgr.GetSelectedObjectType3(i, -1) = swSelectType_e.swSelCOMPONENTS Then
            Set Composant = SelMgr.GetSelectedObject6(i, -1)
            Set Piece = Composant.GetModelDoc2
            Piece.Extension.UpdateFrozenFeatures True
        End If
    Next i
    
    Modele.ClearSelection2 True
    Modele.EditRebuild3
    
End Sub



Public Sub AjouterPlan()
    Dim Sw                  As SldWorks.SldWorks
    Dim ModeleActif         As ModelDoc2
    Dim SelMgr              As SelectionMgr
    Dim FeatMgr             As FeatureManager
    Dim NomPlan             As String
    Dim PlanDeRef           As RefPlane
    Dim FonctionPlanDeRef   As Feature
    Dim i                   As Integer
    
    Set Sw = Application.SldWorks
    Set ModeleActif = Sw.ActiveDoc
    Set SelMgr = ModeleActif.SelectionManager
    
    Set FonctionPlanDeRef = SelMgr.GetSelectedObject6(1, -1)
        
    If Not FonctionPlanDeRef.GetTypeName2 = "RefPlane" Then
        Sw.Frame.SetStatusBarText "L'élément selectionné n'est pas un plan"
        Exit Sub
    End If
    
    NomPlan = ValiderNomFonction(ModeleActif, FonctionPlanDeRef.Name)
    
    Set FeatMgr = ModeleActif.FeatureManager
    Set PlanDeRef = FeatMgr.InsertRefPlane(swRefPlaneReferenceConstraint_Coincident, 0, 0, 0, 0, 0)
    Set FonctionPlanDeRef = PlanDeRef
            
    FonctionPlanDeRef.Name = NomPlan
    
    Set Sw = Nothing
    Set ModeleActif = Nothing
    Set SelMgr = Nothing
    Set FeatMgr = Nothing
    Set PlanDeRef = Nothing
    Set FonctionPlanDeRef = Nothing
    
End Sub

Public Sub Contraindre_Origine()
    Dim Sw                  As SldWorks.SldWorks
    Dim Modele              As ModelDoc2
    Dim Composant           As Component2
    Dim SelMgr              As SelectionMgr
    
    Set Sw = Application.SldWorks
    Set Modele = Sw.ActiveDoc
    Set SelMgr = Modele.SelectionManager
    Set Composant = SelMgr.GetSelectedObject6(1, -1)
    
    If Modele.GetType = swDocASSEMBLY And Not Composant Is Nothing Then
        Call Contraindre(Modele, Composant)
    End If
    
    Modele.EditRebuild3
    
    Set Sw = Nothing
    Set Modele = Nothing
    Set Composant = Nothing
    Set SelMgr = Nothing
    
End Sub

Private Sub Contraindre(Modele As ModelDoc2, Composant As Component2)
    
    Dim FxAssFace       As Feature
    Dim FxAssDessus     As Feature
    Dim FxAssDroite     As Feature
    Dim FxCompFace      As Feature
    Dim FxCompDessus    As Feature
    Dim FxCompDroite    As Feature
    Dim FxContrainte    As Feature
    Dim Assemblage      As AssemblyDoc
    
    Set Assemblage = Modele
    
    '=================================================
    '  Boucle sur les fonctions du composant de base
    '=================================================
    
    'on recupère le composant en cours d'édition
    Set FxAssFace = Assemblage.GetEditTargetComponent.FirstFeature
    
    'si le mode "edition dans le contexte" n'est pas activé, on prend le modele
    If Modele.IsEditingSelf Then
        Set FxAssFace = Modele.FirstFeature
    End If
    
    'On boucle jusqu'au premier plan rencontré, c'est le plan de face. Les autres suivent : dessus, droite
    Do Until FxAssFace Is Nothing
        
        If FxAssFace.GetTypeName2 = "RefPlane" Then
            Set FxAssDessus = FxAssFace.GetNextFeature
            Set FxAssDroite = FxAssDessus.GetNextFeature
            Exit Do
        End If
        
        Set FxAssFace = FxAssFace.GetNextFeature
    Loop
    
    
    '======================================================
    '  Boucle sur les fonction du composant à contraindre
    '======================================================
    Set FxCompFace = Composant.FirstFeature
    
    Do Until FxCompFace Is Nothing
        
        If FxCompFace.GetTypeName2 = "RefPlane" Then
            Set FxCompDessus = FxCompFace.GetNextFeature
            Set FxCompDroite = FxCompDessus.GetNextFeature
            Exit Do
        End If
        
        Set FxCompFace = FxCompFace.GetNextFeature
    Loop
    
    
    '===========================
    '  Création des containtes
    '===========================
        
    'on libère le composant à contraindre
    Assemblage.UnfixComponent
    
    Modele.ClearSelection2 True
    
    'Ajout de la contrainte : Plan de face
    Call Modele.Extension.SelectByID2(FxAssFace.GetNameForSelection(swSelectType_e.swSelDATUMPLANES), "PLANE", 0, 0, 0, False, 1, Nothing, 0)
    Call Modele.Extension.SelectByID2(FxCompFace.GetNameForSelection(swSelectType_e.swSelDATUMPLANES), "PLANE", 0, 0, 0, True, 1, Nothing, 0)
    Set FxContrainte = Modele.AddMate3(swMateCOINCIDENT, swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, Erreur)
    FxContrainte.Name = Composant.Name2 & " " & FxCompFace.Name & " /"
    
    'Ajout de la contrainte : Plan de dessus
    Call Modele.Extension.SelectByID2(FxAssDessus.GetNameForSelection(swSelectType_e.swSelDATUMPLANES), "PLANE", 0, 0, 0, False, 1, Nothing, 0)
    Call Modele.Extension.SelectByID2(FxCompDessus.GetNameForSelection(swSelectType_e.swSelDATUMPLANES), "PLANE", 0, 0, 0, True, 1, Nothing, 0)
    Set FxContrainte = Modele.AddMate3(swMateCOINCIDENT, swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, Erreur)
    FxContrainte.Name = Composant.Name2 & " " & FxCompDessus.Name & " /"
    
    'Ajout de la contrainte : Plan de droite
    Call Modele.Extension.SelectByID2(FxAssDroite.GetNameForSelection(swSelectType_e.swSelDATUMPLANES), "PLANE", 0, 0, 0, False, 1, Nothing, 0)
    Call Modele.Extension.SelectByID2(FxCompDroite.GetNameForSelection(swSelectType_e.swSelDATUMPLANES), "PLANE", 0, 0, 0, True, 1, Nothing, 0)
    Set FxContrainte = Modele.AddMate3(swMateCOINCIDENT, swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, Erreur)
    FxContrainte.Name = Composant.Name2 & " " & FxCompDroite.Name & " /"
    
    Modele.ClearSelection2 True
    
    Set FxAssFace = Nothing
    Set FxAssDessus = Nothing
    Set FxAssDroite = Nothing
    Set FxCompFace = Nothing
    Set FxCompDessus = Nothing
    Set FxCompDroite = Nothing
    Set FxContrainte = Nothing
    Set Assemblage = Nothing

End Sub


Private Function ValiderNomFonction(Modele As ModelDoc2, Nom As String) As String
    Dim NomTmp As String
    Dim i As Integer
    
    i = 2
    NomTmp = Nom
    ValiderNomFonction = NomTmp
    
    While Modele.FeatureManager.IsNameUsed(swFeatureName, NomTmp)
        NomTmp = Nom & " " & i
        ValiderNomFonction = NomTmp
        i = i + 1
    Wend
    
End Function
