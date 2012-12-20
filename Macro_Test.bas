Attribute VB_Name = "Macro_Test"



Public Sub Macro3()
    Dim Sw                  As New ExtSldWorks
    Dim Dessin              As ExtDessin
    Dim Vue                 As ExtVue
    Dim Config              As ExtConfiguration
    Dim Corps               As ExtCorps
    Dim Fonction            As ExtFonction
    Dim Esquisse            As Sketch
    Dim vSegments           As Variant
    Dim vLigne              As Variant
    Dim Ligne               As SketchLine
    Dim PointD              As SketchPoint
    Dim PointA              As SketchPoint
    Dim AngleRadian         As Double

    
    Debug.Print " "
    Debug.Print "------------------------------------"
    
    Set Dessin = Sw.Modele.Dessin
    
    For Each Vue In Dessin.FeuilleActive.ListeDesVues
        Debug.Print Vue.Nom, Vue.ConfigurationDeReference.Nom
        Vue.OrienterVueDepliee cPortrait
    Next Vue
    
    Set Dessin = Nothing
    Set Sw = Nothing
    
End Sub


Public Sub ActiverChaqueConfiguration()
    
    Dim Sw                  As New ExtSldWorks
    Dim ModeleDeBase        As ExtModele
    Dim Composant           As ExtComposant
    Dim Configuration       As ExtConfiguration
    Dim Rechercher          As ExtRechercher
    
    Set ModeleDeBase = Sw.Modele
    Set Rechercher = ModeleDeBase.NouvelleRecherche
    Rechercher.PrendreEnCompteConfig = False
    Rechercher.PrendreEnCompteExclus = True
    
    For Each Composant In Rechercher.ListeDesComposants(cPiece)
        
        Debug.Print Composant.Modele.Fichier.NomDuFichier
        
        For Each Configuration In Composant.Modele.GestDeConfigurations.ListerLesConfigs(cDepliee)
            
            Debug.Print , Configuration.Nom
            
            Configuration.Activer
            
            Debug.Print , Configuration.CorpsDepliee.Dossier.Nom, Configuration.CorpsDepliee.Nom
            
        Next Configuration
        
    Next Composant
    
    Set Rechercher = Nothing
    Set Composant = Nothing
    Set ModeleDeBase = Nothing
    Set Sw = Nothing
    
End Sub
