VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExtVue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private pSwVue                  As View
Private pFeuille                As ExtFeuille
Private pModeleRef              As New ExtModele
Private pConfigurationRef       As New ExtConfiguration
Private Const pNomClasse        As String = "ExtVue"
Private Erreur As Long, Warning As Long

Private Sub Class_Terminate()
    Set pFeuille = Nothing
End Sub

Public Function SetVue(Vue As View, Feuille As ExtFeuille) As Boolean
    
    If Not (Vue Is Nothing) And Not (Feuille Is Nothing) Then
        SetVue = True
        Set pSwVue = Vue
        Set pFeuille = Feuille
        Exit Function
    End If
    
    Debug.Print pNomClasse & ".SetModele : le Modele de base n'est pas valide"
    
End Function

Public Function swVue() As View
    Set swVue = pSwVue
End Function

Public Property Get Nom() As String
    Nom = swVue.GetName2
End Property

Public Property Let Nom(s As String)
    swVue.SetName2 Nom
End Property

Public Property Get Dimensions() As Dimensions
    Dim DimVue As Variant
    
    DimVue = swVue.GetOutline
    Dimensions.Lg = Rectangle.MaxX - Rectangle.MinX
    Dimensions.Ht = Rectangle.MaxY - Rectangle.MinY
    
End Property

Public Property Get Rectangle() As Rectangle
    Dim DimVue As Variant
    DimVue = swVue.GetOutline
    Rectangle.MinX = DimVue(0)
    Rectangle.MinY = DimVue(1)
    Rectangle.MaxX = DimVue(2)
    Rectangle.MaxY = DimVue(3)
End Property

Public Property Get Centre() As Point
    Dim DimVue As Variant
    DimVue = swVue.Position
    
    Centre.X = DimVue(0)
    Centre.Y = DimVue(1)
End Property

Public Property Let Centre(Pt As Point)
    
    Dim DimVue(1) As Double
    DimVue(0) = Pt.X
    DimVue(1) = Pt.Y
    
    swVue.Position = DimVue
End Property

Public Function ModeleDeReference() As ExtModele
    
    pModeleRef.SetModele swVue.ReferencedDocument, pFeuille.Dessin.Modele.Sw
    Set ModeleDeReference = pModeleRef
    
End Function

Public Function ConfigurationDeReference() As ExtConfiguration
    
    pConfigurationRef.SetConfiguration ModeleDeReference.swModele.GetConfigurationByName(swVue.ReferencedConfiguration), Me.ModeleDeReference
    Set ConfigurationDeReference = pConfigurationRef
    
End Function


Public Sub OrienterVueDepliee(Orientation As Orientation_e)
    If swVue.IsFlatPatternView = False Then Exit Sub
    
    Debug.Print , , "Orientation de la vue : "; Nom
    
    Dim Fonction            As ExtFonction
    Dim swEsquisse          As Sketch
    Dim vSegments           As Variant
    Dim vLigne              As Variant
    Dim swLigne             As SketchLine
    Dim Ligne               As ExtLigne
    Dim LigneDeRotation     As New ExtLigne
    Dim LigneDeBase         As New ExtLigne
    Dim swPointDep          As SketchPoint
    Dim swPointArr          As SketchPoint
    Dim AngleRadian         As Double
    
    Dim swMath              As MathUtility
    Dim nVecteurData(2)     As Double
    Dim vVecteurData        As Variant
    Dim swEsquisseVecteur   As MathVector
    Dim swModeleVecteur     As MathVector
    Dim swVueVecteur        As MathVector
    Dim swVueXform          As MathTransform
    Dim swEsquisseXform     As MathTransform
    Dim VectX               As Double
    Dim VectY               As Double
    
    Set swMath = pFeuille.Dessin.Modele.Sw.swSw.GetMathUtility
    Set swVueXform = swVue.ModelToViewTransform
    
    'On récupère l'swEsquisse
    Set Fonction = Me.ConfigurationDeReference.CorpsDepliee.Tolerie.CubeDeVisualisation
    Set swEsquisse = Fonction.swFonction.GetSpecificFeature2
    Set swEsquisseXform = swEsquisse.ModelToSketchTransform
    
    'On liste les lignes de l'swEsquisse, il y en a normalement quatre
    vSegments = swEsquisse.GetSketchSegments
    
    'S'il n'y a pas de segemnts, on sort de la sub
    If IsEmpty(vSegments) Then
        Debug.Print , , , "Pas de segements dans le cube de visualisation"
        Exit Sub
    End If
    
    'Pour comparer, on assigne le première ligne
    Set swLigne = vSegments(0)
    Set swPointDep = swLigne.GetStartPoint2
    Set swPointArr = swLigne.GetEndPoint2
    LigneDeRotation.Depart.X = swPointDep.X: LigneDeRotation.Depart.Y = swPointDep.Y: LigneDeRotation.Depart.Z = swPointDep.Z
    LigneDeRotation.Arrivee.X = swPointArr.X: LigneDeRotation.Arrivee.Y = swPointArr.Y: LigneDeRotation.Arrivee.Z = swPointArr.Z
    'On boucle sur les lignes
    For Each vLigne In vSegments
        Set swLigne = vLigne
        Set swPointDep = swLigne.GetStartPoint2
        Set swPointArr = swLigne.GetEndPoint2
        Set Ligne = New ExtLigne
        Ligne.Depart.X = swPointDep.X: Ligne.Depart.Y = swPointDep.Y: Ligne.Depart.Z = swPointDep.Z
        Ligne.Arrivee.X = swPointArr.X: Ligne.Arrivee.Y = swPointArr.Y: Ligne.Arrivee.Z = swPointArr.Z
        
        If Ligne.Longeur > LigneDeRotation.Longeur Then
            Set LigneDeRotation = Ligne
        End If
        
    Next vLigne
    
    'Recherche l'orientation de l'esquisse par rapport à la vue.
    'on cree un vecteur  dans la direction de X [1;0]
    nVecteurData(0) = 1: nVecteurData(1) = 0: nVecteurData(2) = 0
    vVecteurData = nVecteurData
    'Creation du vecteur
    Set swEsquisseVecteur = swMath.CreateVector(vVecteurData)
    'Convertion du vecteur en coordonnées du Modele
    Set swModeleVecteur = swEsquisseVecteur.MultiplyTransform(swEsquisseXform.Inverse)
    'Convertion du vecteur en coordonnées de la Vue
    Set swVueVecteur = swModeleVecteur.MultiplyTransform(swVueXform)
    'Création d'une ligne pour récupérer son angle
    LigneDeBase.Arrivee.X = ExtMath.Arrondi(swVueVecteur.ArrayData(0), 0)
    LigneDeBase.Arrivee.Y = ExtMath.Arrondi(swVueVecteur.ArrayData(1), 0)
    LigneDeBase.Arrivee.Z = ExtMath.Arrondi(swVueVecteur.ArrayData(2), 0)
    
    'On ajoute les deux angle
    AngleRadian = ExtMath.Angle0To2PI(LigneDeRotation.AngleXYRadian + LigneDeBase.AngleXYRadian)
    
    'Selection de l'orientation et calcul de l'angle mini à parcourir
    If Orientation = cPortrait Then
        'On aligne le plus grand coté avec la verticale
        'Langle de rotation est donc égal à : Angle - 90°
        AngleRadian = AngleRadian - (Pi * 0.5)
    Else
        'On aligne le plus grand coté avec l'horizontal
        'Si l'angle est supérieur à 90° on prend le complémentaire à 180°
        If AngleRadian >= (0.5 * Pi) Then
            AngleRadian = AngleRadian - Pi
        End If
    End If
    
    'Et on l'applique à la vue
    swVue.Angle = AngleRadian * -1#
    
    Set swMath = Nothing
    Set swEsquisseVecteur = Nothing
    Set swModeleVecteur = Nothing
    Set swVueVecteur = Nothing
    Set swVueXform = Nothing
    Set swEsquisseXform = Nothing
    Set Fonction = Nothing
    Set swEsquisse = Nothing
    Set swLigne = Nothing
    Set LigneDeBase = Nothing
    Set Ligne = Nothing
    Set LigneDeRotation = Nothing
    Set swPointDep = Nothing
    Set swPointArr = Nothing
    
End Sub

Public Sub InsererAnnotationLaser()

    If swVue.IsFlatPatternView = False Then Exit Sub
    
    Dim ModeleDoc           As ModelDoc2
    Dim DocExtension        As ModelDocExtension
    Dim Vue                 As View
    Dim vListeComposants    As Variant
    Dim vListeFaces         As Variant
    Dim SelManager          As SelectionMgr
    Dim Face                As Face2
    Dim Entite              As Entity
    Dim SelData             As SelectData
    
    Set ModeleDoc = pFeuille.Dessin.Modele.swModele
    Set DocExtension = ModeleDoc.Extension
    Set SelManager = ModeleDoc.SelectionManager
    Set SelData = SelManager.CreateSelectData
    
    Set Vue = swVue
    vListeComposants = Vue.GetVisibleComponents
    vListeFaces = Vue.GetVisibleEntities(vListeComposants(0), swViewEntityType_e.swViewEntityType_Face)
    Set Face = vListeFaces(0)
    Set Entite = Face
    SelData.View = swVue
    SelData.X = Centre.X
    SelData.Y = Centre.Y
    SelData.Z = 0
    Entite.Select4 False, SelData
    
    Dim Note                As Note
    Dim Annotation          As Annotation
    
    '=========================================================='
    '
    '                  Insertion des notes
    '
    '=========================================================='
    
    'la référence de la pièce à graver
    Set Note = DocExtension.InsertBOMBalloon(swBalloonStyle_e.swBS_None, _
                                            swBalloonFit_e.swBF_Tightest, _
                                            swBalloonTextContent_e.swBalloonTextCustom, _
                                            "$PRPWLD:""SW-Nom de fichier(File Name)""-$PRPWLD:""NoConfig""-$PRPWLD:""" & NO_DOSSIER & """", _
                                            swBalloonTextContent_e.swBalloonTextCustom, _
                                            "", _
                                            swBalloonFit_e.swBF_UserDef, _
                                            False, _
                                            0, _
                                            "")
    
    If Not Note Is Nothing Then
        Set Annotation = Note.GetAnnotation
        
        Note.SetTextJustification swTextJustification_e.swTextJustificationCenter
        Annotation.SetLeader3 swLeaderStyle_e.swNO_LEADER, swLeaderSide_e.swLS_SMART, True, False, False, False
        Annotation.Layer = "GRAVURE"
        
    End If
    
    'la quantité du corps dans la pièce
    
'    Set Note = DocExtension.InsertBOMBalloon(swBalloonStyle_e.swBS_None, _
'                                            swBalloonFit_e.swBF_Tightest, _
'                                            swBalloonTextContent_e.swBalloonTextCustom, _
'                                            "$PRPWLD:""SW-Nom de fichier(File Name)""-$PRPWLD:""NoConfig""-$PRPWLD:""" & NO_DOSSIER & """ - $PRPWLD:""Matériau"" ep $PRPWLD:""Epaisseur de tôlerie""", _
'                                            swBalloonTextContent_e.swBalloonTextCustom, _
'                                            "", _
'                                            swBalloonFit_e.swBF_UserDef, _
'                                            True, _
'                                            0, _
'                                            "×  ")
'                                            '"×"
    
    Set Note = DocExtension.InsertBOMBalloon(swBalloonStyle_e.swBS_None, _
                                            swBalloonFit_e.swBF_Tightest, _
                                            swBalloonTextContent_e.swBalloonTextCustom, _
                                            "$PRPWLD:""SW-Nom de fichier(File Name)""-$PRPWLD:""NoConfig""-$PRPWLD:""" & NO_DOSSIER & """ [ $PRPWLD:""Matériau"" ] ( ep$PRPWLD:""" & EPAISSEUR_DE_TOLE & """ )", _
                                            swBalloonTextContent_e.swBalloonTextCustom, _
                                            "", _
                                            swBalloonFit_e.swBF_UserDef, _
                                            True, _
                                            0, _
                                            "×  ")
                                            '"×"
    If Not Note Is Nothing Then
        Set Annotation = Note.GetAnnotation
        
        Note.SetTextJustification swTextJustification_e.swTextJustificationCenter
        Annotation.SetLeader3 swLeaderStyle_e.swNO_LEADER, swLeaderSide_e.swLS_SMART, True, True, False, False
        Annotation.Layer = "QUANTITE"
        Annotation.SetPosition Centre.X, (Centre.Y - (Dimensions.Ht * 0.5)), 0#
        
        pFeuille.Dessin.Modele.Redessiner
        
    End If
    
    ModeleDoc.ClearSelection2 True
    
    Set SelData = Nothing
    Set Entite = Nothing
    Set Face = Nothing
    Set SelManager = Nothing
    Set Vue = Nothing
    Set DocExtension = Nothing
    Set ModeleDoc = Nothing
    
End Sub
