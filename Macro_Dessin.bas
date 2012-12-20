Attribute VB_Name = "Macro_Dessin"
Option Explicit

Public Sub SauverLesFeuillesEnDXF()

    Dim Sw                  As New ExtSldWorks
    Dim GestDeFichier       As New SysGestDeFichiers
    Dim DossierDXF          As String
    Dim Dessin              As ExtDessin
    Dim Feuille             As ExtFeuille
    
    Set Dessin = Sw.Modele.Dessin
    GestDeFichier.Chemin = Dessin.Modele.Fichier.Chemin
    DossierDXF = GestDeFichier.CreerDossier("DXF")
    
    Debug.Print Dessin.Modele.Fichier.NomDuFichier
    
    For Each Feuille In Dessin.ListeDesFeuilles
        Feuille.SauverEnDXF DossierDXF
    Next Feuille
    
    Set Feuille = Nothing
    Set Dessin = Nothing
    Set GestDeFichier = Nothing
    
End Sub

Public Sub InsererAnnotation()

    Dim Sw                  As SldWorks.SldWorks
    Dim Modele              As ModelDoc2
    Dim SelMgr              As SelectionMgr
    Dim DocExtension        As ModelDocExtension
    Dim Dessin              As DrawingDoc
    Dim Note                As Note
    Dim Annotation          As Annotation
    Dim TextForm            As TextFormat
    
    Set Sw = Application.SldWorks
    Set Modele = Sw.ActiveDoc
    Set SelMgr = Modele.SelectionManager
    Set DocExtension = Modele.Extension
    Set Dessin = Modele
    
    Set Note = Modele.InsertNote("$PRPWLD:""" & NOM_ELEMENT & """" & Chr(13) & Chr(10) & "$PRPWLD:""Profil""")
    
    If Note.GetText Like ("*" & Chr(13) & Chr(10)) Then
'        Note.SetText "$PRPWLD:""" & NOM_ELEMENT & """" & Chr(13) & Chr(10) & "$PRPWLD:""Matériau"" ep $PRPWLD:""Epaisseur de tôlerie"""
        Note.SetText "$PRPWLD:""" & NOM_ELEMENT & """" & Chr(13) & Chr(10) & "$PRPWLD:""Matériau"" ep $PRPWLD:""" & EPAISSEUR_DE_TOLE & """"
    End If
    
    If Not Note Is Nothing Then
        Set Annotation = Note.GetAnnotation
        
        Note.SetTextJustification swTextJustification_e.swTextJustificationLeft
        Annotation.SetLeader3 swLeaderStyle_e.swBENT, swLeaderSide_e.swLS_SMART, True, False, False, False

        Modele.GraphicsRedraw2
        
    End If
    
End Sub

Public Sub InsererProfil()

    Dim Sw                  As SldWorks.SldWorks
    Dim Document            As ModelDoc2
    Dim DocExtension        As ModelDocExtension
    Dim Dessin              As DrawingDoc
    Dim Note                As Note
    Dim Annotation          As Annotation
    
    Set Sw = Application.SldWorks
    Set Document = Sw.ActiveDoc
    Set DocExtension = Document.Extension
    Set Dessin = Document
    
    Set Note = DocExtension.InsertBOMBalloon(swBalloonStyle_e.swBS_None, _
                                            swBalloonFit_e.swBF_Tightest, _
                                            swBalloonTextContent_e.swBalloonTextCustom, _
                                            "$PRPWLD:""cProfil""", _
                                            swBalloonTextContent_e.swBalloonTextCustom, _
                                            "", _
                                            swBalloonFit_e.swBF_UserDef, _
                                            False, _
                                            0, _
                                            "")
    If Not Note Is Nothing Then
        Set Annotation = Note.GetAnnotation
        
        Note.SetTextJustification swTextJustification_e.swTextJustificationLeft
        Annotation.SetLeader3 swLeaderStyle_e.swBENT, swLeaderSide_e.swLS_SMART, True, True, False, False
        
        Document.GraphicsRedraw2
        
    End If
    
End Sub

Public Sub InsererLaserEtQuantite()

    Dim Sw                  As SldWorks.SldWorks
    Dim Document            As ModelDoc2
    Dim DocExtension        As ModelDocExtension
    Dim SelMgr              As SelectionMgr
    Dim Dessin              As DrawingDoc
    Dim Note                As Note
    Dim Annotation          As Annotation
    Dim Vue                 As View
    Dim DrawComp            As DrawingComponent
    Dim Comp                As Component2
    Dim Face                As Face2
    Dim DimFace             As Variant
    Dim FaceBox(1)          As Double
    Dim v                   As Variant
    Dim VueBox(1)           As Variant
    Dim VueCentre(1)        As Variant
    Dim Entite              As Variant
    
    Set Sw = Application.SldWorks
    Set Document = Sw.ActiveDoc
    Set DocExtension = Document.Extension
    Set Dessin = Document
    
    '=========================================================='
    '
    '           Position de la description de la tôle
    '
    '=========================================================='
    Set SelMgr = Document.SelectionManager
    Set Vue = SelMgr.GetSelectedObject6(1, -1)
    
    'On recupere le composant
    Set DrawComp = Vue.RootDrawingComponent
    Set Comp = DrawComp.Component
    
    'On recupere la face
    Entite = Vue.GetVisibleEntities(Comp, swViewEntityType_e.swViewEntityType_Face)
    Set Face = Entite(0)
    DimFace = Face.GetBox
    
    'Dimensions du contour de la vue pour avoir une idée de l'orientation de la face
    v = Vue.GetOutline
    VueBox(0) = v(2) - v(0)
    VueBox(1) = v(3) - v(1)
    
    'Calcul du centre de la face, on estime qu'il est au centre de la vue
    VueCentre(0) = (v(2) + v(0)) * 0.5
    VueCentre(1) = (v(3) + v(1)) * 0.5
    
    'On calcul les dimensions de la bounding box de la face
    'puis on supprime la dimension correpondant à l'épaisseur
    
    Dim i As Integer, j As Integer
    Dim Tmp As Double
    
    j = 0
    Tmp = DimFace(3) - DimFace(0)
    
    For i = 1 To 2
        If (DimFace(i + 3) - DimFace(i)) < Tmp Then
            Tmp = DimFace(i + 3) - DimFace(i)
        End If
    Next i
    
    For i = 0 To 2
        If (DimFace(i + 3) - DimFace(i)) > Tmp Then
            FaceBox(j) = DimFace(i + 3) - DimFace(i)
            j = j + 1
        End If
    Next i
    
    'On recherche la dimension correspondant au y
    If PositionMax(VueBox) = 1 Then
        VueCentre(1) = VueCentre(1) - (FaceBox(PositionMax(FaceBox)) * 0.5) - 0.04
    Else
        VueCentre(1) = VueCentre(1) - (FaceBox(PositionMin(FaceBox)) * 0.5) - 0.04
    End If
    
    
    
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
        Annotation.SetPosition VueCentre(0), VueCentre(1), 0#
        
        Document.GraphicsRedraw2
        
    End If
    
End Sub

Private Function PositionMax(v As Variant) As Integer
    PositionMax = 0
    Dim i As Integer
    
    For i = 1 To UBound(v)
        If v(i) > v(i - 1) Then
            PositionMax = i
        End If
    Next i
End Function

Private Function PositionMin(v As Variant) As Integer
    PositionMin = 0
    Dim i As Integer
    
    For i = 1 To UBound(v)
        If v(i) < v(i - 1) Then
            PositionMin = i
        End If
    Next i
End Function

