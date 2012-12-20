VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExtFeuille"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=======================================================================================
Option Explicit

Private pSwFeuille              As Sheet
Private pDessin                 As ExtDessin
Private Const pNomClasse        As String = "ExtFeuille"
Private Erreur As Long, Warning As Long

Private Sub Class_Terminate()
    Set pSwFeuille = Nothing
    Set pDessin = Nothing
End Sub

Public Function Dessin() As ExtDessin
    Set Dessin = pDessin
End Function

Public Function SetFeuille(Feuille As Sheet, Dessin As ExtDessin) As Boolean
    
    If Not (Feuille Is Nothing) And Not (Dessin Is Nothing) Then
        SetFeuille = True
        Set pSwFeuille = Feuille
        Set pDessin = Dessin
        Exit Function
    End If
    
    Debug.Print pNomClasse & ".SetFeuille : le Modele de base n'est pas valide"
    
End Function

Public Property Get Nom() As String
    Nom = swFeuille.GetName
End Property

Public Property Let Nom(s As String)
    swFeuille.SetName s
End Property

Public Sub Activer()
    Dim Dessin As DrawingDoc
    
    Set Dessin = pDessin.Modele.swModele
    Dessin.ActivateSheet swFeuille.GetName
    
    Set Dessin = Nothing
End Sub

Public Sub Supprimer()
    
    pDessin.Modele.swModele.Extension.SelectByID2 swFeuille.GetName, "SHEET", 0, 0, 0, False, 0, Nothing, 0
    pDessin.Modele.swModele.DeleteSelection False
    pDessin.Modele.swModele.ClearSelection2 True
    
End Sub

Public Sub ZoomEtendu()
    
    pDessin.Modele.swModele.ViewZoomtofit2
    
End Sub

Public Function CreerVueDVP(Piece As ExtPiece, Config As ExtConfiguration) As ExtVue
    
    Dim DessinDoc As DrawingDoc
    Set DessinDoc = pDessin.Modele.swModele
    Set CreerVueDVP = New ExtVue
    
    CreerVueDVP.SetVue DessinDoc.CreateFlatPatternViewFromModelView3(Piece.Modele.Fichier.Chemin, Config.Nom, 0#, 0#, 0#, False, False), Me
    
    CreerVueDVP.Nom = Config.Nom
    
    Set DessinDoc = Nothing

End Function

Public Function PremiereVue() As ExtVue
    
    If ListeDesVues.Count Then
        Set PremiereVue = ListeDesVues.Item(1)
    End If
    
End Function

Public Function ListeDesVues() As Collection
    
    
    Dim vVue    As Variant
    Dim Vue     As View
    Dim VueExt  As ExtVue
    
    Set ListeDesVues = New Collection
    
    vVue = swFeuille.GetViews
    
    If Not IsEmpty(vVue) Then
        For Each vVue In swFeuille.GetViews
            
            Set Vue = vVue
            'If Not (Vue.Name Like "[*]*") Then
                Set VueExt = New ExtVue
                VueExt.SetVue Vue, Me
                ListeDesVues.Add VueExt
                Set VueExt = Nothing
            'End If
        
        Next vVue
    End If
    
    Set Vue = Nothing

End Function

Public Function Enveloppe() As Rectangle
    
    Dim Vue                 As ExtVue
    Dim Rect                As Rectangle
    
    Enveloppe.MaxX = 0#
    Enveloppe.MaxY = 0#
    Enveloppe.MinX = 10000#
    Enveloppe.MinY = 10000#
    
    's'il n'y a pas de vue on met tout à 0
    If ListeDesVues.Count Then
        For Each Vue In ListeDesVues
            
            Enveloppe.MaxX = ExtMath.Max(Enveloppe.MaxX, Vue.Rectangle.MaxX)
            Enveloppe.MaxY = ExtMath.Max(Enveloppe.MaxY, Vue.Rectangle.MaxY)
            Enveloppe.MinX = ExtMath.Min(Enveloppe.MinX, ExtMath.Max(0, Vue.Rectangle.MinX))
            Enveloppe.MinY = ExtMath.Min(Enveloppe.MinY, ExtMath.Max(0, Vue.Rectangle.MinY))
        
        Next Vue
    Else
        Enveloppe.MinX = 0#
        Enveloppe.MinY = 0#
    End If
    
    Set Vue = Nothing
    
End Function

Public Sub RedimensionnerFeuille()
    
    Dim Env         As Rectangle
    Dim DimFeuille  As Dimensions
    
    Env = Enveloppe
    
    DimFeuille.Lg = (Int((Env.MaxX + Env.MinX) * 200) + 1) / 200
    DimFeuille.Ht = (Int((Env.MaxY + Env.MinY) * 200) + 1) / 200
    
    swFeuille.SetSize swDwgPaperSizes_e.swDwgPapersUserDefined, DimFeuille.Lg, DimFeuille.Ht
    
End Sub

Public Sub AjouterListeDesPiecesSoudees()
    
    Debug.Print , "AjouterListeDesPiecesSoudees"
    
    Dim Modele              As ExtModele
    Dim ConfigRacine        As ExtConfiguration
    Dim ModeleDoc           As ModelDoc2
    Dim sVue                As View
    Dim WeldCutList         As WeldmentCutListAnnotation
    Dim WeldCutListFeat     As WeldmentCutListFeature
    Dim Fonction            As Feature
    Dim TemplateCutList     As String
    Dim CutListName         As String
    
    Set Modele = PremiereVue.ModeleDeReference
    Set ConfigRacine = PremiereVue.ConfigurationDeReference.ConfigurationRacine
    
    Set sVue = PremiereVue.swVue
    Set ModeleDoc = sVue.ReferencedDocument
    
    TemplateCutList = pDessin.Modele.Sw.RecupererCle(swFileLocationsWeldmentCutListTemplates) & "\Laser.sldwldtbt"
    
    If sVue.GetKeepLinkedToBOM Then
        Debug.Print sVue.GetKeepLinkedToBOMName
        CutListName = sVue.GetKeepLinkedToBOMName
        Set sVue = sVue.GetNextView
    Else
        Set WeldCutList = sVue.InsertWeldmentTable(False, 0#, 0#, swBOMConfigurationAnchorType_e.swBOMConfigurationAnchor_TopRight, ConfigRacine.Nom, TemplateCutList)
        If Not (WeldCutList Is Nothing) Then
            Set WeldCutListFeat = WeldCutList.WeldmentCutListFeature
            Set Fonction = WeldCutListFeat.GetFeature
            Fonction.Name = "Lst \ " & Me.Nom
            CutListName = Fonction.Name
        Else
            Debug.Print "Table non Cree"
        End If
    End If
    
    Do Until sVue Is Nothing
        
        sVue.SetKeepLinkedToBOM True, CutListName
        
        Set sVue = sVue.GetNextView
    Loop
    
    Set Fonction = Nothing
    Set WeldCutListFeat = Nothing
    Set WeldCutList = Nothing
    Set ConfigRacine = Nothing
    Set sVue = Nothing
    Set ModeleDoc = Nothing
End Sub

Public Sub SauverEnDXF(ByVal CheminDossier As String)
    
    Debug.Print , "SauverEnDXF : "; Me.Nom
    
    Dim Sw As SldWorks.SldWorks
    
    Set Sw = Application.SldWorks
    
    'Laisser les splines en spline
    Sw.SetUserPreferenceToggle swUserPreferenceToggle_e.swDxfExportSplinesAsSplines, True
    'Exporter uniquement la feuille active
    Sw.SetUserPreferenceIntegerValue swUserPreferenceIntegerValue_e.swDxfMultiSheetOption, swDxfMultisheet_e.swDxfActiveSheetOnly
    'Style des lignes : Solidworks
    Sw.SetUserPreferenceIntegerValue swUserPreferenceIntegerValue_e.swDxfOutputLineStyles, 1
    
    Activer
    ZoomEtendu
    CheminDossier = CheminDossier & "\" & Nom & ".dxf"
    Debug.Print , CheminDossier
    pDessin.Modele.swModele.Extension.SaveAs CheminDossier, swSaveAsVersion_e.swSaveAsCurrentVersion, swSaveAsOptions_e.swSaveAsOptions_Silent, Nothing, Erreur, Warning
    
    Set Sw = Nothing
    
End Sub

Public Function swFeuille() As Sheet
    Set swFeuille = pSwFeuille
End Function
