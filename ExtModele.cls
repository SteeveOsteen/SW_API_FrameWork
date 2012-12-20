VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExtModele"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private pSw                     As ExtSldWorks
Private pSwModeleDoc            As ModelDoc2
Private pComposant              As ExtComposant
Private pPiece                  As New ExtPiece
Private pAssemblage             As New ExtAssemblage
Private pDessin                 As New ExtDessin
Private pGestDeConfigurations   As New ExtGestDeConfigurations
Private pGestDeProprietes       As New ExtGestDeProprietes
Private pGestDeMiseAJour        As New GestDeMiseAJourModele
Private pFichier                As New SysFichier
Private Const pNomClasse        As String = "ExtModele"
Private Erreur As Long, Warning As Long

Private Sub Class_Terminate()
    Set pSw = Nothing
    Set pSwModeleDoc = Nothing
    Set pComposant = Nothing
    Set pPiece = Nothing
    Set pAssemblage = Nothing
    Set pDessin = Nothing
    Set pGestDeConfigurations = Nothing
    Set pGestDeProprietes = Nothing
    Set pGestDeMiseAJour = Nothing
    Set pFichier = Nothing
End Sub

Public Function SetModele(ByRef Modele As ModelDoc2, Sw As ExtSldWorks) As Boolean
    
    If Not (Modele Is Nothing) And Not (Sw Is Nothing) Then
        
        SetModele = True
        Set pSw = Sw
        Set pSwModeleDoc = Modele
        pFichier.Chemin = pSwModeleDoc.GetPathName
        pGestDeProprietes.SetGestionnaire swModele.Extension.CustomPropertyManager(""), Me
        pGestDeConfigurations.SetModele Me
        pGestDeMiseAJour.SetModele Me
        
        If Est(cAssemblage + cPiece) Then
            Set pComposant = New ExtComposant
            pComposant.SetComposant swModele.ConfigurationManager.ActiveConfiguration.GetRootComponent3(True), Me
        End If
        
        Select Case TypeDuModele
            Case cAssemblage
                pAssemblage.SetModele Me
            Case cPiece
                pPiece.SetModele Me
            Case cDessin
                pDessin.SetModele Me
            Case Else
        End Select
        
        Exit Function
    End If
    
    Debug.Assert False
    Debug.Print pNomClasse & ".SetModele : le Modele de base n'est pas valide"
    
End Function

Public Function SetComposant(ByRef Composant As ExtComposant) As Boolean
    
    If Not (Composant Is Nothing) Then
        Set pComposant = Composant
        Exit Function
    End If
    
    Debug.Print pNomClasse & ".SetComposant : le Modele de base n'est pas valide"
    
End Function

Public Function TypeDuModele() As TypeFichier_e
    
    Dim GestDeFichier As New SysGestDeFichiers
    GestDeFichier.Chemin = swModele.GetPathName
    
    TypeDuModele = GestDeFichier.TypeDuFichier
    
    Set GestDeFichier = Nothing
    
End Function

Public Function Est(T As TypeFichier_e) As Boolean
    Est = T And TypeDuModele
End Function

Public Function Fichier() As SysFichier
    Set Fichier = pFichier
End Function

Public Function GestDeConfigurations() As ExtGestDeConfigurations
    Set GestDeConfigurations = pGestDeConfigurations
End Function

Public Function GestDeProprietes() As ExtGestDeProprietes
    Set GestDeProprietes = pGestDeProprietes
End Function

Public Function GestDeMiseAJour() As GestDeMiseAJourModele
    Set GestDeMiseAJour = pGestDeMiseAJour
End Function

Public Function Composant() As ExtComposant
    Set Composant = pComposant
End Function

Public Function Piece() As ExtPiece
    Set Piece = pPiece
End Function

Public Function Assemblage() As ExtAssemblage
    Set Assemblage = pAssemblage
End Function

Public Function Dessin() As ExtDessin
    Set Dessin = pDessin
End Function

Public Function NouvelleRecherche() As ExtRechercher
    Set NouvelleRecherche = New ExtRechercher
    NouvelleRecherche.SetModele Me
End Function

Public Function Sw() As ExtSldWorks
    Set Sw = pSw
End Function

Public Sub Activer()
    pSw.swSw.ActivateDoc3 pSwModeleDoc.GetPathName, True, 0, Erreur
    Me.ZoomEtendu
    Me.Redessiner
End Sub

Public Sub Sauver()
    swModele.Save3 swSaveAsOptions_e.swSaveAsOptions_Silent, Erreur, Warning
End Sub

Public Sub Fermer()
    pSw.swSw.CloseDoc swModele.GetPathName
End Sub

Public Sub Redessiner()
    swModele.ActiveView.GraphicsRedraw Nothing
End Sub

Public Sub Reconstruire()
    swModele.EditRebuild3
End Sub

Public Sub ForcerAToutReconstruire()
    swModele.ForceRebuild3 False
End Sub

Public Sub ZoomEtendu()
    swModele.ViewZoomtofit2
End Sub

Public Function swModele() As ModelDoc2
    Set swModele = pSwModeleDoc
End Function

