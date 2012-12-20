VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExtDessin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=======================================================================================
Option Explicit

Private pModele             As ExtModele
Private pSwDessinDoc        As DrawingDoc
Private Const pNomClasse    As String = "ExtDessin"
Private Erreur As Long, Warning As Long

Private Sub Class_Terminate()
    Debug.Assert False
    Set pModele = Nothing
    Set pSwDessinDoc = Nothing
End Sub

Public Function Modele() As ExtModele
    Set Modele = pModele
End Function

Public Function SetModele(Modele As ExtModele) As Boolean
    
    If Not (Modele Is Nothing) Then
        If Modele.TypeDuModele = cDessin Then
            SetModele = True
            Set pModele = Modele
            Set pSwDessinDoc = Modele.swModele
            Exit Function
        End If
    End If
    
    Debug.Print pNomClasse & ".SetModele : le Modele de base n'est pas valide"
    
End Function

Public Function AjouterFeuille(NomFeuille As String) As ExtFeuille
    If FeuilleExiste(NomFeuille) Then
        Set AjouterFeuille = Feuille(NomFeuille)
    Else
        Set AjouterFeuille = New ExtFeuille
        swDessin.NewSheet3 NomFeuille, swDwgPaperSizes_e.swDwgPapersUserDefined, swDwgTemplates_e.swDwgTemplateNone, 1#, 1#, True, "", 10#, 10#, "Par défaut"
        AjouterFeuille.SetFeuille swDessin.Sheet(NomFeuille), Me
    End If
End Function

Public Function Feuille(ByVal NomFeuille As String) As ExtFeuille
    
    Set Feuille = New ExtFeuille
    
    Feuille.SetFeuille swDessin.Sheet(NomFeuille), Me
    
End Function

Public Function FeuilleActive() As ExtFeuille
    
    Set FeuilleActive = New ExtFeuille
    
    FeuilleActive.SetFeuille swDessin.GetCurrentSheet, Me
    
End Function

Public Function ListeDesFeuilles(Optional ByVal NomARechercher As String) As Collection
    
    Dim vNomFeuille As Variant
    Dim NomFeuille  As String
    
    Set ListeDesFeuilles = New Collection
    
    NomARechercher = NomARechercher & "*"
    
    For Each vNomFeuille In swDessin.GetSheetNames
        NomFeuille = vNomFeuille
        If NomFeuille Like NomARechercher Then
            ListeDesFeuilles.Add Feuille(vNomFeuille)
        End If
    Next vNomFeuille
    
End Function

Public Function FeuilleExiste(ByVal NomARechercher As String) As Boolean
    
    Dim vNomFeuille As Variant
    Dim NomFeuille  As String
    
    For Each vNomFeuille In swDessin.GetSheetNames
        NomFeuille = vNomFeuille
        If NomFeuille Like NomARechercher Then
            FeuilleExiste = True
        End If
    Next vNomFeuille
    
End Function

Public Function swDessin() As DrawingDoc
    Set swDessin = pSwDessinDoc
End Function
