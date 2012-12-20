VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExtRechercher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private pModele                 As ExtModele
Private pCollPieces             As Collection
Private pPrendreEnCompteConfig  As Boolean
Private pPrendreEnCompteExclus  As Boolean
Private Const pNomClasse        As String = "ExtRechercher"
Private Erreur As Long, Warning As Long

Private Sub Class_Initialize()
    
    pPrendreEnCompteConfig = True
    pPrendreEnCompteExclus = False
End Sub

Private Sub Class_Terminate()
    
    Set pModele = Nothing
    Set pCollPieces = Nothing
End Sub

Public Property Get PrendreEnCompteConfig() As Boolean
    PrendreEnCompteConfig = pPrendreEnCompteConfig
End Property

Public Property Let PrendreEnCompteConfig(T As Boolean)
    pPrendreEnCompteConfig = T
End Property

Public Property Get PrendreEnCompteExclus() As Boolean
    PrendreEnCompteConfig = pPrendreEnCompteExclus
End Property

Public Property Let PrendreEnCompteExclus(T As Boolean)
    pPrendreEnCompteExclus = T
End Property

Public Function Modele() As ExtModele
    Set Modele = pModele
End Function

Public Function SetModele(Modele As ExtModele) As Boolean
    
    If Not (Modele Is Nothing) Then
        SetModele = True
        Set pModele = Modele
        Exit Function
    End If
    
    Debug.Print pNomClasse & ".SetModele : le Modele de base n'est pas valide"
    
End Function

Public Function NomCle(ByRef Composant As ExtComposant) As String
    NomCle = Composant.Modele.Fichier.NomDuFichier
    If pPrendreEnCompteConfig Then
        NomCle = NomCle & " " & Composant.Configuration.Nom
    End If
End Function

Private Sub ListerLesComposants(ByRef ComposantRacine As ExtComposant, TypeComposant As TypeFichier_e, NomComposant As String)
    
    On Error GoTo GestErreur
    
    Dim ComposantListe      As ExtComposant
    Dim Composant           As ExtComposant
    Dim Cle                 As String
    
    For Each ComposantListe In ComposantRacine.ListedesComposantsEnfants
        
        If (ComposantListe.EstExclu Imp pPrendreEnCompteExclus) Then
            
            If ComposantListe.Modele.Est(TypeComposant) And (ComposantListe.Modele.Fichier.NomDuFichier Like NomComposant) Then
                
                Cle = NomCle(ComposantListe)
                
                If CleExiste(pCollPieces, Cle) Then
                    Set Composant = pCollPieces.Item(Cle)
                    Composant.Nb = Composant.Nb + 1
                Else
                    Set Composant = New ExtComposant
                    Set Composant = ComposantListe
                    Composant.Nb = 1
                    pCollPieces.Add Composant, Cle
                End If
                    
            End If
            
            If ComposantListe.Modele.Est(cAssemblage) Then
                ListerLesComposants ComposantListe, TypeComposant, NomComposant
            End If
            
        End If
        
    Next ComposantListe
    
    Exit Sub
    
GestErreur:
    Debug.Print "Erreur [" & pNomClasse & ".ListerComposants] : " & Err.Number & " ->  " & Err.Description
    Resume Next
    
End Sub
 
Public Function ListeDesComposants(TypeComposant As TypeFichier_e, Optional NomComposant As String) As Collection
    
    On Error GoTo GestErreur
    
    Dim Composant   As New ExtComposant
    
    Set ListeDesComposants = New Collection
    Set pCollPieces = New Collection
    
    Select Case pModele.TypeDuModele
        Case cAssemblage
            If pModele.swModele.GetComponentCount(False) Then
                Composant.SetComposant pModele.swModele.ConfigurationManager.ActiveConfiguration.GetRootComponent3(True), pModele
                NomComposant = NomComposant & "*"
                ListerLesComposants Composant, TypeComposant, NomComposant
            Else
                pCollPieces.Add pModele.Composant
            End If
        Case cPiece
            pCollPieces.Add pModele.Composant
    End Select
    
    Set ListeDesComposants = pCollPieces
    Set pCollPieces = Nothing
    
    Exit Function
    
GestErreur:
    Debug.Print "Erreur [" & pNomClasse & ".ListeDesComposants] : " & Err.Number & " ->  " & Err.Description
    Resume Next
    
End Function

Private Function CleExiste(Coll As Collection, Cle As String) As Boolean
    
    On Error Resume Next
    Coll.Item Cle
    CleExiste = (Err.Number = 0)
    Err.Clear
    
End Function
