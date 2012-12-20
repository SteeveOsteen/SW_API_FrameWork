VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GestDeMiseAJourPiece"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=======================================================================================
Option Explicit

Private pPiece      As ExtPiece
Private Const pNomClasse As String = "GestDeMiseAJourPiece"

Private Sub Class_Terminate()
    Set pPiece = Nothing
End Sub

Public Function SetModele(ByRef Piece As ExtPiece) As Boolean
    
    If Not (Piece Is Nothing) Then
        Set pPiece = Piece
        SetModele = True
        Exit Function
    End If
    
    Debug.Print pNomClasse & ".SetModele : le modele de base n'est pas valide"
    
End Function

Public Sub MettreAJourLesNomsDeConfigs()
    
    On Error GoTo GestErreur
    
    Debug.Print , "MettreAJourLesNomsDeConfigs"
    
    Dim Configuration       As ExtConfiguration
    
    For Each Configuration In pPiece.Modele.GestDeConfigurations.ListerLesConfigs(cDepliee)
        
        Configuration.GestDeProprietes.AjouterPropriete "NoConfig", swCustomInfoText, Configuration.ConfigurationParent.Nom
        
    Next Configuration
    
    Set Configuration = Nothing
    
    Exit Sub
    
GestErreur:
    Debug.Print "Erreur [" & pNomClasse & ".MettreAJourLesNomsDeConfigs] : " & Err.Number & " ->  " & Err.Description
    Resume Next
    
End Sub

Public Sub MettreAJourLaListeDesPiecesSoudees()
    
    On Error GoTo GestErreur
    
    Debug.Print , "MettreAJourLaListeDesPiecesSoudees"
    
    Dim Dossier             As ExtDossier
    Dim NomElement          As String
    Dim NoDossier           As Integer
    
    NoDossier = 1
    
    For Each Dossier In pPiece.ListeDesDossiers(cTousLesTypesDeCorps, True)
        
        Dossier.GestDeProprietes.SupprimerPropriete NOM_ELEMENT
        Dossier.GestDeProprietes.SupprimerPropriete NO_DOSSIER
        Dossier.GestDeProprietes.SupprimerPropriete "NoCorps"
        
        NomElement = vbNullString
        If Not (Dossier.Nom Like (ARTICLE_LISTE_DES_PIECES_SOUDEES & "*")) Then
            NomElement = Trim(Dossier.Nom)
        End If
        
        Dossier.GestDeProprietes.AjouterPropriete NOM_ELEMENT, swCustomInfoText, NomElement
        Dossier.GestDeProprietes.AjouterPropriete NO_DOSSIER, swCustomInfoText, NoDossier
        
        NoDossier = NoDossier + 1
        
    Next Dossier
    
    Set Dossier = Nothing
    
    Exit Sub
    
GestErreur:
    Debug.Print "Erreur [" & pNomClasse & ".MettreAJourLaListeDesPiecesSoudees] : " & Err.Number & " ->  " & Err.Description
    Resume Next
    
End Sub

