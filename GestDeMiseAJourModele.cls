VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GestDeMiseAJourModele"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private pModele     As ExtModele
Private Const pNomClasse As String = "GestDeMiseAJourModele"

Private Sub Class_Terminate()
    Set pModele = Nothing
End Sub

Public Function SetModele(ByRef Modele As ExtModele) As Boolean
    
    If Not (Modele Is Nothing) Then
        Set pModele = Modele
        SetModele = True
        Exit Function
    End If
    
    Debug.Print pNomClasse & ".SetModele : le modele de base n'est pas valide"
    
End Function

Sub ImporterLesInfosClient(NomFichier As String)
    
    Debug.Print pModele.Fichier.NomDuFichier
    Debug.Print , "ImporterLesInfosClient"
    
    Dim GestFichiers        As New SysGestDeFichiers
    Dim Propriete           As ExtPropriete
    
    GestFichiers.Chemin = pModele.Fichier.Chemin
    For Each Propriete In GestFichiers.RecupererPropriete(NomFichier)
        pModele.GestDeProprietes.AjouterPropriete Propriete.Nom, swCustomInfoText, Propriete.Valeur
    Next Propriete
    
    Set GestFichiers = Nothing
    Set Propriete = Nothing
End Sub
