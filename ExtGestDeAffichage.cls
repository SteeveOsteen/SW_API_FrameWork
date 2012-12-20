VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExtGestDeAffichage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=======================================================================================
Option Explicit

Private pConfig             As ExtConfiguration
Private Const pNomClasse    As String = "ExtConfiguration"

Private Sub Class_Terminate()
    Set pConfig = Nothing
End Sub

Public Function SetConfiguration(ByRef Config As ExtConfiguration) As Boolean
    
    If Not (Config Is Nothing) Then
        SetConfiguration = True
        Set pConfig = Config
        
        Exit Function
    End If
    
    Debug.Assert False
    Debug.Print pNomClasse & ".SetConfiguration : le Modele de base n'est pas valide"
    
End Function

Public Function ListeDesAffichages() As Collection
    Dim vNomAff As Variant
    Dim NomAff  As String
    
    Set ListeDesAffichages = New Collection
    
    For Each vNomAff In pConfig.swConfiguration.GetDisplayStates
        NomAff = vNomAff
        ListeDesAffichages.Add NomAff
    Next vNomAff
    
End Function

Public Function SupprimerAffichage(ByVal NomAffichage As String) As Boolean
    SupprimerAffichage = pConfig.swConfiguration.DeleteDisplayState(NomAffichage)
    If Not (SupprimerAffichage = False) Then
        SupprimerAffichage = True
    End If
End Function

Public Function AjouterAffichage(ByVal NomAffichage As String) As Boolean
    AjouterAffichage = pConfig.swConfiguration.CreateDisplayState(NomAffichage)
    If Not (AjouterAffichage = False) Then
        AjouterAffichage = True
    End If
End Function

Public Function RenommerAffichage(ByVal AncienNom As String, ByVal NouveauNom As String) As Boolean
    RenommerAffichage = pConfig.swConfiguration.RenameDisplayState(AncienNom, NouveauNom)
    If Not (RenommerAffichage = False) Then
        RenommerAffichage = True
    End If
End Function
