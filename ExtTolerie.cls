VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExtTolerie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=======================================================================================
Option Explicit

Private pCorps                  As ExtCorps
Private pPiece                  As ExtPiece
Private Const pNomClasse        As String = "ExtTolerie"

Private Sub Class_Terminate()
    
    Set pCorps = Nothing
    Set pPiece = Nothing
End Sub

Public Function SetTolerie(ByRef Corps As ExtCorps, ByRef Piece As ExtPiece) As Boolean
    
    If Not (Corps Is Nothing) And Not (Piece Is Nothing) Then
        SetTolerie = True
        Set pCorps = Corps
        Set pPiece = Piece
        Exit Function
    End If
    
    Debug.Print pNomClasse & ".SetTolerie : le modele de base n'est pas valide"
    
End Function

Public Sub Deplier()
    FonctionDepliee.Activer True
End Sub

Public Sub Plier()
    FonctionDepliee.Desactiver False
End Sub

Public Function FonctionDepliee() As ExtFonction
        
    Dim vFonction           As Variant
    Dim Fonction            As Feature
    
    For Each vFonction In pCorps.swCorps.GetFeatures
        Set Fonction = vFonction
        If Fonction.GetTypeName2 = "FlatPattern" Then
            Set FonctionDepliee = New ExtFonction
            FonctionDepliee.SetFonction Fonction, pPiece
            Exit For
        End If
    Next vFonction
    
    Set Fonction = Nothing
    
End Function

Public Function FonctionTole() As ExtFonction
    
    If pCorps.Est(cTole) Then
        
        Dim vFonction           As Variant
        Dim Fonction            As Feature
        
        For Each vFonction In pCorps.swCorps.GetFeatures
            Set Fonction = vFonction
            If Fonction.GetTypeName2 = "SheetMetal" Then
                Set FonctionTole = New ExtFonction
                FonctionTole.SetFonction Fonction, pPiece
                Exit For
            End If
        Next vFonction
        
        Set Fonction = Nothing
    End If
    
End Function

Public Function CubeDeVisualisation() As ExtFonction
        Set CubeDeVisualisation = FonctionDepliee.ListeDesSousFonctions(CUBE_DE_VISUALISATION & "*").Item(1)
End Function


Private Function NomConfigDepliee() As String

    NomConfigDepliee = pPiece.Modele.GestDeConfigurations.ConfigurationActive.Nom & CONFIG_DEPLIEE & pCorps.Dossier.GestDeProprietes.RecupererPropriete(NO_DOSSIER)

End Function

Public Function RecupererConfigurationDepliee() As ExtConfiguration
    Dim ConfigDepliee           As ExtConfiguration
    
    For Each ConfigDepliee In pPiece.Modele.GestDeConfigurations.ListerLesConfigs(cDepliee)
        If ConfigDepliee.Nom = NomConfigDepliee Then
            Set RecupererConfigurationDepliee = ConfigDepliee
            Exit For
        End If
    Next ConfigDepliee
    
    Set ConfigDepliee = Nothing
    
End Function

Public Function CreerConfigurationDepliee() As Boolean
    
    Debug.Print , , "CreerConfigurationDepliee"
    
    Dim ConfigPliee                 As ExtConfiguration
    Dim ConfigDepliee               As ExtConfiguration
    Dim Fonction                    As ExtFonction
    Dim ListeDesFonctionsASupprimer As New Collection
    
    Set ConfigPliee = pPiece.Modele.GestDeConfigurations.ConfigurationActive
    
    If ConfigPliee.Est(cPliee) Then
        
                Debug.Print , , "Corps : "; pCorps.Nom; " -> Config à déplier : "; ConfigPliee.Nom
        CreerConfigurationDepliee = True
        
        Set ConfigDepliee = RecupererConfigurationDepliee
        
        If Not (ConfigDepliee Is Nothing) Then
            ConfigDepliee.Supprimer
        End If
        Set ConfigDepliee = ConfigPliee.AjouterUneConfigDerivee(NomConfigDepliee)
                Debug.Print , , "Nom de la configuration depliée : "; ConfigDepliee.Nom
        
        'On enregistre l'état des fonctions à supprimer du corps
        For Each Fonction In pPiece.ListeDesFonctions("S_*")
            Fonction.EnregistrerEtat
            ListeDesFonctionsASupprimer.Add Fonction
        Next Fonction
                
        ConfigDepliee.Activer
        Me.Deplier
        For Each Fonction In pPiece.ListeDesFonctions("S_*")
            Fonction.Desactiver
        Next Fonction
        
        ConfigPliee.Activer
        Me.Plier
        For Each Fonction In ListeDesFonctionsASupprimer
            Fonction.RestaurerEtat
        Next Fonction
        
    End If
    
    Set ListeDesFonctionsASupprimer = Nothing
    Set ConfigPliee = Nothing
    Set ConfigDepliee = Nothing
    Set Fonction = Nothing
    
End Function
