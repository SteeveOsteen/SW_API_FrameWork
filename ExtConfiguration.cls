VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExtConfiguration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=======================================================================================
Option Explicit

Private pSwConfig           As Configuration
Private pModele             As ExtModele
Private pGestDeProprietes   As New ExtGestDeProprietes
Private pGestDeAffichage    As New ExtGestDeAffichage
Private Const pNomClasse    As String = "ExtConfiguration"

Private Sub Class_Terminate()
    Set pGestDeProprietes = Nothing
    Set pGestDeAffichage = Nothing
    Set pSwConfig = Nothing
End Sub

Public Property Get Nom() As String
    Nom = swConfiguration.Name
End Property

Public Property Let Nom(s As String)
    swConfiguration.Name = s
End Property

Public Function SetConfiguration(ByRef Config As Configuration, ByRef Modele As ExtModele) As Boolean
    
    If Not (Config Is Nothing) And Not (Modele Is Nothing) Then
        SetConfiguration = True
        Set pSwConfig = Config
        Set pModele = Modele
        
        pGestDeProprietes.SetGestionnaire swConfiguration.CustomPropertyManager, pModele
        pGestDeAffichage.SetConfiguration Me
        
        Exit Function
    End If
    
    Debug.Assert False
    Debug.Print pNomClasse & ".SetConfiguration : le Modele de base n'est pas valide"
    
End Function

Public Function GestDeProprietes() As ExtGestDeProprietes
    Set GestDeProprietes = pGestDeProprietes
End Function

Public Function GestDeAffichage() As ExtGestDeAffichage
    Set GestDeAffichage = pGestDeAffichage
End Function

Public Function Activer() As Boolean
    
'    If Not (Configuration.IsDirty = False) Or Not (Configuration.NeedsRebuild = False) Then
'        pModele.Reconstruire
'        pModele.Redessiner
'    End If
    
    Activer = pModele.swModele.ShowConfiguration2(Nom)
    
End Function

Public Function Supprimer() As Boolean
    
    Supprimer = pModele.swModele.DeleteConfiguration2(Nom)
End Function

Public Function AjouterUneConfigDerivee(ByVal NomConfigDerivee As String) As ExtConfiguration
    Set AjouterUneConfigDerivee = New ExtConfiguration
    AjouterUneConfigDerivee.SetConfiguration pModele.swModele.ConfigurationManager.AddConfiguration(NomConfigDerivee, NomConfigDerivee, "", 0, Nom, ""), pModele
End Function

Public Function ConfigurationParent() As ExtConfiguration
    
    If Est(cDerivee) Then
        Set ConfigurationParent = New ExtConfiguration
        ConfigurationParent.SetConfiguration swConfiguration.GetParent, pModele
    End If
    
End Function

Public Function ConfigurationRacine() As ExtConfiguration
    
    Dim Config  As ExtConfiguration
    
    If Est(cDerivee) Then
        Set Config = Me
        
        Do Until Config.Est(cDeBase)
            Set Config = Config.ConfigurationParent
        Loop
        
        Set ConfigurationRacine = Config
        Set Config = Nothing
    Else
        Set ConfigurationRacine = Me
    End If
    
End Function

Public Function CorpsDepliee() As ExtCorps
    
    If Est(cDepliee) Then
        Dim Dossier As ExtDossier
        Dim Corps   As ExtCorps
        
        For Each Dossier In pModele.Piece.ListeDesDossiers(cTole, True)
            'En 2012 le corps déplié s'appelle "Etat déplié#"
            'En 2013 le corps déplié s'appelle "Etat déplié"
            'Pour avoir une compatiblité 2012-2013, je filtre sur "Etat déplié*"
            'A voir s'il n'y a pas de pb avec un corps nommé "Etat déplié" d'origine
            
            For Each Corps In Dossier.ListeDesCorps(NOM_CORPS_DEPLIEE & "*")
                Set CorpsDepliee = Corps
            Next Corps
        Next Dossier
        
        Set Corps = Nothing
        Set Dossier = Nothing
        
    End If
    
End Function

Public Function Est(T As TypeConfig_e) As Boolean
    Est = TypeConfig And T
End Function

Public Function TypeConfig() As TypeConfig_e
    Dim Configuration   As Configuration
    
    Set Configuration = swConfiguration
    
    If Configuration.Name Like ("*" & CONFIG_DEPLIEE & "*") Then
        TypeConfig = TypeConfig_e.cDepliee
    ElseIf Configuration.Name Like (CONFIG_PLIEE & "*") Then
        TypeConfig = TypeConfig_e.cPliee
    End If
    
    If Configuration.IsDerived Then
        TypeConfig = TypeConfig + TypeConfig_e.cDerivee
    Else
        TypeConfig = TypeConfig + TypeConfig_e.cDeBase
    End If
    
    Set Configuration = Nothing
    
End Function

Public Function swConfiguration() As Configuration
    Set swConfiguration = pSwConfig
End Function
