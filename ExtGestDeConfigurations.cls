VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExtGestDeConfigurations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private pModele                 As ExtModele
Private Const pNomClasse        As String = "ExtGestDeConfigurations"

Public Function SetModele(ByRef Modele As ExtModele) As Boolean
    
    If Not (Modele Is Nothing) Then
        SetModele = True
        Set pModele = Modele
        Exit Function
    End If
    
    Debug.Print pNomClasse & ".SetModele : " & Err.Number & " ->  " & Err.Description
    
End Function

Public Function ConfigurationActive() As ExtConfiguration

    Set ConfigurationActive = New ExtConfiguration
    
    ConfigurationActive.SetConfiguration pModele.swModele.ConfigurationManager.ActiveConfiguration, pModele
    
End Function

Public Function Configuration(ByVal NomConfig As String) As ExtConfiguration
    Dim Config          As Configuration
    
    Set Config = pModele.swModele.GetConfigurationByName(NomConfig)
    
    If Not (Config Is Nothing) Then
        Set Configuration = New ExtConfiguration
        Configuration.SetConfiguration Config, pModele
    End If
    
    Set Config = Nothing
End Function

Public Function ListerLesConfigs(Optional ByVal TypeConfig As TypeConfig_e = TypeConfig_e.cToutesLesTypesDeConfig, Optional NomConfigDeBase As String) As Collection
    
    Dim swConfig            As Configuration
    Dim Configuration       As ExtConfiguration
    Dim ConfigurationParent As ExtConfiguration
    Dim vNomConfig          As Variant
    Dim NomConfig           As String
    Dim NomConfigParent     As String
    
    Set ListerLesConfigs = New Collection
    
    If NomConfigDeBase = vbNullString Then NomConfigDeBase = "*"
    
    For Each vNomConfig In pModele.swModele.GetConfigurationNames
        
        NomConfig = vNomConfig
        Set swConfig = pModele.swModele.GetConfigurationByName(NomConfig)
        
        Set Configuration = New ExtConfiguration
        Configuration.SetConfiguration swConfig, pModele
        
        Set ConfigurationParent = Configuration.ConfigurationParent
        If Not (ConfigurationParent Is Nothing) Then
            NomConfigParent = ConfigurationParent.Nom
        End If
        
        If Configuration.Est(TypeConfig) And (NomConfigParent Like NomConfigDeBase) Then
            ListerLesConfigs.Add Configuration
        End If
        
        Set Configuration = Nothing
        
    Next vNomConfig
    
End Function

Public Function AjouterUneConfigDeBase(ByVal NomConfig As String) As ExtConfiguration
    
    Set AjouterUneConfigDeBase = New ExtConfiguration
    
    AjouterUneConfigDeBase.SetConfiguration pModele.swModele.ConfigurationManager.AddConfiguration(NomConfig, NomConfig, "", 0, "", ""), pModele
    
End Function

Public Sub SupprimerLesConfigurationsDeplies(Optional NomConfigDeBase As String)
    
    Dim Config              As ExtConfiguration
    
    Debug.Print , "SupprimerLesConfigurationsDeplies"
    
    For Each Config In Me.ListerLesConfigs(cDepliee, NomConfigDeBase)
        Debug.Print , , Config.Nom
        Config.Supprimer
    Next Config
    
    Set Config = Nothing
    
End Sub

