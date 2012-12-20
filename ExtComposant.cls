VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExtComposant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=======================================================================================
Option Explicit

Private pNb                 As Integer
Private pSwComposant        As Component2
Private pModele             As ExtModele
Private pConfiguration      As New ExtConfiguration
Private Const pNomClasse    As String = "ExtComposant"

Private Sub Class_Terminate()
    
    Set pSwComposant = Nothing
    Set pModele = Nothing
    Set pConfiguration = Nothing
End Sub

Public Function SetComposant(ByRef Composant As Component2, ByRef ModeleExt As ExtModele) As Boolean
    
    If Not (Composant Is Nothing) And Not (ModeleExt Is Nothing) Then
        SetComposant = True
        Set pSwComposant = Composant
        Set pModele = ModeleExt
        pNb = 1
        
        Dim Config          As Configuration
        
        'Quand le composant est le composant racine [ModeleDoc.ConfigurationManager.ActiveConfiguration.GetRootComponent3(True)] ,
        'le nom de la configuration referencée n'est pas dispo, on passe donc par le ModelDoc
        'pour recuperer la configuration active
        If Composant.ReferencedConfiguration = vbNullString Then
            Set Config = pModele.swModele.GetActiveConfiguration
        Else
            Set Config = pModele.swModele.GetConfigurationByName(Composant.ReferencedConfiguration)
        End If
        
        pConfiguration.SetConfiguration Config, pModele
        
        Set Config = Nothing
        
        Exit Function
    End If
    
    Debug.Print pNomClasse & ".SetComposant : le composant de base n'est pas valide"
    
End Function

Public Property Let Nb(N As Integer)
    pNb = N
End Property

Public Property Get Nb() As Integer
    Nb = pNb
End Property

Public Property Get EstExclu() As Boolean
    If Not (swComposant.ExcludeFromBOM = False) Then
        EstExclu = True
    End If
End Property

Public Property Let EstExclu(T As Boolean)
    swComposant.ExcludeFromBOM = T
End Property

Public Property Get EstSupprime() As Boolean
    If Not (swComposant.IsSuppressed = False) Then
        EstSupprime = True
    End If
End Property

Public Function Modele() As ExtModele
    Set Modele = pModele
End Function

Public Function Configuration() As ExtConfiguration
    Set Configuration = pConfiguration
End Function

Public Function ListedesComposantsEnfants(Optional PrendreEnCompteSupprime As Boolean = False) As Collection
    
    Dim vComposant  As Variant
    Dim Composant   As Component2
    Dim CompExt     As ExtComposant
    Dim ModeleExt   As ExtModele
    
    Set ListedesComposantsEnfants = New Collection
    
    For Each vComposant In swComposant.GetChildren
        
        Set Composant = vComposant
        
        If Not (Composant.IsSuppressed = False) Imp PrendreEnCompteSupprime Then
            Set ModeleExt = New ExtModele
            Set CompExt = New ExtComposant
            
            'Si on utilise Composant.GetModelDoc2 dans le cas ou celui ci est supprimé, on recupère un NULL
            'On passe donc par Composant.GetPathName
            'Peut peut-être poser des pb dans de gros assemblage, puisque cette méthode ouvre le fichier en arrière plan
            Set ModeleExt = pModele.Sw.Modele(Composant.GetPathName)
            
            CompExt.SetComposant Composant, ModeleExt
            ModeleExt.SetComposant CompExt
            
            ListedesComposantsEnfants.Add CompExt
        End If
    
    Next vComposant
    
    Set Composant = Nothing
    Set CompExt = Nothing
    Set ModeleExt = Nothing

End Function

Public Function swComposant() As Component2
    Set swComposant = pSwComposant
End Function
