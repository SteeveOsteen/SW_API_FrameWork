VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExtGestDeProprietes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private pSw                     As SldWorks.SldWorks
Private pModele             As ExtModele
Private pGestionnaire       As CustomPropertyManager
Private Const pNomClasse    As String = "ExtGestDeProprietes"

Private Sub Class_Initialize()
    Set pSw = Application.SldWorks
End Sub

Private Sub Class_Terminate()
    Set pSw = Nothing
    Set pModele = Nothing
    Set pGestionnaire = Nothing
End Sub

Public Function SetGestionnaire(ByRef Gestionnaire As CustomPropertyManager, ByRef Modele As ExtModele) As Boolean
    
    If Not (Gestionnaire Is Nothing) And Not (Modele Is Nothing) Then
        SetGestionnaire = True
        Set pModele = Modele
        Set pGestionnaire = Gestionnaire
        Exit Function
    End If
    
    Debug.Print pNomClasse & ".SetGestionnaire : le modele de base n'est pas valide"
    
End Function

Public Sub AjouterPropriete(ByVal NomPropriete As String, TypePropriete As swCustomInfoType_e, ByVal ValeurPropriete As String)
    
    pGestionnaire.Delete NomPropriete
    pGestionnaire.Add2 NomPropriete, TypePropriete, ValeurPropriete
    
End Sub

Public Function RecupererPropriete(ByVal NomPropriete As String) As String
    
    Dim Val     As String
    Dim ResVal  As String
    Dim base As String, vers As String, hot As String
    
    
    pSw.GetBuildNumbers2 base, vers, hot
    
    If base Like "sw2013*" Then
        pGestionnaire.Get4 NomPropriete, True, Val, ResVal
    Else
        pGestionnaire.Get2 NomPropriete, Val, ResVal
    End If
    
    RecupererPropriete = ResVal
    
End Function

Public Function SupprimerPropriete(ByVal NomPropriete As String)
    
    pGestionnaire.Delete NomPropriete
    
End Function
