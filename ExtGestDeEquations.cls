VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExtGestDeEquations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private pModele             As ExtModele
Private pGestionnaire       As EquationMgr
Private Const pNomClasse    As String = "ExtGestDeEquations"

Private Sub Class_Terminate()
    Set pModele = Nothing
    Set pGestionnaire = Nothing
End Sub

Public Function SetGestionnaire(ByRef Gestionnaire As EquationMgr, ByRef Modele As ExtModele) As Boolean
    
    If Not (Gestionnaire Is Nothing) And Not (Modele Is Nothing) Then
        SetGestionnaire = True
        Set pModele = Modele
        Set pGestionnaire = Gestionnaire
        Exit Function
    End If
    
    Debug.Print pNomClasse & ".SetGestionnaire : le modele de base n'est pas valide"
    
End Function

Public Sub ToutRecalculer()
    pGestionnaire.EvaluateAll
End Sub
