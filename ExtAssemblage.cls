VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExtAssemblage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private pModele             As ExtModele
Private Const pNomClasse    As String = "ExtAssemblage"

Private Sub Class_Terminate()
    
    Set pModele = Nothing
End Sub

Public Function Modele() As ExtModele
    Set Modele = pModele
End Function

Public Function SetModele(Modele As ExtModele) As Boolean
    
    If Not (Modele Is Nothing) Then
        If Modele.Est(cAssemblage) Then
            SetModele = True
            Set pModele = Modele
            Exit Function
        End If
    End If
    
    Debug.Print pNomClasse & ".SetModele : le modele de base n'est pas valide"
    
End Function
