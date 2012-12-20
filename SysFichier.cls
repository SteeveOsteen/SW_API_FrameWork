VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SysFichier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private pFs                 As Scripting.FileSystemObject
Private pChemin             As String
Private Const pNomClasse    As String = "SysFichier"

Private Sub Class_Initialize()
    Set pFs = CreateObject("Scripting.FileSystemObject")
End Sub

Private Sub Class_Terminate()
    Set pFs = Nothing
End Sub

Public Property Get Chemin() As String
    Chemin = pChemin
End Property

Public Property Let Chemin(ByVal s As String)
    pChemin = s
End Property

Public Property Get NomDuFichier(Optional SansExt As Boolean = False) As String
    
    NomDuFichier = pFs.GetFileName(pChemin)
    
    If SansExt Then
        NomDuFichier = pFs.GetBaseName(pChemin)
    End If

End Property

Public Property Get NomDuDossier() As String
    
    NomDuDossier = pFs.GetParentFolderName(pChemin)

End Property

