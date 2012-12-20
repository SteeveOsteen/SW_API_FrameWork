VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SysGestDeFichiers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=======================================================================================
Option Explicit

Private pFs                 As Scripting.FileSystemObject
Private pChemin             As String
Private Const pNomClasse    As String = "SysGestDeFichiers"

Private Sub Class_Initialize()
    Set pFs = CreateObject("Scripting.FileSystemObject")
End Sub

Private Sub Class_Terminate()
    Set pFs = Nothing
End Sub

Public Property Get Chemin() As String
    Chemin = pChemin
End Property

Public Property Let Chemin(ByVal Chaine As String)
    pChemin = Chaine
End Property

Public Function Fichier() As SysFichier
    
    On Error GoTo GestErreur
    
    Set Fichier = New SysFichier
    
    Fichier.Chemin = pChemin
    
    Exit Function
    
GestErreur:
    Debug.Print "Erreur [Fichier] : " & Err.Number & " ->  " & Err.Description
    Resume Next
    
End Function

Public Function TypeDuFichier() As TypeFichier_e
    
    TypeDuFichier = swDocNONE
    
    If pChemin Like "*.SLDASM" Then
        TypeDuFichier = cAssemblage
    ElseIf pChemin Like "*.SLDPRT" Then
        TypeDuFichier = cPiece
    ElseIf pChemin Like "*.SLDDRW" Then
        TypeDuFichier = cDessin
    End If

End Function

Public Function FichierExiste() As Boolean
    
    If pFs.FileExists(pChemin) Then
        FichierExiste = True
    End If
    
End Function
Public Function DossierExiste() As Boolean
    
    If pFs.FolderExists(pChemin) Then
        DossierExiste = True
    End If
    
End Function

Public Function CreerDossier(Dossier As String) As String
    
    Dim CheminDossier As String
    CheminDossier = pFs.GetParentFolderName(pChemin) & "\" & Dossier
    
    If Not pFs.FolderExists(CheminDossier) Then
        pFs.CreateFolder CheminDossier
    End If
    
    CreerDossier = CheminDossier
    
End Function

Public Function RecupererPropriete(NomDuFichier As String) As Collection
    
    On Error GoTo GestErreur
    
    Dim Fichier             As Scripting.File
    Dim Texte               As Scripting.TextStream
    Dim Ligne               As String
    Dim Tmp()               As String
    Dim Propriete           As ExtPropriete
    
    Set RecupererPropriete = New Collection
    Set Fichier = pFs.GetFile(Me.Fichier.NomDuDossier & "\" & NomDuFichier)
    
    If Fichier Is Nothing Then Exit Function
    
    Set Texte = Fichier.OpenAsTextStream(ForReading)
    
    While Not Texte.AtEndOfStream
        Ligne = Texte.ReadLine
        If Not (Ligne = vbNullString) Then
            Tmp = Split(Ligne, ":")
            Set Propriete = New ExtPropriete
            Propriete.Nom = Trim(Tmp(0))
            Propriete.Valeur = Trim(Tmp(1))
            RecupererPropriete.Add Propriete
        End If
    Wend
    
    Exit Function
    
GestErreur:
    Debug.Print Me.Fichier.NomDuDossier & "\" & NomDuFichier
    Debug.Print "Erreur [" & pNomClasse & ".RecupererPropriete] : " & Err.Number & " ->  " & Err.Description
    Resume Next

End Function
