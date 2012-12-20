VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExtSldWorks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=======================================================================================
Option Explicit

Private pSw                 As SldWorks.SldWorks
Private Const pNomClasse    As String = "ExtSldWorks"
Private Erreur As Long, Warning As Long

Private Sub Class_Initialize()
    Set pSw = Application.SldWorks
    Dim VersionDeBase   As String
    Dim VersionCourante As String
    Dim Hotfixe         As String
    pSw.GetBuildNumbers2 VersionDeBase, VersionCourante, Hotfixe
    Debug.Print "    "
    Debug.Print "================================================================================================"
    Debug.Print "SOLIDWORKS"
    Debug.Print "Version de base : "; VersionDeBase, "Version courante : "; VersionCourante, "Hotfixe : "; Hotfixe
    Debug.Print "------------------------------------------------------------------------------------------------"
End Sub

Private Sub Class_Terminate()
    Debug.Assert False
    Set pSw = Nothing
    Debug.Print "    "
    Debug.Print "================================================================================================"
End Sub

Public Function TypeDuModeleActif() As TypeFichier_e
    
    Dim Modele As ModelDoc2
    Dim GestDeFichier As New SysGestDeFichiers
    
    Set Modele = pSw.ActiveDoc
    GestDeFichier.Chemin = Modele.GetPathName
    
    TypeDuModeleActif = GestDeFichier.TypeDuFichier
    
    Set GestDeFichier = Nothing
    Set Modele = Nothing
    
End Function

Public Function Modele(Optional ByVal Chemin As String) As ExtModele
    
    Dim ModeleActif As ModelDoc2
    Dim ModeleDoc   As ModelDoc2
    Dim GestDeFichier As New SysGestDeFichiers
    
    Set Modele = New ExtModele
    
    If Chemin = vbNullString Then
        Set ModeleActif = pSw.ActiveDoc
        Modele.SetModele ModeleActif, Me
    Else
        GestDeFichier.Chemin = Chemin
        Select Case GestDeFichier.TypeDuFichier
            Case cAssemblage
                Set ModeleDoc = Ouvrir(Chemin, swDocASSEMBLY)
                Modele.SetModele ModeleDoc, Me
            Case cPiece
                Set ModeleDoc = Ouvrir(Chemin, swDocPART)
                Modele.SetModele ModeleDoc, Me
            Case cDessin
                Set ModeleDoc = Ouvrir(Chemin, swDocDRAWING)
                Modele.SetModele ModeleDoc, Me
        End Select
    End If
    
    Set ModeleActif = Nothing
    
End Function

Public Function CreerDessin(ByVal Dossier As String, ByVal NomDuDessin As String, ByVal NomDuModele As String) As ExtModele
    
    Dim CheminModeleDeDessin As String
    Dim Modele As ModelDoc2
    
    Set CreerDessin = New ExtModele
    
    CheminModeleDeDessin = pSw.GetUserPreferenceStringValue(swFileLocationsDocumentTemplates) & "\" & NomDuModele & ".drwdot"
    
    Set Modele = pSw.NewDocument(CheminModeleDeDessin, swDwgPaperSizes_e.swDwgPapersUserDefined, 10#, 10#)
    
    Modele.Extension.SaveAs Dossier & "\" & NomDuDessin & ".SLDDRW", swSaveAsVersion_e.swSaveAsCurrentVersion, swSaveAsOptions_e.swSaveAsOptions_Silent, Nothing, Erreur, Warning
    
    CreerDessin.SetModele Modele, Me
    
    Set Modele = Nothing
    
End Function

Private Function Ouvrir(ByVal Chemin As String, TypeDoc As swDocumentTypes_e) As ModelDoc2
    
    Dim vDoc        As Variant
    Dim Doc         As ModelDoc2
    
    'On liste les documents ouverts
    For Each vDoc In pSw.GetDocuments
        Set Doc = vDoc
        'Si les chemin correspondent c'est bon
        If Doc.GetPathName = Chemin Then
            Set Ouvrir = Doc
        End If
    Next vDoc
    
    'Si le document n'est pas ouvert, on le fait
    If Ouvrir Is Nothing Then
        Set Ouvrir = pSw.OpenDoc6(Chemin, TypeDoc, swOpenDocOptions_e.swOpenDocOptions_Silent, "", Erreur, Warning)
    End If
    
    Set Doc = Nothing
    
End Function

Public Function RecupererCle(ByVal NomCle As swUserPreferenceStringValue_e) As String
    
    On Error GoTo GestErreur
    
    RecupererCle = pSw.GetUserPreferenceStringValue(NomCle)
    
    Exit Function
    
GestErreur:
    Debug.Print "Erreur [" & pNomClasse & ".RecupererCle] : " & Err.Number & " ->  " & Err.Description
    Resume Next
End Function

Public Function swSw() As SldWorks.SldWorks
    Set swSw = pSw
End Function
