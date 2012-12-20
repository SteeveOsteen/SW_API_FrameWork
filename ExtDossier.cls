VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExtDossier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private pSwDossier              As BodyFolder
Private pNomDossier             As String
Private pPiece                  As ExtPiece
Private pGestDeProprietes       As New ExtGestDeProprietes
Private Const pNomClasse        As String = "ExtDossier"

Private Sub Class_Terminate()
    
    Set pPiece = Nothing
    Set pSwDossier = Nothing
    Set pGestDeProprietes = Nothing
End Sub

Public Function GestDeProprietes() As ExtGestDeProprietes
    Set GestDeProprietes = pGestDeProprietes
End Function


Public Function SetDossier(ByRef Dossier As BodyFolder, ByRef Piece As ExtPiece) As Boolean
    
    If Not (Dossier Is Nothing) And Not (Piece Is Nothing) Then
        Set pSwDossier = Dossier
        pNomDossier = Dossier.GetFeature.Name
        Set pPiece = Piece
        If Dossier.GetBodyCount Then
            SetDossier = True
            pGestDeProprietes.SetGestionnaire Dossier.GetFeature.CustomPropertyManager, pPiece.Modele
            Exit Function
        Else
            Set pPiece = Nothing
            Set pSwDossier = Nothing
            pNomDossier = vbNullString
        End If
        Exit Function
    End If
    
    Debug.Print pNomClasse & ".SetDossier : le dossier n'est pas valide"
    
End Function

Public Property Get Nom(Optional Nettoyer As Boolean = False) As String
    
    Nom = swDossier.GetFeature.Name
    
    If Nettoyer Then Nom = Trim(Nom)
    
End Property

Public Function Exclu() As Boolean
    
    On Error GoTo GestErreur
    
    Exclu = True
    If swDossier.GetFeature.ExcludeFromCutList = False Then Exclu = False
    
    Exit Function
    
GestErreur:
    Debug.Print "Erreur [" & pNomClasse & ".Exclu] : " & Err.Number & " ->  " & Err.Description
    Resume Next
    
End Function

Public Function PremierCorps() As ExtCorps
    
    Set PremierCorps = New ExtCorps
    
    Dim vCorps  As Variant
    Dim Corps   As Body2
    Dim Dossier As BodyFolder
    
    Set Dossier = swDossier
    
    If Dossier.GetBodyCount Then
        
        vCorps = Dossier.GetBodies
        Set Corps = vCorps(0)
        PremierCorps.SetCorps Corps, pPiece
        
    End If
    
    Exit Function
    
    Set Corps = Nothing
    Set Dossier = Nothing
    
End Function

Public Function ListeDesCorps(Optional ByVal NomARechercher As String) As Collection
    Set ListeDesCorps = New Collection
    
    Dim vCorps      As Variant
    Dim Corps       As Body2
    Dim CorpsExt    As ExtCorps
    
    If NomARechercher = vbNullString Then NomARechercher = "*"
    
    For Each vCorps In swDossier.GetBodies
        Set Corps = vCorps
        If Corps.Name Like NomARechercher Then
            Set CorpsExt = New ExtCorps
            CorpsExt.SetCorps Corps, pPiece
            ListeDesCorps.Add CorpsExt
            Set CorpsExt = Nothing
        End If
    Next vCorps
    
    Set Corps = Nothing
    Set CorpsExt = Nothing
    
End Function

Public Function Est(T As TypeCorps_e) As Boolean
    
    Est = TypeDeCorps And T
    
End Function

Private Function TypeDeCorps() As TypeCorps_e
    
    TypeDeCorps = PremierCorps.TypeDeCorps
    
End Function

Public Function swDossier() As BodyFolder
    
    Set swDossier = pSwDossier
    
End Function

Private Function ListeDesPiecesSoudees() As Feature
    Dim PiecesSoudees   As Feature
    Dim DossierSw       As BodyFolder
    
    'Recherche la liste des corps
    Set PiecesSoudees = pPiece.Modele.swModele.FirstFeature
    
    Do Until PiecesSoudees Is Nothing
    
        If PiecesSoudees.GetTypeName2 = "SolidBodyFolder" Then
        
            Set ListeDesPiecesSoudees = PiecesSoudees
            Set DossierSw = PiecesSoudees.GetSpecificFeature2
            DossierSw.UpdateCutList
            Exit Do
        End If
        Set PiecesSoudees = PiecesSoudees.GetNextFeature
    Loop
    
    Set PiecesSoudees = Nothing
    Set DossierSw = Nothing

End Function

