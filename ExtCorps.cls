VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExtCorps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=======================================================================================
Option Explicit

Private pSwCorps                As Body2
Private pPiece                  As ExtPiece
Private pTolerie                As ExtTolerie
Private Const pNomClasse        As String = "ExtCorps"

Private Sub Class_Terminate()
    Set pSwCorps = Nothing
    Set pPiece = Nothing
    Set pTolerie = Nothing
End Sub

Public Property Get Nom() As String
    Nom = swCorps.Name
End Property

Public Function SetCorps(ByRef Corps As Body2, ByRef Piece As ExtPiece) As Boolean
    
    If Not (Corps Is Nothing) And Not (Piece Is Nothing) Then
        SetCorps = True
        Set pPiece = Piece
        Set pSwCorps = Corps
        If Est(cTole) Then
            Set pTolerie = New ExtTolerie
            pTolerie.SetTolerie Me, pPiece
        End If
        Exit Function
    End If
    
    Debug.Print pNomClasse & ".SetCorps : le modele de base n'est pas valide"
    
End Function

Public Function Tolerie() As ExtTolerie
    If Est(cTole) Then
        Set Tolerie = pTolerie
    End If
End Function

Public Function PremiereFonction() As ExtFonction
    
    Dim vFonction           As Variant
    Dim Fonction            As Feature
    
    Set PremiereFonction = New ExtFonction
    vFonction = swCorps.GetFeatures
    Set Fonction = vFonction(0)
    PremiereFonction.SetFonction Fonction, pPiece
    
    Set Fonction = Nothing
    
End Function

Public Function ListeDesFonctions(Optional ByVal NomARechercher As String) As Collection
    
    Dim Fonction            As Feature
    Dim FonctionExt         As ExtFonction
    
    If NomARechercher = vbNullString Then NomARechercher = "*"
    
    Set ListeDesFonctions = New Collection
    
    Set Fonction = pPiece.Modele.swModele.FirstFeature
    
    Do Until Fonction Is Nothing
    
        If Fonction.Name Like NomARechercher Then
            Set FonctionExt = New ExtFonction
            FonctionExt.SetFonction Fonction, pPiece
            ListeDesFonctions.Add FonctionExt
        End If
        
        Set Fonction = Fonction.GetNextFeature
    Loop
    
    Set FonctionExt = Nothing
    Set Fonction = Nothing
    
End Function

Public Function Est(T As TypeCorps_e) As Boolean
    Est = TypeDeCorps And T
    Exit Function
End Function

Public Function TypeDeCorps() As TypeCorps_e
    Dim vFonction   As Variant
    Dim Fonction    As Feature
    
     For Each vFonction In swCorps.GetFeatures
        Set Fonction = vFonction
        Select Case Fonction.GetTypeName2
            Case "WeldMemberFeat"
                TypeDeCorps = cProfil
                Exit For
            Case "FlatPattern"
                TypeDeCorps = cTole
                Exit For
            Case Else
                TypeDeCorps = cAutre
         End Select
    Next vFonction
    
End Function

Public Function Dossier() As ExtDossier
    
    Dim lDossier        As ExtDossier
    Dim Corps           As ExtCorps
    Dim Sortir          As Boolean
    
    For Each lDossier In pPiece.ListeDesDossiers(cTousLesTypesDeCorps, True)
        For Each Corps In lDossier.ListeDesCorps
            If Corps.Nom = Me.Nom Then
                Set Dossier = lDossier
                Sortir = True
                Exit For
            End If
        Next Corps
        If Sortir Then Exit For
    Next lDossier
    
    Set lDossier = Nothing
    Set Corps = Nothing
    
End Function

Public Function swCorps() As Body2
    Set swCorps = pSwCorps
End Function

