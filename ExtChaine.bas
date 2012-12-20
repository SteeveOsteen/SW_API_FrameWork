Attribute VB_Name = "ExtChaine"
Public Function EchapperCaractereLike(ByVal Chaine As String) As String
    EchapperCaractere = Chaine
    EchapperCaractere = Replace(EchapperCaractere, "*", "[*]")
    EchapperCaractere = Replace(EchapperCaractere, "?", "[?]")
    EchapperCaractere = Replace(EchapperCaractere, "#", "[#]")
    EchapperCaractere = Replace(EchapperCaractere, "[", "[[]")
End Function

Public Function ValiderLeNomDuFichier(ByVal Chaine As String) As String
    
    ValiderLeNomDuFichier = Chaine
    
    Dim i           As Long
    Dim Lookup      As String
    Dim ReplaceBy   As String
    
    Lookup = "����������������/\:*?><|"
    ReplaceBy = "aaaeeeeiiioouuuc________"
    
    If Len(ReplaceBy) < Len(Lookup) Then ReplaceBy = vbNullString
    
    For i = 1 To Len(Lookup)
        ' on remplace tous les caract�res de Lookup 1 par 1 dans iString
        ValiderLeNomDuFichier = Replace(ValiderLeNomDuFichier, Mid(Lookup, i, 1), (IIf(ReplaceBy = vbNullString, "", Mid(ReplaceBy, i, 1))), , , vbTextCompare)
    Next i
    
End Function
