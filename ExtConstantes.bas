Attribute VB_Name = "ExtConstantes"
Public Const CONFIG_DEPLIEE As String = "SM-FLAT-PATTERN"
Public Const CONFIG_PLIEE As String = "#"
Public Const ARTICLE_LISTE_DES_PIECES_SOUDEES As String = "Article-liste-des-pièces-soudées"
Public Const FEUILLE_DE_BASE_LASER As String = "Base"
Public Const EPAISSEUR_DE_TOLE As String = "Epaisseur de la tôle"
Public Const NO_DOSSIER As String = "NoDossier"
Public Const NOM_ELEMENT As String = "Element"
Public Const CUBE_DE_VISUALISATION As String = "Cube de visualisation"
Public Const MODELE_DE_DESSIN_LASER As String = "MacroLaser" 'MacroLaser
Public Const NOM_CORPS_DEPLIEE As String = "Etat déplié"
Public Const ETAT_D_AFFICHAGE As String = "Etat d'affichage-"

Public Enum TypeFichier_e
    cAssemblage = 1
    cPiece = 2
    cDessin = 4
    cTousLesTypesDeFichier = 7
End Enum

Public Enum TypeCorps_e
    cTole = 1
    cProfil = 2
    cAutre = 4
    cTousLesTypesDeCorps = 7
End Enum

Public Enum TypeConfig_e
    cDeBase = 1
    cDerivee = 2
    cDepliee = 4
    cPliee = 8
    cToutesLesTypesDeConfig = 15
End Enum

Public Enum EtatFonction_e
    cDesactivee = 0
    cActivee = 1
End Enum

Public Enum Orientation_e
    cPortrait = 1
    cPaysage = 2
End Enum

Public Type Point
    X As Double
    Y As Double
    Z As Double
End Type

Public Type Dimensions
    Lg As Double
    Ht As Double
End Type

Public Type Rectangle
    MinX As Double
    MinY As Double
    MaxX As Double
    MaxY As Double
End Type
