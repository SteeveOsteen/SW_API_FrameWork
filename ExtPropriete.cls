VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExtPropriete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=======================================================================================
Option Explicit

Private pNom                As String
Private pValeur             As String
Private Const pNomClasse    As String = "ExtPropriete"

Public Property Let Nom(s As String)
    pNom = s
End Property

Public Property Get Nom() As String
    Nom = pNom
End Property

Public Property Let Valeur(s As String)
    pValeur = s
End Property

Public Property Get Valeur() As String
    Valeur = pValeur
End Property
