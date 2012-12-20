VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExtPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private pX As Double
Private pY As Double
Private pZ As Double
Private Const pNomClasse    As String = "ExtPoint"

Private Sub Class_Initialize()
    pX = 0#
    pY = 0#
    pZ = 0#
End Sub

Public Property Let X(v As Double)
    pX = v
End Property

Public Property Get X() As Double
    X = pX
End Property

Public Property Let Y(v As Double)
    pY = v
End Property

Public Property Get Y() As Double
    Y = pY
End Property

Public Property Let Z(v As Double)
    pZ = v
End Property

Public Property Get Z() As Double
    Z = pZ
End Property
