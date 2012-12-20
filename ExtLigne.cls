VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExtLigne"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private pDepart As New ExtPoint
Private pArrivee As New ExtPoint
Private Const pNomClasse    As String = "ExtPropriete"

Public Property Let Depart(p As ExtPoint)
    Set pDepart = p
End Property

Public Property Get Depart() As ExtPoint
    Set Depart = pDepart
End Property

Public Property Let Arrivee(p As ExtPoint)
    Set pArrivee = p
End Property

Public Property Get Arrivee() As ExtPoint
    Set Arrivee = pArrivee
End Property

Public Property Get DeltaX() As Double
    DeltaX = pArrivee.X - pDepart.X
End Property

Public Property Get DeltaY() As Double
    DeltaY = pArrivee.Y - pDepart.Y
End Property

Public Property Get DeltaZ() As Double
    DeltaZ = pArrivee.Z - pDepart.Z
End Property

Public Property Get Longeur() As Double
    Longeur = Sqr(DeltaX ^ 2 + DeltaY ^ 2 + DeltaZ ^ 2)
End Property

Public Property Get AngleXYRadian() As Double
    If DeltaX <> 0# Then
        AngleXYRadian = Atn(DeltaY / DeltaX)
    Else
        If DeltaY = 0 Then
            AngleXYRadian = 0
        ElseIf DeltaY > 0 Then
            AngleXYRadian = 0.5 * Pi
        Else
            AngleXYRadian = -0.5 * Pi
        End If
    End If
    AngleXYRadian = ExtMath.AngleAbs(AngleXYRadian)
End Property

Public Property Get AngleXZRadian() As Double
    If DeltaX <> 0# Then
        AngleXZRadian = Atn(DeltaZ / DeltaX)
    Else
        If DeltaZ = 0 Then
            AngleXZRadian = 0
        ElseIf DeltaZ > 0 Then
            AngleXZRadian = 0.5 * Pi
        Else
            AngleXZRadian = -0.5 * Pi
        End If
    End If
    AngleXZRadian = ExtMath.AngleAbs(AngleXZRadian)
End Property

Public Property Get AngleZYRadian() As Double
    If DeltaZ <> 0# Then
        AngleZYRadian = Atn(DeltaY / DeltaZ)
    Else
        If DeltaY = 0 Then
            AngleZYRadian = 0
        ElseIf DeltaY > 0 Then
            AngleZYRadian = 0.5 * Pi
        Else
            AngleZYRadian = -0.5 * Pi
        End If
    End If
    AngleZYRadian = ExtMath.AngleAbs(AngleZYRadian)
End Property
