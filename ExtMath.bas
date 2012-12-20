Attribute VB_Name = "ExtMath"
Option Explicit

'Constantes mathématiques
Public Const Pi As Double = 3.14159265358979

'Renvoi le maxi de deux nombres
Public Function Max(ByVal a As Double, ByVal b As Double) As Double
    Max = b
    If a > b Then Max = a
End Function

'Renvoi le min de deux nombres
Public Function Min(ByVal a As Double, ByVal b As Double) As Double
    Min = b
    If a < b Then Min = a
End Function

'Converti les radian en degrés
Public Function Degree(ByVal Rad As Double) As Double
    Degree = Rad * 180# / Pi
End Function

'Converti les degrés en radians
Public Function Radian(ByVal Deg As Double) As Double
    Radian = Deg * Pi / 180#
End Function

'Renvoi l'angle en absolu, sans signe
'ex : rad = -0.5Pi -> 1.5Pi
Public Function AngleAbs(ByVal Rad As Double) As Double
    If Rad < 0 Then
        Rad = Angle0To2PI(Abs(Rad))
        AngleAbs = (2 * Pi) - Rad
    Else
        AngleAbs = Rad
    End If
End Function

'Renvoi l'angle compris entre 0 et 2Pi
'ex : rad = 3Pi -> 1Pi
Public Function Angle0To2PI(ByVal Rad As Double) As Double
    If Rad >= Pi Then
        Angle0To2PI = Rad - ((Rad \ Pi) * Pi)
    Else
        Angle0To2PI = Rad
    End If
End Function

'Arrondi par exces au nb de décimales indiqué
Public Function Arrondi(ByVal Nb As Double, ByVal NbDecimale As Integer) As Double
    Arrondi = Nb * (10 ^ NbDecimale)
    If Abs(Arrondi - Fix(Arrondi)) >= 0.5 Then
        Arrondi = Arrondi + 0.5
    End If
    Arrondi = Fix(Arrondi) / (10 ^ NbDecimale)
End Function
