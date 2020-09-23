Attribute VB_Name = "Module1"
Type pointapi
    x As Long
    y As Long
End Type

Type pointSng
    x As Single
    y As Single
End Type

Type Particle
    Position As pointSng
    Velocity As pointSng
    NaturalLength As Single
    Extension As Single
    Mass As Single
    Modulus As Single
    Radius As Single
End Type

Declare Function PolyBezier Lib "gdi32" (ByVal hdc As Long, lppt As pointapi, ByVal cPoints As Long) As Long
Declare Function FloodFill Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Declare Function ExtFloodFill Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long

Sub ApplyForce(P As Particle, ForceX As Single, ForceY As Single, Time_Of_Force As Single)
    P.Velocity.x = P.Velocity.x + (ForceX * Time_Of_Force / P.Mass)
    P.Velocity.y = P.Velocity.y + (ForceY * Time_Of_Force / P.Mass)
End Sub

