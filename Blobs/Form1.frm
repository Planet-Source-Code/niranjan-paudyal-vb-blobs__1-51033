VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8595
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   573
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   792
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Pt 
      AutoRedraw      =   -1  'True
      FillColor       =   &H00008000&
      ForeColor       =   &H00008000&
      Height          =   6975
      Left            =   0
      ScaleHeight     =   461
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1013
      TabIndex        =   0
      Top             =   0
      Width           =   15255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Running As Boolean
Dim t(5) As Particle
Dim Last As pointapi

Dim P(3) As pointapi
Dim I
Dim p1 As pointapi, p2 As pointapi
Const pi = 3.14
Sub run()
    
    While Running
        Pt.Cls
        MoveParticles t(0), Last.x, Last.y
        
        For I = 0 To UBound(t)
            If I <> 0 Then MoveParticles t(I), t(I - 1).Position.x, t(I - 1).Position.y
        
            t(I).Velocity.y = t(I).Velocity.y / 1.2
            t(I).Velocity.x = t(I).Velocity.x / 1.2
            
            t(I).Position.x = t(I).Position.x + t(I).Velocity.x
            t(I).Position.y = t(I).Position.y + t(I).Velocity.y '+ T(i).Mass * 3
            ApplyForce t(I), 0, t(I).Mass * 9.8, 0.1

            
            If t(I).Position.x - t(I).Radius < 0 Then
                t(I).Velocity.x = -t(I).Velocity.x
                t(I).Position.x = t(I).Radius
            End If
            If t(I).Position.x + t(I).Radius > Pt.ScaleWidth Then
                t(I).Velocity.x = -t(I).Velocity.x
                t(I).Position.x = Pt.ScaleWidth - t(I).Radius
            End If
            If t(I).Position.y - t(I).Radius < 0 Then
                t(I).Velocity.y = -t(I).Velocity.y
                t(I).Position.y = t(I).Radius
            End If
            If t(I).Position.y + t(I).Radius > Pt.ScaleHeight Then
                t(I).Velocity.y = -t(I).Velocity.y
                t(I).Position.y = Pt.ScaleHeight - t(I).Radius
            End If
            
            'If (T(i).Velocity.X * T(i).Velocity.X + T(i).Velocity.Y * T(i).Velocity.Y) <= 1 Then
            '    T(i).Velocity.X = 0
            '    T(i).Velocity.Y = -T(i).Mass * 3
            'End If
            
            
            Pt.Circle (t(I).Position.x, t(I).Position.y), t(I).Radius
            
            
            
        Next I
        
        For I = 0 To UBound(t) - 1
            
            DrawBlob t(I).Position, t(I + 1).Position, t(I).Radius, 1
            
        Next I
        
            
        DoEvents
        'Exit Sub
    Wend
    
    
End Sub


Private Sub DrawBlob(X1 As pointSng, X2 As pointSng, Radius As Single, Elasticty As Single)
    
    d = Sqr((X1.x - X2.x) ^ 2 + (X1.y - X2.y) ^ 2)
    If d = 0 Then d = 0.000001
    'If d = 0 Then seperation = Elasticty * Radius * 2 / 1 Else seperation = Elasticty * Radius * 2 / d
    
 
    
    If (X1.x - X2.x) = 0 Then ang = pi / 2 Else ang = Atn((X1.y - X2.y) / (X1.x - X2.x))
    'Me.Caption = ang
    
    v = pi / 2
    
    s1x = Cos(ang - v) * Radius + X1.x
    s1y = Sin(ang - v) * Radius + X1.y
    e1x = Cos(ang - v) * Radius + X2.x
    e1y = Sin(ang - v) * Radius + X2.y
    
 
    s2x = Cos(ang + v) * Radius + X1.x
    s2y = Sin(ang + v) * Radius + X1.y
    e2x = Cos(ang + v) * Radius + X2.x
    e2y = Sin(ang + v) * Radius + X2.y
    
    h = 2 * Radius * Radius / (d) '- 2
    
    
    'Pt.Line (e1x, s1y)-(e1x, e1y)
    
    
    
    If h > Radius Then h = Radius
    'Me.Caption = h
    
    dx = Cos(ang - v) * h
    dy = Sin(ang - v) * h
    halfdx = dx + X1.x + (X2.x - X1.x) / 2
    halfdy = dy + X1.y + (X2.y - X1.y) / 2
    
    
    
    
    P(0).x = s1x: P(0).y = s1y
    P(1).x = halfdx: P(1).y = halfdy
    P(2).x = halfdx: P(2).y = halfdy
    P(3).x = e1x: P(3).y = e1y
    Module1.PolyBezier Pt.hdc, P(0), 4
    
    
    
    
    halfdx = -dx + X1.x + (X2.x - X1.x) / 2
    halfdy = -dy + X1.y + (X2.y - X1.y) / 2
    
    P(0).x = s2x: P(0).y = s2y
    P(1).x = halfdx: P(1).y = halfdy
    P(2).x = halfdx: P(2).y = halfdy
    P(3).x = e2x: P(3).y = e2y
    Module1.PolyBezier Pt.hdc, P(0), 4
    

    'Pt.PSet (halfdx, halfdy), vbRed
    'FloodFill Pt.hdc, halfdx, halfdy, Pt.ForeColor 'Pt.BackColor
    

    'extFloodFill Pt.hdc, 0, 0, Pt.BackColor, 0
    'Pt.Circle (X1.X, X1.Y), 10
    'Pt.Refresh
End Sub


Private Sub MoveParticles(Parti As Particle, x, y)
            
        If (x - Parti.Position.x) = 0 Then ang = 0.0000001 Else ang = Atn((Parti.Position.y - y) / (Parti.Position.x - x))
    
        d = Sqr((Parti.Position.x - x) ^ 2 + (Parti.Position.y - y) ^ 2)
        ext = d - Parti.NaturalLength
        
        energy = (Parti.Modulus * ext * ext) / (2 * Parti.NaturalLength)
        speed = Sqr(2 * energy / Parti.Mass) / 10
        
        If x > Parti.Position.x Then
            R = speed * Cos(ang)
            v = speed * Sin(ang)
        Else
            R = -speed * Cos(ang)
            v = -speed * Sin(ang)
        End If
        
        Parti.Velocity.x = Parti.Velocity.x + R '-speed * Cos(ang)
        Parti.Velocity.y = Parti.Velocity.y + v 'speed * Sin(ang)
End Sub


Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
    For I = 0 To UBound(t)
        
        t(I).Position.y = Pt.ScaleHeight / 2
        t(I).NaturalLength = 0.1
        t(I).Extension = 0
        t(I).Modulus = 0.5
        t(I).Mass = I * I * 4 + 0.7
        t(I).Radius = 15
        t(I).Position.x = (UBound(t) * t(I).Radius * 2) - I * t(I).Radius * 2 ' Pt.ScaleWidth / 2
    Next I
    
    Last.x = Pt.ScaleWidth / 2
    Last.y = t(0).Position.y
    Me.Show
    Running = True
    run
End Sub

Private Sub Form_Resize()
    Pt.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Running = False
End Sub

Private Sub Pt_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Last.x = x
    Last.y = y
    
End Sub

