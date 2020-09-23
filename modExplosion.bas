Attribute VB_Name = "mExplosion"
'Creates explosion sequence on screen.
'Utilizes  an explosion particle strip which is included with this project

Public Type objTrail        'trailing bits object definition
    LifeTime As Integer     'number of loops to show bit
    XPos As Single          'x-coordinate of bit
    YPos As Single          'y-coordinate of bit
End Type

Public Type objParticle         'explosion particle object definition
    LifeTime As Integer         'number of loops to show particle
    XMov As Single              'delta x distance particle has moved from centre
    YMov As Single              'delta y distance particle has moved from centre
    Angle As Single             'angle at which particle is moving away
    Speed As Single             'speed of particle movement
    Smoke(8) As objTrail   'array of trailing bits for the particle
    picno As Integer
End Type

Public Type objExplosion        'explosion object definition
    Particle(9) As objParticle 'array of main explosion particles
    Dot(9) As objParticle      'array of smaller dot-type explosion particles (no trailing bits)
    Counter As Long             'controls duration of explosion
    ExpCtrX As Long             'x-coordinate of centre point of the explosion
    ExpCtrY As Long             'y-coordinate of centre point of the explosion
    BMImgNo As Integer          'Image counter for explosion bitmap
    ImgCounter As Integer       'image counter for particle bitmap
    DoExplosion As Boolean      'explosion trigger
    ExpType As Integer          'what type of explosion to execute
End Type

Public Explode(8) As objExplosion   'you can modify array size - more than 3 or 4 on screen gets slow

Public Sub ShowExplosions()
Dim a As Integer, i As Integer, j As Integer
Dim XP As Long, YP As Long
    
For a = 0 To UBound(Explode)
    With Explode(a)
    
    If .DoExplosion = True Then
        If .ExpType = 1 Or .ExpType = 4 Then
            If .Counter < 192 Then
                .Counter = .Counter + 1
                If .ExpType = 1 Then
                    If .BMImgNo <= 14 Then
                        recExplode.Left = .BMImgNo * 88
                        recExplode.Right = recExplode.Left + 88
                        .BMImgNo = .BMImgNo + 1
                        backbuffer.BltFast .ExpCtrX - 44, .ExpCtrY - 60, ddsExplode, recExplode, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
                    End If
                ElseIf .ExpType = 4 Then
                    If .BMImgNo <= 25 Then
                        recExplode2.Left = .BMImgNo * 128
                        recExplode2.Right = recExplode2.Left + 128
                        .BMImgNo = .BMImgNo + 1
                        backbuffer.BltFast .ExpCtrX - 64, .ExpCtrY - 64, ddsExplode2, recExplode2, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
                    End If
                End If
                For i = 0 To 6
                    If .Particle(i).LifeTime > 0 Then
                        .Particle(i).LifeTime = .Particle(i).LifeTime - 1
                        .Particle(i).XMov = .Particle(i).XMov + Cos(.Particle(i).Angle) * .Particle(i).Speed
                        .Particle(i).YMov = .Particle(i).YMov + Sin(.Particle(i).Angle) * .Particle(i).Speed
                        XP = .ExpCtrX + .Particle(i).XMov
                        YP = .ExpCtrY + .Particle(i).YMov
                        recTrails.Left = Int((192 - .Particle(i).LifeTime) / 9.6) * 10
                        recTrails.Right = recTrails.Left + 10   'the next line must refer to your back buffer surface
                        backbuffer.BltFast XP - 5, YP - 5, ddsTrails, recTrails, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
                    End If
            
                    If .Dot(i).LifeTime > 0 Then
                        .Dot(i).LifeTime = .Dot(i).LifeTime - 1
                        .Dot(i).XMov = .Dot(i).XMov + Cos(.Dot(i).Angle) * .Dot(i).Speed
                        .Dot(i).YMov = .Dot(i).YMov + Sin(.Dot(i).Angle) * .Dot(i).Speed
                        XP = .ExpCtrX + .Dot(i).XMov
                        YP = .ExpCtrY + .Dot(i).YMov
                        recTrails.Left = Int((240 - .Dot(i).LifeTime) / 9.6) * 10
                        recTrails.Right = recTrails.Left + 10   'the next line must refer to your back buffer surface
                        backbuffer.BltFast XP - 5, YP - 5, ddsTrails, recTrails, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
                    End If
                Next i
                For i = 7 To 9
                    If .Particle(i).LifeTime > 0 Then
                        .Particle(i).LifeTime = .Particle(i).LifeTime - 1
                        .Particle(i).XMov = .Particle(i).XMov + Cos(.Particle(i).Angle) * .Particle(i).Speed
                        .Particle(i).YMov = .Particle(i).YMov + Sin(.Particle(i).Angle) * .Particle(i).Speed
                        XP = .ExpCtrX + .Particle(i).XMov
                        YP = .ExpCtrY + .Particle(i).YMov
                        recSmoke.Left = Int((192 - .Particle(i).LifeTime) / 18) * 8
                        recSmoke.Right = recSmoke.Left + 8   'the next line must refer to your back buffer surface
                        backbuffer.BltFast XP - 4, YP - 4, ddsSmoke, recSmoke, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
                        For j = 8 To 0 Step -1
                            If j = 0 Then
                                .Particle(i).Smoke(j).XPos = XP + .Particle(i).XMov
                                .Particle(i).Smoke(j).YPos = YP + .Particle(i).YMov
                                recSmoke.Left = Int((192 - .Particle(i).LifeTime) / 18) * 8
                            Else
                                .Particle(i).Smoke(j).XPos = .Particle(i).Smoke(j - 1).XPos
                                .Particle(i).Smoke(j).YPos = .Particle(i).Smoke(j - 1).YPos
                                recSmoke.Left = Int((192 - .Particle(i).LifeTime) / 18) * 8 + j * 8
                            End If
                            recSmoke.Right = recSmoke.Left + 5
                            backbuffer.BltFast .Particle(i).Smoke(j).XPos - 4, .Particle(i).Smoke(j).YPos - 4, ddsSmoke, recSmoke, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
                        Next j
                    End If
                Next i
            Else
                .DoExplosion = False
            End If
        ElseIf .ExpType = 2 Then
            If .Counter < 192 Then
                .Counter = .Counter + 1
                For i = 0 To UBound(.Particle)
                    If .Particle(i).LifeTime > 0 Then
                        .Particle(i).LifeTime = .Particle(i).LifeTime - 1
                        .Particle(i).XMov = .Particle(i).XMov + Cos(.Particle(i).Angle) * .Particle(i).Speed
                        .Particle(i).YMov = .Particle(i).YMov + Sin(.Particle(i).Angle) * .Particle(i).Speed
                        XP = .ExpCtrX + .Particle(i).XMov
                        YP = .ExpCtrY + .Particle(i).YMov
                        For j = 8 To 0 Step -1
                            If j = 0 Then
                                .Particle(i).Smoke(j).XPos = XP + .Particle(i).XMov
                                .Particle(i).Smoke(j).YPos = YP + .Particle(i).YMov
                                recSmoke.Left = Int((192 - .Particle(i).LifeTime) / 18) * 8
                            Else
                                .Particle(i).Smoke(j).XPos = .Particle(i).Smoke(j - 1).XPos
                                .Particle(i).Smoke(j).YPos = .Particle(i).Smoke(j - 1).YPos
                                recSmoke.Left = Int((192 - .Particle(i).LifeTime) / 18) * 8 + j * 8
                            End If
                            recSmoke.Right = recSmoke.Left + 5
                            backbuffer.BltFast .Particle(i).Smoke(j).XPos - 4, .Particle(i).Smoke(j).YPos - 4, ddsSmoke, recSmoke, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
                        Next j
                        recAstExplode.Left = .Particle(i).picno * 10
                        recAstExplode.Right = recAstExplode.Left + 10
                        backbuffer.BltFast XP - 5, YP - 5, ddsAstExplode, recAstExplode, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
                    End If
                Next i
            Else
                .DoExplosion = False
            End If
        ElseIf .ExpType = 3 Then
            If .BMImgNo <= 11 Then
                recSmExplode.Right = recSmExplode.Left + 32
                backbuffer.BltFast .ExpCtrX - 16, .ExpCtrY - 16, ddsSmExplode, recSmExplode, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
                .BMImgNo = .BMImgNo + 1
                recSmExplode.Left = .BMImgNo * 32
            Else
                .DoExplosion = False
            End If
        End If
    End If
    
    End With
Next a

End Sub

Public Sub StartExplosion(GetX As Single, GetY As Single, GetType As Integer)
Dim i As Integer, j As Integer, k As Integer

    recExplode.Left = 0
    recExplode.Right = recExplode.Left + 128
    
    For k = 0 To UBound(Explode)
        If Explode(k).DoExplosion = False Then
            With Explode(k)
                .ExpCtrX = GetX
                .ExpCtrY = GetY
                .Counter = 0
                .DoExplosion = True
                .ImgCounter = 0
                .BMImgNo = 0
                .ExpType = GetType
        
                For i = 0 To UBound(.Particle)
                    .Particle(i).LifeTime = Int(Rnd * 92 + 96)
                    .Particle(i).Angle = Rnd * 6.3
                    .Particle(i).Speed = Rnd * 5
                    .Particle(i).XMov = 0
                    .Particle(i).YMov = 0
                    .Particle(i).picno = Int(Rnd * 6)
                    For j = 0 To UBound(.Particle(0).Smoke)
                        .Particle(i).Smoke(j).LifeTime = Int(Rnd * 144 + 48)
                    Next j
            
                    .Dot(i).LifeTime = Int(Rnd * 92 + 100)
                    .Dot(i).Angle = Rnd * 6.3
                    .Dot(i).Speed = Rnd * 3
                    .Dot(i).XMov = 0
                    .Dot(i).YMov = 0
                Next i
            
            End With
            Exit Sub
        
        End If
    Next k

End Sub

Public Sub ResetExplosions()
Dim i As Integer

    For i = 0 To UBound(Explode)
        Explode(i).DoExplosion = False
    Next i

End Sub

